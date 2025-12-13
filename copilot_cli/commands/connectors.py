"""Connector commands for listing available Power Platform connectors."""
import typer
import json
import yaml
from pathlib import Path
from typing import Optional

from ..client import get_client
from ..output import print_json, print_table, handle_api_error, print_success


app = typer.Typer(help="List and inspect Power Platform connectors")


def is_custom_connector(connector: dict) -> bool:
    """
    Determine if a connector is custom or managed.

    Custom connectors have different properties structure than managed connectors.
    """
    props = connector.get("properties", {})

    # Explicit isCustomApi flag from Power Apps API
    if props.get("isCustomApi", False):
        return True

    # Check for custom connector indicators (Dataverse-sourced connectors)
    if "environment" in props:
        return True

    # Check publisher - custom connectors often have user/org as publisher
    publisher = (props.get("publisher") or "").lower()
    if publisher and publisher not in ["microsoft", "microsoft corporation", "azure"]:
        # Could be custom, but also third-party managed
        # Check tier - custom connectors typically don't have tier or have "NotSpecified"
        tier = props.get("tier", "")
        if not tier or tier == "NotSpecified":
            return True

    return False


def format_connector_for_display(connector: dict) -> dict:
    """Format a connector for display."""
    props = connector.get("properties", {})
    dataverse = connector.get("_dataverse", {})

    description = props.get("description") or ""
    if len(description) > 60:
        description = description[:57] + "..."

    is_custom = is_custom_connector(connector)

    return {
        "name": props.get("displayName") or connector.get("name", ""),
        "id": connector.get("name", ""),  # connectorinternalid (e.g., shared_cr83c-5fasana-...)
        "logical_name": dataverse.get("name", ""),  # Dataverse logical name (e.g., cr83c_5fasana)
        "type": "Custom" if is_custom else "Managed",
        "publisher": props.get("publisher") or "",
        "tier": props.get("tier") or "N/A",
        "description": description,
        "source": connector.get("_source", ""),
    }


def generate_api_properties(openapi_def: dict, icon_brand_color: str = "#007ee5") -> dict:
    """
    Generate apiProperties.json content from OpenAPI definition.

    Extracts authentication configuration from securityDefinitions and creates
    the properties structure needed by Power Platform custom connectors.

    Args:
        openapi_def: OpenAPI 2.0 definition
        icon_brand_color: Icon brand color in hex format

    Returns:
        dict: API properties structure
    """
    properties = {
        "iconBrandColor": icon_brand_color,
        "capabilities": [],
        "policyTemplateInstances": []
    }

    # Extract connection parameters from securityDefinitions
    security_defs = openapi_def.get("securityDefinitions", {})
    connection_params = {}

    for sec_name, sec_def in security_defs.items():
        sec_type = sec_def.get("type", "")

        if sec_type == "oauth2":
            # OAuth2 configuration
            connection_params["token"] = {
                "type": "oauthSetting",
                "oAuthSettings": {
                    "identityProvider": "oauth2",
                    "clientId": "",  # Must be set via UI
                    "scopes": list(sec_def.get("scopes", {}).keys()),
                    "redirectMode": "Global",
                    "redirectUrl": "https://global.consent.azure-apim.net/redirect",
                    "properties": {
                        "IsFirstParty": "False"
                    },
                    "customParameters": {
                        "authorizationUrl": {
                            "value": sec_def.get("authorizationUrl", "")
                        },
                        "tokenUrl": {
                            "value": sec_def.get("tokenUrl", "")
                        },
                        "refreshUrl": {
                            "value": sec_def.get("tokenUrl", "")  # Often same as tokenUrl
                        }
                    }
                }
            }
        elif sec_type == "apiKey":
            # API Key authentication
            param_name = sec_def.get("name", "api_key")
            in_location = sec_def.get("in", "header")

            connection_params[param_name] = {
                "type": "securestring",
                "uiDefinition": {
                    "displayName": f"API Key ({param_name})",
                    "description": f"The API Key for authentication (sent in {in_location})",
                    "tooltip": "Provide your API Key",
                    "constraints": {
                        "required": "true",
                        "clearText": False
                    }
                }
            }
        elif sec_type == "basic":
            # Basic authentication
            connection_params["username"] = {
                "type": "string",
                "uiDefinition": {
                    "displayName": "Username",
                    "description": "Username for basic authentication",
                    "tooltip": "Provide your username",
                    "constraints": {
                        "required": "true"
                    }
                }
            }
            connection_params["password"] = {
                "type": "securestring",
                "uiDefinition": {
                    "displayName": "Password",
                    "description": "Password for basic authentication",
                    "tooltip": "Provide your password",
                    "constraints": {
                        "required": "true",
                        "clearText": False
                    }
                }
            }

    if connection_params:
        properties["connectionParameters"] = connection_params

    return {"properties": properties}


def validate_openapi_definition(openapi_def: dict) -> tuple[bool, str]:
    """
    Validate that the OpenAPI definition is in the correct format.

    Args:
        openapi_def: Parsed OpenAPI definition dict

    Returns:
        tuple: (is_valid, error_message)
    """
    # Check for OpenAPI 2.0 (Swagger)
    swagger_version = openapi_def.get("swagger")
    if not swagger_version or not swagger_version.startswith("2."):
        openapi_version = openapi_def.get("openapi", "")
        if openapi_version.startswith("3."):
            return False, (
                "OpenAPI 3.0 is not supported by Power Platform. "
                "Please convert to OpenAPI 2.0 (Swagger) format. "
                "You can use tools like swagger-cli or API Transformer for conversion."
            )
        return False, "OpenAPI definition must be in OpenAPI 2.0 (Swagger) format."

    # Check required fields
    required_fields = ["swagger", "info", "host", "basePath", "schemes"]
    missing = [f for f in required_fields if f not in openapi_def]
    if missing:
        return False, f"Missing required fields: {', '.join(missing)}"

    # Check info section
    info = openapi_def.get("info", {})
    if not info.get("title"):
        return False, "info.title is required"
    if not info.get("version"):
        return False, "info.version is required"

    # Check size (must be < 1MB)
    json_str = json.dumps(openapi_def)
    size_mb = len(json_str.encode('utf-8')) / (1024 * 1024)
    if size_mb >= 1.0:
        return False, f"OpenAPI definition is too large ({size_mb:.2f} MB). Must be less than 1 MB."

    return True, ""


@app.command("list")
def connectors_list(
    custom: bool = typer.Option(
        False,
        "--custom",
        "-c",
        help="Show only custom connectors",
    ),
    managed: bool = typer.Option(
        False,
        "--managed",
        "-m",
        help="Show only managed (Microsoft) connectors",
    ),
    filter_text: Optional[str] = typer.Option(
        None,
        "--filter",
        "-f",
        help="Filter by name or publisher (case-insensitive)",
    ),
    table: bool = typer.Option(
        False,
        "--table",
        "-t",
        help="Display output as a formatted table instead of JSON",
    ),
    raw: bool = typer.Option(
        False,
        "--raw",
        "-r",
        help="Output raw JSON including all metadata",
    ),
):
    """
    List all available connectors in the environment.

    This command queries TWO APIs to get complete connector coverage:
      - Dataverse connector table: ALL custom connectors (including MCP connectors)
      - Power Apps API: Managed (Microsoft) connectors

    Connectors are proxies/wrappers around APIs that define what actions
    are available (e.g., Asana, SharePoint, SQL Server). They represent
    the "type" of service you can connect to.

    Connector Types:
      - Custom: User-created connectors in the environment
      - Managed: Built-in connectors published by Microsoft

    Examples:
        copilot connectors list --table
        copilot connectors list --custom --table
        copilot connectors list --managed --table
        copilot connectors list --filter "asana" --table
        copilot connectors list --raw
    """
    if custom and managed:
        typer.echo("Error: Cannot specify both --custom and --managed", err=True)
        raise typer.Exit(1)

    try:
        client = get_client()
        connectors = client.list_connectors(
            custom_only=custom,
            managed_only=managed,
        )

        if not connectors:
            if custom:
                typer.echo("No custom connectors found in this environment.")
            elif managed:
                typer.echo("No managed connectors found in this environment.")
            else:
                typer.echo("No connectors found.")
            return

        # Filter by text
        if filter_text:
            filter_lower = filter_text.lower()
            connectors = [
                c for c in connectors
                if filter_lower in c.get("properties", {}).get("displayName", "").lower()
                or filter_lower in c.get("properties", {}).get("publisher", "").lower()
                or filter_lower in c.get("name", "").lower()
                or filter_lower in c.get("_dataverse", {}).get("name", "").lower()
            ]

        if not connectors:
            typer.echo("No connectors match the filter criteria.")
            return

        # Raw output includes all metadata
        if raw:
            print_json(connectors)
            return

        formatted = [format_connector_for_display(c) for c in connectors]

        # Sort by type (Custom first) then name
        formatted.sort(key=lambda x: (0 if x["type"] == "Custom" else 1, x["name"].lower()))

        if table:
            print_table(
                formatted,
                columns=["name", "type", "publisher", "tier", "source", "id"],
                headers=["Name", "Type", "Publisher", "Tier", "Source", "ID"],
            )
        else:
            print_json(formatted)
    except Exception as e:
        exit_code = handle_api_error(e)
        raise typer.Exit(exit_code)


def extract_operations(
    connector: dict,
    include_deprecated: bool = False,
    include_internal: bool = False,
) -> list:
    """
    Extract operations (actions/triggers) from connector swagger definition.

    Args:
        connector: The connector definition with swagger
        include_deprecated: If True, include deprecated operations
        include_internal: If True, include internal-visibility operations
                         (internal ops cannot be used as Copilot agent tools)

    Returns:
        List of operation dicts with id, name, description, type, deprecated, visibility
    """
    operations = []
    swagger = connector.get("properties", {}).get("swagger", {})
    paths = swagger.get("paths", {})

    for path, methods in paths.items():
        for method, details in methods.items():
            if method in ["get", "post", "put", "patch", "delete"]:
                op_id = details.get("operationId")
                if not op_id:
                    continue

                is_deprecated = details.get("deprecated", False)
                visibility = details.get("x-ms-visibility", "normal")
                is_internal = visibility == "internal"

                # Skip deprecated unless explicitly requested
                if is_deprecated and not include_deprecated:
                    continue

                # Skip internal unless explicitly requested
                # Internal operations cannot be used as Copilot agent tools
                if is_internal and not include_internal:
                    continue

                # Determine if trigger or action
                is_trigger = details.get("x-ms-trigger") is not None
                op_type = "Trigger" if is_trigger else "Action"

                # Get description, truncate if too long
                description = details.get("description") or details.get("summary") or ""
                if len(description) > 80:
                    description = description[:77] + "..."

                operations.append({
                    "id": op_id,
                    "name": details.get("summary") or op_id,
                    "type": op_type,
                    "method": method.upper(),
                    "deprecated": is_deprecated,
                    "visibility": visibility,
                    "description": description,
                })

    # Sort: Actions first, then Triggers, then by name
    operations.sort(key=lambda x: (0 if x["type"] == "Action" else 1, x["id"].lower()))

    return operations


@app.command("get")
def connectors_get(
    connector_id: str = typer.Argument(
        ...,
        help="The connector's unique identifier (e.g., shared_asana, shared_office365)",
    ),
    table: bool = typer.Option(
        False,
        "--table",
        "-t",
        help="Display operations as a formatted table",
    ),
    include_deprecated: bool = typer.Option(
        False,
        "--include-deprecated",
        "-d",
        help="Include deprecated operations (hidden by default)",
    ),
    include_internal: bool = typer.Option(
        False,
        "--include-internal",
        "-i",
        help="Include internal operations (cannot be used as agent tools)",
    ),
    raw: bool = typer.Option(
        False,
        "--raw",
        "-r",
        help="Output raw JSON connector definition (ignores --table)",
    ),
    openapi: bool = typer.Option(
        False,
        "--openapi",
        "--swagger",
        help="Output the full OpenAPI/Swagger definition (JSON format)",
    ),
    output_file: Optional[str] = typer.Option(
        None,
        "--output",
        "-o",
        help="Write OpenAPI definition to file (use with --openapi)",
    ),
):
    """
    Get details for a specific connector including available operations.

    By default, shows only operations that can be used as Copilot agent tools.
    Deprecated and internal-visibility operations are hidden by default.

    Examples:
        copilot connectors get shared_asana --table
        copilot connectors get shared_asana --table --include-deprecated
        copilot connectors get shared_asana --table --include-internal
        copilot connectors get shared_office365 --raw
        copilot connectors get shared_asana --openapi
        copilot connectors get shared_asana --openapi --output ./asana-spec.json
    """
    try:
        client = get_client()
        connector = client.get_connector(connector_id)

        # OpenAPI/Swagger output
        if openapi:
            swagger = connector.get("properties", {}).get("swagger", {})
            if not swagger:
                typer.echo("Error: No OpenAPI/Swagger definition found for this connector.", err=True)
                raise typer.Exit(1)

            if output_file:
                # Write to file
                output_path = Path(output_file)
                try:
                    output_path.write_text(json.dumps(swagger, indent=2))
                    props = connector.get("properties", {})
                    typer.echo(f"OpenAPI definition for '{props.get('displayName', connector_id)}' written to: {output_file}")
                except Exception as e:
                    typer.echo(f"Error writing to file: {e}", err=True)
                    raise typer.Exit(1)
            else:
                # Output to stdout
                print_json(swagger)
            return

        # Raw output - full JSON
        if raw:
            print_json(connector)
            return

        # Extract and display operations
        operations = extract_operations(connector, include_deprecated, include_internal)
        props = connector.get("properties", {})

        # Show connector summary
        typer.echo(f"\nConnector: {props.get('displayName', connector_id)}")
        typer.echo(f"ID: {connector.get('name', connector_id)}")
        typer.echo(f"Publisher: {props.get('publisher', 'N/A')}")

        if not operations:
            typer.echo("\nNo usable operations found.")
            typer.echo("Use --include-deprecated and/or --include-internal to see hidden operations.")
            return

        # Count hidden operations
        all_ops = extract_operations(connector, True, True)
        deprecated_count = len([o for o in all_ops if o["deprecated"]])
        internal_count = len([o for o in all_ops if o["visibility"] == "internal"])

        hidden_parts = []
        if deprecated_count > 0 and not include_deprecated:
            hidden_parts.append(f"{deprecated_count} deprecated")
        if internal_count > 0 and not include_internal:
            hidden_parts.append(f"{internal_count} internal")

        hidden_msg = f" ({', '.join(hidden_parts)} hidden)" if hidden_parts else ""
        typer.echo(f"\nOperations: {len(operations)}{hidden_msg}")

        if table:
            # Table format
            display_ops = []
            for op in operations:
                row = {
                    "id": op["id"],
                    "name": op["name"][:40] + "..." if len(op["name"]) > 40 else op["name"],
                    "type": op["type"],
                    "method": op["method"],
                }
                if include_deprecated:
                    row["deprecated"] = "Yes" if op["deprecated"] else "No"
                if include_internal:
                    row["visibility"] = op["visibility"]
                display_ops.append(row)

            columns = ["id", "name", "type", "method"]
            headers = ["Operation ID", "Name", "Type", "Method"]
            if include_deprecated:
                columns.append("deprecated")
                headers.append("Deprecated")
            if include_internal:
                columns.append("visibility")
                headers.append("Visibility")

            print_table(display_ops, columns=columns, headers=headers)
        else:
            # JSON format - just operations list
            print_json(operations)

    except Exception as e:
        exit_code = handle_api_error(e)
        raise typer.Exit(exit_code)


@app.command("create")
def connectors_create(
    name: str = typer.Option(..., "--name", "-n", help="Display name for the connector"),
    swagger_file: str = typer.Option(..., "--swagger-file", "-f", help="Path to OpenAPI 2.0 (Swagger) definition file"),
    description: Optional[str] = typer.Option(None, "--description", "-d", help="Connector description"),
    icon_brand_color: Optional[str] = typer.Option("#007ee5", "--icon-brand-color", help="Icon brand color (hex format)"),
    environment: Optional[str] = typer.Option(None, "--environment", "--env", help="Environment ID (defaults to configured environment)"),
    oauth_client_id: Optional[str] = typer.Option(None, "--oauth-client-id", help="OAuth 2.0 Client ID (required for OAuth connectors)"),
    oauth_client_secret: Optional[str] = typer.Option(None, "--oauth-client-secret", help="OAuth 2.0 Client Secret (required for OAuth connectors)"),
    oauth_redirect_url: Optional[str] = typer.Option(None, "--oauth-redirect-url", help="Custom OAuth redirect URL (overrides default Power Platform redirect URL)"),
    script: Optional[str] = typer.Option(None, "--script", "-x", help="Path to C# script file (.csx) for custom code transformations"),
    script_operations: Optional[str] = typer.Option(None, "--script-operations", help="Comma-separated list of operationIds that use the script (defaults to all operations)"),
):
    """
    Create a new custom connector from an OpenAPI 2.0 (Swagger) definition.

    The OpenAPI definition must be in OpenAPI 2.0 format (not 3.0).

    For OAuth connectors, provide --oauth-client-id and --oauth-client-secret.
    Without these, the connector will be created but connections cannot authenticate.

    For custom code (request/response transformations), use --script to specify
    a C# script file (.csx). The script can modify requests before they're sent
    to the API and responses before they're returned.

    After creating the connector, use 'copilot connections create' to create a
    connection and authenticate.

    Examples:
      # Basic connector (non-OAuth)
      copilot connectors create --name "My API" --swagger-file ./api.json

      # OAuth connector with credentials
      copilot connectors create --name "My API" --swagger-file ./api.json \\
        --oauth-client-id "client123" --oauth-client-secret "secret456"

      # Connector with custom code script
      copilot connectors create --name "My API" --swagger-file ./api.json \\
        --script ./code.csx --oauth-client-id "client123" --oauth-client-secret "secret456"

      # Script for specific operations only
      copilot connectors create --name "My API" --swagger-file ./api.json \\
        --script ./code.csx --script-operations "CreateTask,UpdateTask"

      # Then create a connection
      copilot connections create --connector-id <connector-id> --name "My Connection" --oauth
    """
    try:
        # Read and parse OpenAPI file
        swagger_path = Path(swagger_file)
        if not swagger_path.exists():
            typer.echo(f"Error: File not found: {swagger_file}", err=True)
            raise typer.Exit(1)

        try:
            file_content = swagger_path.read_text()

            # Try JSON first, then YAML
            try:
                openapi_def = json.loads(file_content)
            except json.JSONDecodeError:
                try:
                    openapi_def = yaml.safe_load(file_content)
                except yaml.YAMLError as yaml_err:
                    typer.echo(f"Error: Invalid JSON/YAML format: {yaml_err}", err=True)
                    raise typer.Exit(1)
        except Exception as e:
            typer.echo(f"Error reading file: {e}", err=True)
            raise typer.Exit(1)

        # Validate OpenAPI definition
        is_valid, error_msg = validate_openapi_definition(openapi_def)
        if not is_valid:
            typer.echo(f"Error: {error_msg}", err=True)
            raise typer.Exit(1)

        # Check if connector uses OAuth and validate credentials
        security_defs = openapi_def.get("securityDefinitions", {})
        uses_oauth = any(sec.get("type") == "oauth2" for sec in security_defs.values())

        if uses_oauth:
            if not oauth_client_id or not oauth_client_secret:
                typer.echo("Error: This connector uses OAuth 2.0 authentication.", err=True)
                typer.echo("You must provide --oauth-client-id and --oauth-client-secret", err=True)
                typer.echo()
                typer.echo("Without OAuth credentials, the connector will be created but")
                typer.echo("connections cannot authenticate until credentials are added via UI.")
                typer.echo()
                typer.echo("Continue anyway? (y/N): ", nl=False)

                response = input().strip().lower()
                if response != 'y':
                    typer.echo("Cancelled.")
                    raise typer.Exit(0)

        # Validate script file if provided
        script_file = None
        if script:
            script_path = Path(script)
            if not script_path.exists():
                typer.echo(f"Error: Script file not found: {script}", err=True)
                raise typer.Exit(1)
            if not script_path.suffix.lower() in ['.csx', '.cs']:
                typer.echo(f"Warning: Script file should be a C# script (.csx or .cs): {script}", err=True)
            script_file = str(script_path.resolve())

        # Parse script operations if provided
        ops_list = None
        if script_operations:
            ops_list = [op.strip() for op in script_operations.split(',') if op.strip()]
            if not ops_list:
                typer.echo("Error: --script-operations requires at least one operation ID", err=True)
                raise typer.Exit(1)

        # Create connector
        client = get_client()
        result = client.create_custom_connector(
            name=name,
            openapi_definition=openapi_def,
            description=description,
            icon_brand_color=icon_brand_color or "#007ee5",
            environment_id=environment,
            oauth_client_id=oauth_client_id,
            oauth_client_secret=oauth_client_secret,
            oauth_redirect_url=oauth_redirect_url,
            script_file=script_file,
            script_operations=ops_list,
        )

        connector_id = result["connector_id"]
        environment_id = result["environment_id"]

        print_success(f"Custom connector '{name}' created successfully!")
        typer.echo(f"Connector ID: {connector_id}")
        typer.echo(f"Environment: {environment_id}")
        if script_file:
            typer.echo(f"Custom Code: Enabled ({Path(script_file).name})")
        typer.echo()

        # Show next steps with redirect URL info for OAuth connectors
        typer.echo("⚠️  Next Steps:")

        if uses_oauth:
            # Note: Power Platform strips the "shared_" prefix from connector_id for OAuth redirect URL
            redirect_connector_id = connector_id.replace("shared_", "", 1) if connector_id.startswith("shared_") else connector_id
            typer.echo()
            typer.echo("1. Register this redirect URL in your OAuth app settings:")
            typer.echo(f"   https://global.consent.azure-apim.net/redirect/{redirect_connector_id}")
            typer.echo()
            typer.echo("   💡 Or use wildcard (if supported): https://global.consent.azure-apim.net/redirect/*")
            typer.echo()
            typer.echo("2. Create a connection:")
            typer.echo(f"   copilot connections create --connector-id {connector_id} --name \"My Connection\" --oauth")
        else:
            typer.echo(f"1. Test the connector: copilot connectors get {connector_id}")
            typer.echo(f"2. Create a connection: copilot connections create --connector-id {connector_id} --name \"My Connection\"")

    except Exception as e:
        exit_code = handle_api_error(e)
        raise typer.Exit(exit_code)


@app.command("update")
def connectors_update(
    connector_id: str = typer.Argument(..., help="The connector's unique identifier (e.g., shared_cr83c-5fasana-...)"),
    swagger_file: Optional[str] = typer.Option(None, "--swagger-file", "-f", help="Path to OpenAPI 2.0 (Swagger) definition file"),
    description: Optional[str] = typer.Option(None, "--description", "-d", help="Connector description"),
    icon_brand_color: Optional[str] = typer.Option(None, "--icon-brand-color", help="Icon brand color (hex format)"),
    environment: Optional[str] = typer.Option(None, "--environment", "--env", help="Environment ID (defaults to configured environment)"),
    oauth_client_id: Optional[str] = typer.Option(None, "--oauth-client-id", help="OAuth 2.0 Client ID"),
    oauth_client_secret: Optional[str] = typer.Option(None, "--oauth-client-secret", help="OAuth 2.0 Client Secret"),
    oauth_redirect_url: Optional[str] = typer.Option(None, "--oauth-redirect-url", help="Custom OAuth redirect URL"),
    script: Optional[str] = typer.Option(None, "--script", "-x", help="Path to C# script file (.csx) for custom code transformations"),
    script_operations: Optional[str] = typer.Option(None, "--script-operations", help="Comma-separated list of operationIds that use the script"),
):
    """
    Update an existing custom connector.

    You can update the OpenAPI definition, description, custom code script, or
    authentication settings. Only provided options will be updated; other settings
    are preserved.

    To add or update custom code, use --script to specify a C# script file (.csx).
    The script can modify requests before they're sent to the API and responses
    before they're returned.

    Examples:
      # Update OpenAPI definition
      copilot connectors update shared_myapi-... --swagger-file ./api-v2.json

      # Add custom code script to existing connector
      copilot connectors update shared_myapi-... --script ./code.csx

      # Update script for specific operations only
      copilot connectors update shared_myapi-... --script ./code.csx --script-operations "CreateTask,UpdateTask"

      # Update description only
      copilot connectors update shared_myapi-... --description "Updated API connector"
    """
    try:
        # Parse OpenAPI file if provided
        openapi_def = None
        if swagger_file:
            swagger_path = Path(swagger_file)
            if not swagger_path.exists():
                typer.echo(f"Error: File not found: {swagger_file}", err=True)
                raise typer.Exit(1)

            try:
                file_content = swagger_path.read_text()

                # Try JSON first, then YAML
                try:
                    openapi_def = json.loads(file_content)
                except json.JSONDecodeError:
                    try:
                        openapi_def = yaml.safe_load(file_content)
                    except yaml.YAMLError as yaml_err:
                        typer.echo(f"Error: Invalid JSON/YAML format: {yaml_err}", err=True)
                        raise typer.Exit(1)
            except Exception as e:
                typer.echo(f"Error reading file: {e}", err=True)
                raise typer.Exit(1)

            # Validate OpenAPI definition
            is_valid, error_msg = validate_openapi_definition(openapi_def)
            if not is_valid:
                typer.echo(f"Error: {error_msg}", err=True)
                raise typer.Exit(1)

        # Validate script file if provided
        script_file = None
        if script:
            script_path = Path(script)
            if not script_path.exists():
                typer.echo(f"Error: Script file not found: {script}", err=True)
                raise typer.Exit(1)
            if not script_path.suffix.lower() in ['.csx', '.cs']:
                typer.echo(f"Warning: Script file should be a C# script (.csx or .cs): {script}", err=True)
            script_file = str(script_path.resolve())

        # Parse script operations if provided
        ops_list = None
        if script_operations:
            ops_list = [op.strip() for op in script_operations.split(',') if op.strip()]
            if not ops_list:
                typer.echo("Error: --script-operations requires at least one operation ID", err=True)
                raise typer.Exit(1)

        # Check if any update options provided
        if not any([swagger_file, description, icon_brand_color, script, oauth_client_id, oauth_client_secret]):
            typer.echo("Error: No update options provided. Use --help to see available options.", err=True)
            raise typer.Exit(1)

        # Update connector
        client = get_client()
        result = client.update_custom_connector(
            connector_id=connector_id,
            openapi_definition=openapi_def,
            description=description,
            icon_brand_color=icon_brand_color,
            environment_id=environment,
            oauth_client_id=oauth_client_id,
            oauth_client_secret=oauth_client_secret,
            oauth_redirect_url=oauth_redirect_url,
            script_file=script_file,
            script_operations=ops_list,
        )

        typer.echo(f"Display Name: {result.get('display_name', 'N/A')}")
        print_success(f"Connector '{connector_id}' updated successfully!")
        if result.get("script_uploaded"):
            typer.echo(f"Custom Code: Updated ({Path(script_file).name})")

        # Warn about connection refresh when swagger is updated
        if swagger_file:
            typer.echo("")
            typer.secho(
                "⚠️  WARNING: Existing connections may cache the old schema.",
                fg=typer.colors.YELLOW,
                bold=True
            )
            typer.echo("   To use new/modified operations, you may need to:")
            typer.echo("   1. Delete and recreate connections for this connector")
            typer.echo("   2. Update connection references with the new connection ID")
            typer.echo("")
            typer.echo("   Commands:")
            typer.echo(f"   copilot connections list --connector-id {connector_id} --table")
            typer.echo(f"   copilot connections delete <connection-id> -c {connector_id} --force")
            typer.echo(f"   copilot connections create -c {connector_id} -n \"<name>\" --oauth")
            typer.echo(f"   copilot connection-references update <ref-id> --connection-id <new-connection-id>")

    except Exception as e:
        exit_code = handle_api_error(e)
        raise typer.Exit(exit_code)


@app.command("register")
def connectors_register(
    connector_id: str = typer.Argument(
        ...,
        help="The connector's unique identifier (e.g., shared_asana-20custom-...)",
    ),
    swagger_file: str = typer.Option(
        ...,
        "--swagger-file",
        "-f",
        help="Path to the original OpenAPI 2.0 (Swagger) definition file used to create the connector",
    ),
    force: bool = typer.Option(
        False,
        "--force",
        help="Skip confirmation prompt",
    ),
):
    """
    Register a custom connector in Dataverse.

    Custom connectors created via Power Apps API are not automatically registered
    in Dataverse. This command creates a record in the Dataverse connector table
    so that connection references can properly link to the connector via
    CustomConnectorId.

    This is required for connector operations to be properly discovered by
    Copilot Studio agents. Without Dataverse registration, you may see
    "ConnectorOperationNotFound" errors.

    IMPORTANT: You must provide the ORIGINAL OpenAPI schema file that was used
    to create the connector. The schema stored in Power Apps is modified and
    cannot be used directly.

    Examples:
        copilot connectors register shared_asana-20custom-... --swagger-file ./connector.json
        copilot connectors register <connector-id> -f ./api.json --force
    """
    try:
        client = get_client()

        # Read and parse the original OpenAPI file
        swagger_path = Path(swagger_file)
        if not swagger_path.exists():
            typer.echo(f"Error: File not found: {swagger_file}", err=True)
            raise typer.Exit(1)

        try:
            file_content = swagger_path.read_text()
            try:
                swagger = json.loads(file_content)
            except json.JSONDecodeError:
                try:
                    swagger = yaml.safe_load(file_content)
                except yaml.YAMLError as yaml_err:
                    typer.echo(f"Error: Invalid JSON/YAML format: {yaml_err}", err=True)
                    raise typer.Exit(1)
        except Exception as e:
            typer.echo(f"Error reading file: {e}", err=True)
            raise typer.Exit(1)

        # Get connector details from Power Apps API
        typer.echo(f"Looking up connector: {connector_id}...")
        connector = client.get_connector(connector_id)

        props = connector.get("properties", {})
        display_name = props.get("displayName", connector_id)
        description = props.get("description", "")

        # Check if already registered in Dataverse
        dataverse = connector.get("_dataverse", {})
        existing_entity_id = dataverse.get("connectorid")

        if existing_entity_id:
            typer.echo(f"Connector '{display_name}' is already registered in Dataverse.")
            typer.echo(f"Entity ID: {existing_entity_id}")
            return

        source = connector.get("_source", "unknown")
        if source != "powerapps":
            typer.echo(f"Connector source: {source}")
            typer.echo("Only Power Apps connectors need to be registered in Dataverse.")
            typer.echo("This connector may already be properly registered.")
            if not force:
                typer.echo("Use --force to attempt registration anyway.")
                raise typer.Exit(0)

        # Confirm registration
        if not force:
            typer.echo(f"\nConnector: {display_name}")
            typer.echo(f"ID: {connector_id}")
            typer.echo(f"Source: {source}")
            typer.echo()
            typer.echo("This will register the connector in Dataverse, enabling proper")
            typer.echo("connection reference linking for Copilot Studio agents.")
            typer.echo()
            typer.echo("Continue? (y/N): ", nl=False)

            response = input().strip().lower()
            if response != 'y':
                typer.echo("Cancelled.")
                raise typer.Exit(0)

        # Register in Dataverse
        typer.echo("\nRegistering connector in Dataverse...")
        entity_id = client.register_connector_in_dataverse(
            connector_id=connector_id,
            display_name=display_name,
            openapi_definition=swagger,
            description=description,
        )

        if entity_id:
            print_success(f"Connector '{display_name}' registered successfully!")
            typer.echo(f"Dataverse Entity ID: {entity_id}")
            typer.echo()
            typer.echo("Next steps:")
            typer.echo("1. Recreate the connection reference to link it properly:")
            typer.echo(f"   copilot connection-references list --table")
            typer.echo(f"   copilot connection-references remove <ref-id> --force")
            typer.echo(f"   copilot connection-references create --name \"<name>\" --connection-id <conn-id>")
        else:
            typer.echo("Failed to register connector in Dataverse.", err=True)
            typer.echo("This may be a permissions issue or the connector schema may be invalid.", err=True)
            raise typer.Exit(1)

    except Exception as e:
        exit_code = handle_api_error(e)
        raise typer.Exit(exit_code)


@app.command("delete")
@app.command("remove")
def connectors_delete(
    connector_id: str = typer.Argument(
        ...,
        help="The connector's unique identifier (e.g., shared_asana-20test-5fd251...)",
    ),
    environment: Optional[str] = typer.Option(
        None,
        "--environment",
        "--env",
        help="Power Platform environment ID. Uses DATAVERSE_ENVIRONMENT_ID if not specified.",
    ),
    cascade: bool = typer.Option(
        False,
        "--cascade",
        help="Also delete all connections, connection references, and agent tools associated with this connector",
    ),
    force: bool = typer.Option(
        False,
        "--force",
        "-f",
        help="Skip confirmation prompt",
    ),
):
    """
    Delete a custom connector.

    This permanently removes a custom connector from the Power Platform environment.
    Only custom connectors can be deleted; managed (Microsoft) connectors cannot be deleted.

    Warning: Deleting a connector may break flows or agents that depend on it.

    Use --cascade to recursively delete everything associated with this connector:
      1. Agent connector tools using connections for this connector
      2. Connection references pointing to connections for this connector
      3. Connections created for this connector
      4. The connector itself

    Examples:
        copilot connectors delete shared_asana-20test-5fd251d00ef0afcb57-5fe2f45645c919b585
        copilot connectors delete <connector-id> --force
        copilot connectors delete <connector-id> --env Default-xxx
        copilot connectors delete <connector-id> --cascade
    """
    try:
        client = get_client()

        # Resolve environment ID if not provided
        if not environment:
            from ..config import get_config
            config = get_config()
            environment = config.environment_id
            if not environment:
                typer.echo("Error: Environment ID not found. Please set DATAVERSE_ENVIRONMENT_ID.", err=True)
                raise typer.Exit(1)

        # Get connector details to show what will be deleted
        try:
            connector = client.get_connector(connector_id, environment)
            props = connector.get("properties", {})
            connector_name = props.get("displayName", connector_id)

            # Check if it's a custom connector
            if not is_custom_connector(connector):
                typer.echo(f"Error: Cannot delete managed connector '{connector_name}'", err=True)
                typer.echo("Only custom connectors can be deleted.", err=True)
                raise typer.Exit(1)
                
            # Valid connector found
            typer.echo(f"Verified connector: {connector_name}")

        except Exception as e:
            # Handle specific error cases
            error_msg = str(e)
            
            # Check for 404 Not Found (often buried in the error message)
            if "404" in error_msg or "NotFound" in error_msg:
                typer.echo(f"Error: Connector '{connector_id}' not found in environment '{environment}'.", err=True)
                if not force:
                    typer.echo("Aborting. Use --force to attempt deletion anyway.", err=True)
                    raise typer.Exit(1)
                typer.echo("Warning: Connector not found, but proceeding due to --force.", err=True)
                connector_name = connector_id

            # Check for 403 Forbidden
            elif "403" in error_msg or "Forbidden" in error_msg or "Access Denied" in error_msg:
                typer.echo(f"Error: Permission denied for connector '{connector_id}'.", err=True)
                typer.echo("This is likely due to a Data Loss Prevention (DLP) policy blocking access.", err=True)
                if not force:
                    typer.echo("Aborting. Use --force to attempt deletion anyway.", err=True)
                    raise typer.Exit(1)
                typer.echo("Warning: Permission denied, but proceeding due to --force.", err=True)
                connector_name = connector_id
            
            else:
                # Unknown error
                typer.echo(f"Warning: Could not verify connector details: {e}", err=True)
                if not force:
                    typer.echo("Aborting. Use --force to attempt deletion anyway.", err=True)
                    raise typer.Exit(1)
                connector_name = connector_id

        # Confirm deletion unless --force
        if not force:
            typer.echo(f"\nConnector: {connector_name}")
            typer.echo(f"ID: {connector_id}")
            typer.echo()
            typer.echo("⚠️  Warning: This will permanently delete the connector.")
            typer.echo("   Any flows or agents using this connector may break.")
            typer.echo()
            typer.echo("Continue? (y/N): ", nl=False)

            response = input().strip().lower()
            if response != 'y':
                typer.echo("Cancelled.")
                raise typer.Exit(0)

        # Handle cascade deletion
        if cascade:
            typer.echo("\nCascading deletion requested. Checking for associated resources...")

            # 1. Find and delete all agent connector tools using this connector
            all_connector_tools = client.list_tools(category='connector')
            # Filter tools that use this connector (connector_id appears in connectionReference path)
            matching_tools = [
                tool for tool in all_connector_tools
                if connector_id in (tool.get("data") or "")
            ]

            if matching_tools:
                typer.echo(f"Found {len(matching_tools)} agent tool(s) using this connector. Deleting...")
                for tool in matching_tools:
                    tool_id = tool.get("botcomponentid")
                    tool_name = tool.get("name") or tool.get("schemaname") or tool_id
                    try:
                        # Connection refs are deleted separately in cascade logic below
                        client.remove_tool(tool_id)
                        typer.echo(f"  ✓ Deleted tool: {tool_name}")
                    except Exception as e:
                        typer.echo(f"  ✗ Failed to delete tool {tool_name}: {e}", err=True)
            else:
                typer.echo("No agent tools found using this connector.")

            # 2. Delete all connection references for this connector (by connector_id)
            refs = client.list_connection_references(connector_id=connector_id)
            if refs:
                typer.echo(f"Found {len(refs)} connection reference(s). Deleting...")
                for ref in refs:
                    ref_id = ref.get("connectionreferenceid")
                    ref_name = ref.get("connectionreferencedisplayname", "Unnamed")
                    try:
                        client.delete_connection_reference(ref_id)
                        typer.echo(f"  ✓ Deleted reference: {ref_name}")
                    except Exception as e:
                        typer.echo(f"  ✗ Failed to delete reference {ref_name}: {e}", err=True)
            else:
                typer.echo("No connection references found for this connector.")

            # 3. Get all connections for this connector and delete them
            connections = client.list_connections(connector_id, environment)
            if connections:
                typer.echo(f"Found {len(connections)} connection(s). Deleting...")
                for conn in connections:
                    conn_id = conn.get("name")
                    conn_name = conn.get("properties", {}).get("displayName", conn_id)
                    try:
                        client.delete_connection(conn_id, connector_id, environment)
                        typer.echo(f"  ✓ Deleted connection: {conn_name}")
                    except Exception as e:
                        typer.echo(f"  ✗ Failed to delete connection {conn_name}: {e}", err=True)
            else:
                typer.echo("No connections found for this connector.")

            typer.echo("\nProceeding to delete connector...")

        # Delete the connector
        client.delete_custom_connector(connector_id, environment)

        print_success(f"Connector '{connector_name}' deleted successfully!")
        typer.echo(f"Connector ID: {connector_id}")

    except Exception as e:
        exit_code = handle_api_error(e)
        raise typer.Exit(exit_code)
