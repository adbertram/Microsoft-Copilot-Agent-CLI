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

    # Check for custom connector indicators
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

    description = props.get("description") or ""
    if len(description) > 60:
        description = description[:57] + "..."

    is_custom = is_custom_connector(connector)

    return {
        "name": props.get("displayName") or connector.get("name", ""),
        "id": connector.get("name", ""),
        "type": "Custom" if is_custom else "Managed",
        "publisher": props.get("publisher") or "",
        "tier": props.get("tier") or "N/A",
        "description": description,
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
):
    """
    List all available connectors in the environment.

    Connectors are proxies/wrappers around APIs that define what actions
    are available (e.g., Asana, SharePoint, SQL Server). They represent
    the "type" of service you can connect to.

    Connector Types:
      - Managed: Built-in connectors published by Microsoft
      - Custom: User-created connectors in the environment

    Examples:
        copilot connectors list
        copilot connectors list --table
        copilot connectors list --custom --table
        copilot connectors list --managed --table
        copilot connectors list --filter "asana" --table
    """
    if custom and managed:
        typer.echo("Error: Cannot specify both --custom and --managed", err=True)
        raise typer.Exit(1)

    try:
        client = get_client()
        connectors = client.list_connectors()

        if not connectors:
            typer.echo("No connectors found.")
            return

        # Filter by custom/managed
        if custom:
            connectors = [c for c in connectors if is_custom_connector(c)]
        elif managed:
            connectors = [c for c in connectors if not is_custom_connector(c)]

        # Filter by text
        if filter_text:
            filter_lower = filter_text.lower()
            connectors = [
                c for c in connectors
                if filter_lower in c.get("properties", {}).get("displayName", "").lower()
                or filter_lower in c.get("properties", {}).get("publisher", "").lower()
                or filter_lower in c.get("name", "").lower()
            ]

        if not connectors:
            typer.echo("No connectors match the filter criteria.")
            return

        formatted = [format_connector_for_display(c) for c in connectors]

        # Sort by type (Custom first) then name
        formatted.sort(key=lambda x: (0 if x["type"] == "Custom" else 1, x["name"].lower()))

        if table:
            print_table(
                formatted,
                columns=["name", "type", "publisher", "tier", "id"],
                headers=["Name", "Type", "Publisher", "Tier", "ID"],
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
    """
    try:
        client = get_client()
        connector = client.get_connector(connector_id)

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
):
    """
    Create a new custom connector from an OpenAPI 2.0 (Swagger) definition.

    The OpenAPI definition must be in OpenAPI 2.0 format (not 3.0).

    For OAuth connectors, provide --oauth-client-id and --oauth-client-secret.
    Without these, the connector will be created but connections cannot authenticate.

    After creating the connector, use 'copilot connections create' to create a
    connection and authenticate.

    Examples:
      # Basic connector (non-OAuth)
      copilot connectors create --name "My API" --swagger-file ./api.json

      # OAuth connector with credentials
      copilot connectors create --name "My API" --swagger-file ./api.json \\
        --oauth-client-id "client123" --oauth-client-secret "secret456"

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
        )

        connector_id = result["connector_id"]
        environment_id = result["environment_id"]

        print_success(f"Custom connector '{name}' created successfully!")
        typer.echo(f"Connector ID: {connector_id}")
        typer.echo(f"Environment: {environment_id}")
        typer.echo()

        # Show next steps with redirect URL info for OAuth connectors
        typer.echo("⚠️  Next Steps:")

        if uses_oauth:
            typer.echo()
            typer.echo("1. Register this redirect URL in your OAuth app settings:")
            typer.echo(f"   https://global.consent.azure-apim.net/redirect/{connector_id}")
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


@app.command("delete")
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

    Examples:
        copilot connectors delete shared_asana-20test-5fd251d00ef0afcb57-5fe2f45645c919b585
        copilot connectors delete <connector-id> --force
        copilot connectors delete <connector-id> --env Default-xxx
    """
    try:
        client = get_client()

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

        except Exception as e:
            # If we can't get the connector, it might not exist
            typer.echo(f"Warning: Could not verify connector details: {e}", err=True)
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

        # Delete the connector
        client.delete_custom_connector(connector_id, environment)

        print_success(f"Connector '{connector_name}' deleted successfully!")
        typer.echo(f"Connector ID: {connector_id}")

    except Exception as e:
        exit_code = handle_api_error(e)
        raise typer.Exit(exit_code)
