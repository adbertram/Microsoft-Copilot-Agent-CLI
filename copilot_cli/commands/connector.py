"""Connector commands for listing available Power Platform connectors."""
import typer
from typing import Optional

from ..client import get_client
from ..output import print_json, print_table, print_success, handle_api_error


app = typer.Typer(help="Manage Power Platform connectors")

# Subcommand group for connections
connections_app = typer.Typer(help="Manage connector connections")
app.add_typer(connections_app, name="connections")

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


def format_connection_reference_for_display(conn_ref: dict) -> dict:
    """Format a connection reference for display."""
    display_name = conn_ref.get("connectionreferencedisplayname") or ""
    if len(display_name) > 40:
        display_name = display_name[:37] + "..."

    logical_name = conn_ref.get("connectionreferencelogicalname") or ""
    connector_id = conn_ref.get("connectorid") or ""

    # Extract connector name from connector_id path
    # Format: /providers/Microsoft.PowerApps/apis/shared_office365
    connector_name = ""
    if connector_id:
        parts = connector_id.split("/")
        if parts:
            connector_name = parts[-1]

    state = conn_ref.get("statecode")
    state_str = "Active" if state == 0 else "Inactive" if state == 1 else "Unknown"

    return {
        "name": display_name,
        "logical_name": logical_name,
        "id": conn_ref.get("connectionreferenceid") or "",
        "connector": connector_name,
        "state": state_str,
    }


def format_connection_for_display(connection: dict, connector_id: str = "") -> dict:
    """Format a connection for display."""
    props = connection.get("properties", {})
    statuses = props.get("statuses", [])

    # Extract status info
    status_str = "Unknown"
    error_msg = ""
    if statuses:
        first_status = statuses[0] if isinstance(statuses, list) else statuses
        status_str = first_status.get("status", "Unknown")
        if first_status.get("error"):
            err = first_status["error"]
            if isinstance(err, dict):
                error_msg = err.get("message", "")[:50]
            else:
                error_msg = str(err)[:50]

    display_name = props.get("displayName") or connection.get("name", "")
    if len(display_name) > 40:
        display_name = display_name[:37] + "..."

    created = props.get("createdTime", "")
    if created:
        created = created[:10]

    return {
        "name": display_name,
        "id": connection.get("name", ""),
        "connector_id": connector_id,
        "status": status_str,
        "created": created,
        "error": error_msg,
    }


@app.command("list")
def connector_list(
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

    This command lists all connectors available in the Power Platform
    environment, including both managed (Microsoft) and custom connectors.

    Connector Types:
      - Managed: Built-in connectors published by Microsoft
      - Custom: User-created connectors in the environment

    Examples:
        copilot connector list
        copilot connector list --table
        copilot connector list --custom --table
        copilot connector list --managed --table
        copilot connector list --filter "office365" --table
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


@app.command("get")
def connector_get(
    connector_id: str = typer.Argument(
        ...,
        help="The connector's unique identifier (e.g., shared_office365)",
    ),
):
    """
    Get details for a specific connector.

    Examples:
        copilot connector get shared_office365
        copilot connector get shared_sharepointonline
    """
    try:
        client = get_client()
        connector = client.get_connector(connector_id)
        print_json(connector)
    except Exception as e:
        exit_code = handle_api_error(e)
        raise typer.Exit(exit_code)


# =============================================================================
# Connections Subcommands
# =============================================================================


@connections_app.command("list")
def connections_list(
    connector_id: Optional[str] = typer.Option(
        None,
        "--connector-id",
        "-c",
        help="The connector's unique identifier (e.g., shared_office365). If not provided, lists all connection references.",
    ),
    connection_id: Optional[str] = typer.Option(
        None,
        "--connection-id",
        help="Filter to a specific connection ID",
    ),
    table: bool = typer.Option(
        False,
        "--table",
        "-t",
        help="Display output as a formatted table instead of JSON",
    ),
):
    """
    List connections or connection references.

    When --connector-id is provided, lists authenticated connections for that
    specific connector. Each connection represents a user's authentication
    to the connector's backend service.

    When no --connector-id is provided, lists all connection references in the
    Dataverse environment. Connection references are the links between Power
    Platform solutions and connector connections.

    Examples:
        copilot connector connections list --table
        copilot connector connections list --connector-id shared_office365
        copilot connector connections list -c shared_commondataserviceforapps --table
        copilot connector connections list -c shared_podio --connection-id abc123
    """
    try:
        client = get_client()

        # If no connector_id provided, list all connection references
        if not connector_id:
            conn_refs = client.list_connection_references()

            if not conn_refs:
                typer.echo("No connection references found in the environment.")
                return

            formatted = [format_connection_reference_for_display(cr) for cr in conn_refs]

            if table:
                print_table(
                    formatted,
                    columns=["name", "logical_name", "connector", "state", "id"],
                    headers=["Name", "Logical Name", "Connector", "State", "ID"],
                )
            else:
                print_json(formatted)
            return

        # Otherwise, list connections for specific connector
        connections = client.list_connections(connector_id)

        if not connections:
            typer.echo(f"No connections found for connector '{connector_id}'.")
            typer.echo("\nThis could mean:")
            typer.echo("  - No connections have been created for this connector")
            typer.echo("  - The connector ID might be incorrect")
            typer.echo("\nUse 'copilot connector list --table' to see available connectors.")
            return

        # Filter to specific connection if requested
        if connection_id:
            connections = [c for c in connections if c.get("name") == connection_id]
            if not connections:
                typer.echo(f"Connection '{connection_id}' not found for connector '{connector_id}'.")
                raise typer.Exit(1)

        formatted = [format_connection_for_display(c, connector_id) for c in connections]

        if table:
            print_table(
                formatted,
                columns=["name", "id", "status", "created", "error"],
                headers=["Name", "Connection ID", "Status", "Created", "Error"],
            )
        else:
            print_json(formatted)

    except Exception as e:
        exit_code = handle_api_error(e)
        raise typer.Exit(exit_code)


@connections_app.command("auth-test")
def connections_auth_test(
    connector_id: str = typer.Option(
        ...,
        "--connector-id",
        "-c",
        help="The connector's unique identifier (e.g., shared_office365)",
    ),
    connection_id: Optional[str] = typer.Option(
        None,
        "--connection-id",
        help="Test a specific connection ID. If not provided, tests all connections.",
    ),
    table: bool = typer.Option(
        False,
        "--table",
        "-t",
        help="Display output as a formatted table instead of JSON",
    ),
    test_api: bool = typer.Option(
        False,
        "--test-api",
        help="Also call the testConnection API endpoint (not all connectors support this).",
    ),
):
    """
    Test authentication for connector connections.

    This command checks the authentication status of connections for a
    given connector. By default, it reads the status from the connection
    properties. Use --test-api to also call the testConnection API endpoint
    (note: not all connectors implement this endpoint).

    Connection statuses:
      - Connected: Connection is authenticated and ready to use
      - Error: Connection has an authentication or configuration issue
      - Unauthenticated: Connection needs to be authenticated

    Examples:
        copilot connector connections auth-test --connector-id shared_office365
        copilot connector connections auth-test -c shared_commondataserviceforapps --table
        copilot connector connections auth-test -c shared_podio --connection-id abc123
        copilot connector connections auth-test -c shared_office365 --test-api
    """
    try:
        client = get_client()

        # List connections for this connector
        typer.echo(f"Finding connections for connector: {connector_id}...")
        connections = client.list_connections(connector_id)

        if not connections:
            typer.echo(f"No connections found for connector '{connector_id}'.")
            typer.echo("\nThis could mean:")
            typer.echo("  - No connections have been created for this connector")
            typer.echo("  - The connector ID might be incorrect")
            typer.echo("\nUse 'copilot connector list --table' to see available connectors.")
            return

        # If specific connection requested, filter to that one
        if connection_id:
            connections = [c for c in connections if c.get("name") == connection_id]
            if not connections:
                typer.echo(f"Connection '{connection_id}' not found for connector '{connector_id}'.")
                raise typer.Exit(1)

        typer.echo(f"Found {len(connections)} connection(s). Checking authentication status...\n")

        results = []
        for conn in connections:
            conn_id = conn.get("name", "")
            props = conn.get("properties", {})
            display_name = props.get("displayName") or conn_id

            # Get current status from connection object
            current_status = "Unknown"
            status_error = ""
            statuses = props.get("statuses", [])
            if statuses:
                first_status = statuses[0] if isinstance(statuses, list) else statuses
                current_status = first_status.get("status", "Unknown")
                if first_status.get("error"):
                    err = first_status["error"]
                    if isinstance(err, dict):
                        status_error = err.get("message", str(err))[:60]
                    else:
                        status_error = str(err)[:60]

            # Determine if connection is healthy based on status
            is_healthy = current_status.lower() == "connected"
            auth_result = "✓ OK" if is_healthy else "✗ " + current_status

            result = {
                "connection_id": conn_id,
                "display_name": display_name[:40] if len(display_name) > 40 else display_name,
                "status": current_status,
                "auth_result": auth_result,
                "healthy": is_healthy,
                "error": status_error,
            }

            # Optionally test via API endpoint
            if test_api:
                test_result = client.test_connection(connector_id, conn_id)
                if test_result.get("status_code") == 404:
                    result["api_test"] = "N/A (not implemented)"
                elif test_result.get("success"):
                    result["api_test"] = "✓ Passed"
                else:
                    result["api_test"] = "✗ Failed"
                    if test_result.get("error"):
                        result["error"] = test_result["error"][:60]

            results.append(result)

            # Print progress
            status_icon = "✓" if is_healthy else "✗"
            typer.echo(f"  {status_icon} {display_name[:50]} ({current_status})")

        typer.echo("")

        # Summary
        healthy_count = sum(1 for r in results if r["healthy"])
        unhealthy_count = len(results) - healthy_count

        if table:
            columns = ["display_name", "status", "auth_result", "error"]
            headers = ["Connection Name", "Status", "Auth Check", "Error"]
            if test_api:
                columns.insert(3, "api_test")
                headers.insert(3, "API Test")
            print_table(results, columns=columns, headers=headers)
        else:
            print_json(results)

        if unhealthy_count > 0:
            typer.echo(
                f"\nSummary: {healthy_count} healthy, {unhealthy_count} unhealthy "
                f"out of {len(results)} connection(s)"
            )
            raise typer.Exit(1)
        else:
            typer.echo(f"\nSummary: All {len(results)} connection(s) are healthy ✓")

    except Exception as e:
        exit_code = handle_api_error(e)
        raise typer.Exit(exit_code)


@connections_app.command("create")
def connections_create(
    connector_id: str = typer.Option(
        ...,
        "--connector-id",
        "-c",
        help="The connector's unique identifier (e.g., shared_asana, shared_office365)",
    ),
    name: str = typer.Option(
        ...,
        "--name",
        "-n",
        help="Display name for the connection",
    ),
    parameters: Optional[str] = typer.Option(
        None,
        "--parameters",
        "-p",
        help="JSON string of connection parameters (connector-specific). For OAuth connectors, use --oauth to initiate browser-based auth.",
    ),
    oauth: bool = typer.Option(
        False,
        "--oauth",
        help="Initiate OAuth authentication flow (opens browser). Use for OAuth-based connectors like Asana, SharePoint, etc.",
    ),
    environment: Optional[str] = typer.Option(
        None,
        "--environment",
        "--env",
        help="Power Platform environment ID. Uses DATAVERSE_ENVIRONMENT_ID if not specified.",
    ),
):
    """
    Create a new connection for a connector.

    Connections authenticate access to external services. Different connectors
    require different authentication methods:

    OAuth Connectors (--oauth):
      Use browser-based authentication for connectors like Asana, SharePoint,
      Dynamics 365, etc. This will output a consent URL to complete in browser.

    API Key Connectors (--parameters):
      Provide credentials directly via JSON. For example:
        --parameters '{"api_key": "xxx"}' for API key auth
        --parameters '{"username": "x", "password": "y"}' for basic auth

    Azure AI Search:
      Use specific parameters for Azure AI Search connections:
        --parameters '{"endpoint": "https://search.windows.net", "api_key": "xxx"}'

    Examples:
        # OAuth connector (Asana, SharePoint, etc.)
        copilot connector connections create -c shared_asana -n "My Asana" --oauth

        # Azure AI Search
        copilot connector connections create -c shared_azureaisearch -n "My Search" \\
            --parameters '{"endpoint": "https://mysearch.search.windows.net", "api_key": "xxx"}'

        # API key connector
        copilot connector connections create -c shared_sendgrid -n "SendGrid" \\
            --parameters '{"api_key": "SG.xxx"}'
    """
    import json
    import uuid

    try:
        client = get_client()

        # Get environment ID from config if not provided
        if not environment:
            from ..config import get_config
            config = get_config()
            environment = config.environment_id
            if not environment:
                typer.echo(
                    "Error: Environment ID not found. Please set DATAVERSE_ENVIRONMENT_ID "
                    "in your .env file or use --environment.",
                    err=True
                )
                raise typer.Exit(1)

        # Parse parameters if provided
        params_dict = {}
        if parameters:
            try:
                params_dict = json.loads(parameters)
            except json.JSONDecodeError as e:
                typer.echo(f"Error: Invalid JSON in --parameters: {e}", err=True)
                raise typer.Exit(1)

        if oauth:
            # OAuth flow - create connection and return consent URL
            result = client.create_oauth_connection(
                connector_id=connector_id,
                connection_name=name,
                environment_id=environment,
            )

            connection_id = result.get("name", "")
            consent_url = result.get("properties", {}).get("connectionParameters", {}).get("token", {}).get("oAuthSettings", {}).get("consentUrl")

            # Try to extract consent link from various locations
            if not consent_url:
                # Check for consent link in statuses
                statuses = result.get("properties", {}).get("statuses", [])
                for status in statuses:
                    if status.get("status") == "Unauthenticated":
                        error = status.get("error", {})
                        if isinstance(error, dict):
                            consent_url = error.get("message", "")
                            # Extract URL if embedded in message
                            if "https://" in consent_url:
                                import re
                                urls = re.findall(r'https://[^\s\'"]+', consent_url)
                                if urls:
                                    consent_url = urls[0]

            print_success(f"Connection '{name}' created.")
            typer.echo(f"Connection ID: {connection_id}")
            typer.echo(f"Connector: {connector_id}")
            typer.echo("")

            if consent_url and consent_url.startswith("https://"):
                typer.echo("OAuth authentication required. Complete the consent flow:")
                typer.echo(f"\n  {consent_url}\n")
                typer.echo("After completing authentication in your browser, the connection will be ready.")
            else:
                typer.echo("Note: Connection created but requires authentication.")
                typer.echo("Complete the OAuth flow in Power Platform admin center:")
                typer.echo(f"  https://make.powerapps.com/environments/{environment}/connections")

        elif connector_id == "shared_azureaisearch":
            # Azure AI Search has specific parameters
            endpoint = params_dict.get("endpoint") or params_dict.get("ConnectionEndpoint")
            api_key = params_dict.get("api_key") or params_dict.get("AdminKey")

            if not endpoint or not api_key:
                typer.echo(
                    "Error: Azure AI Search requires 'endpoint' and 'api_key' in --parameters",
                    err=True
                )
                typer.echo('Example: --parameters \'{"endpoint": "https://mysearch.search.windows.net", "api_key": "xxx"}\'')
                raise typer.Exit(1)

            result = client.create_azure_ai_search_connection(
                connection_name=name,
                search_endpoint=endpoint,
                api_key=api_key,
                environment_id=environment,
            )

            connection_id = result.get("name", "")
            display_name = result.get("properties", {}).get("displayName", name)
            statuses = result.get("properties", {}).get("statuses", [])
            status = statuses[0].get("status", "Unknown") if statuses else "Unknown"

            print_success(f"Connection '{display_name}' created successfully.")
            typer.echo(f"Connection ID: {connection_id}")
            typer.echo(f"Status: {status}")

        else:
            # Generic connection creation with parameters
            result = client.create_connection(
                connector_id=connector_id,
                connection_name=name,
                environment_id=environment,
                parameters=params_dict,
            )

            connection_id = result.get("name", "")
            props = result.get("properties", {})
            display_name = props.get("displayName", name)
            statuses = props.get("statuses", [])
            status = statuses[0].get("status", "Unknown") if statuses else "Unknown"

            print_success(f"Connection '{display_name}' created.")
            typer.echo(f"Connection ID: {connection_id}")
            typer.echo(f"Connector: {connector_id}")
            typer.echo(f"Status: {status}")

            if status == "Unauthenticated":
                typer.echo("")
                typer.echo("Note: Connection requires authentication.")
                typer.echo("Complete setup in Power Platform:")
                typer.echo(f"  https://make.powerapps.com/environments/{environment}/connections")

    except Exception as e:
        exit_code = handle_api_error(e)
        raise typer.Exit(exit_code)


@connections_app.command("delete")
def connections_delete(
    connection_id: str = typer.Argument(
        ...,
        help="The connection's unique identifier (GUID)",
    ),
    connector_id: str = typer.Option(
        ...,
        "--connector-id",
        "-c",
        help="The connector's unique identifier (e.g., shared_asana, shared_office365)",
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
    Delete a connector connection.

    Permanently removes a connection from the Power Platform environment.
    This may break flows or agents that depend on this connection.

    Examples:
        copilot connector connections delete <guid> -c shared_asana
        copilot connector connections delete <guid> -c shared_office365 --force
        copilot connector connections delete <guid> -c shared_azureaisearch --env Default-xxx
    """
    try:
        client = get_client()

        # Get environment ID from config if not provided
        if not environment:
            from ..config import get_config
            config = get_config()
            environment = config.environment_id
            if not environment:
                typer.echo(
                    "Error: Environment ID not found. Please set DATAVERSE_ENVIRONMENT_ID "
                    "in your .env file or use --environment.",
                    err=True
                )
                raise typer.Exit(1)

        # Try to get connection details first
        try:
            connections = client.list_connections(connector_id, environment)
            conn = next((c for c in connections if c.get("name") == connection_id), None)
            if conn:
                display_name = conn.get("properties", {}).get("displayName", connection_id)
                typer.echo(f"Connection: {display_name}")
                typer.echo(f"ID: {connection_id}")
                typer.echo(f"Connector: {connector_id}")
        except Exception:
            pass

        if not force:
            typer.echo("\nWARNING: This may break flows or agents using this connection.")
            confirm = typer.confirm("Are you sure you want to delete this connection?")
            if not confirm:
                typer.echo("Cancelled.")
                raise typer.Exit(0)

        client.delete_connection(connection_id, connector_id, environment)
        print_success(f"Connection {connection_id} deleted successfully.")

    except typer.Exit:
        raise
    except Exception as e:
        exit_code = handle_api_error(e)
        raise typer.Exit(exit_code)


@connections_app.command("remove")
def connections_remove(
    connection_ref_id: str = typer.Argument(
        ...,
        help="The connection reference's unique identifier (GUID)",
    ),
    force: bool = typer.Option(
        False,
        "--force",
        "-f",
        help="Skip confirmation prompt",
    ),
):
    """
    Remove a connection reference from the environment.

    This command deletes a connection reference from the Dataverse environment.
    Connection references are solution-aware links between Power Platform
    solutions and connector connections.

    Note: To delete an actual connection (not a reference), use:
        copilot connector connections delete <id> -c <connector-id>

    WARNING: Removing a connection reference may break flows or agents that
    depend on it. Use with caution.

    Examples:
        copilot connector connections remove 3562bdae-3fbb-f011-bbd3-000d3a8ba54e
        copilot connector connections remove 3562bdae-3fbb-f011-bbd3-000d3a8ba54e --force
    """
    try:
        client = get_client()

        # Try to get the connection reference first to show details
        try:
            conn_refs = client.list_connection_references()
            conn_ref = next(
                (cr for cr in conn_refs if cr.get("connectionreferenceid") == connection_ref_id),
                None
            )
        except Exception:
            conn_ref = None

        if conn_ref:
            display_name = conn_ref.get("connectionreferencedisplayname") or "Unknown"
            logical_name = conn_ref.get("connectionreferencelogicalname") or "Unknown"
            typer.echo(f"Connection Reference: {display_name}")
            typer.echo(f"Logical Name: {logical_name}")
            typer.echo(f"ID: {connection_ref_id}")

            # Block deletion of system-managed connection references
            if logical_name.startswith("msdyn_"):
                typer.echo("")
                typer.echo("Error: Cannot remove system-managed connection references.")
                typer.echo(
                    "Connection references with 'msdyn_' prefix are managed by Microsoft "
                    "and used internally by Power Platform/Copilot Studio."
                )
                raise typer.Exit(1)
        else:
            typer.echo(f"Connection Reference ID: {connection_ref_id}")

        if not force:
            typer.echo("\nWARNING: This may break flows or agents using this connection reference.")
            confirm = typer.confirm("Are you sure you want to remove this connection reference?")
            if not confirm:
                typer.echo("Cancelled.")
                raise typer.Exit(0)

        typer.echo("\nRemoving connection reference...")
        client.delete_connection_reference(connection_ref_id)
        typer.echo("Connection reference removed successfully.")

    except typer.Exit:
        raise
    except Exception as e:
        exit_code = handle_api_error(e)
        raise typer.Exit(exit_code)
