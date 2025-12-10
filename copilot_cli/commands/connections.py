"""Connection commands for managing Power Platform connections."""
import typer
from typing import Optional

from ..client import get_client
from ..config import get_config
from ..output import print_json, print_table, print_success, handle_api_error


app = typer.Typer(help="Manage Power Platform connections (authenticated credentials)")


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

    # Extract connector name from apiId if connector_id not provided
    # apiId format: /providers/Microsoft.PowerApps/apis/shared_asana
    if not connector_id:
        api_id = props.get("apiId", "")
        if api_id:
            parts = api_id.split("/")
            connector_id = parts[-1] if parts else ""

    return {
        "name": display_name,
        "id": connection.get("name", ""),
        "connector": connector_id,
        "status": status_str,
        "created": created,
        "error": error_msg,
    }


@app.command("get")
def connections_get(
    connection_id: str = typer.Argument(
        ...,
        help="The connection's unique identifier (GUID)",
    ),
    environment: Optional[str] = typer.Option(
        None,
        "--environment",
        "--env",
        help="Power Platform environment ID. Uses DATAVERSE_ENVIRONMENT_ID if not specified.",
    ),
):
    """
    Get details for a specific connection by ID.

    Returns the full connection object including connector ID, display name,
    authentication status, and creation time.

    Examples:
        copilot connections get 5d8c58af-19db-4b51-b63b-cb543e53d9ba
        copilot connections get abc123 --env Default-xxx
    """
    try:
        client = get_client()

        # Get environment ID from config if not provided
        if not environment:
            config = get_config()
            environment = config.environment_id

        connection = client.get_connection(connection_id, environment)
        formatted = format_connection_for_display(connection)
        print_json(formatted)

    except Exception as e:
        exit_code = handle_api_error(e)
        raise typer.Exit(exit_code)


@app.command("list")
def connections_list(
    connector_id: Optional[str] = typer.Option(
        None,
        "--connector-id",
        "-c",
        help="Filter to a specific connector (e.g., shared_asana, shared_office365). If not provided, lists all connections.",
    ),
    table: bool = typer.Option(
        False,
        "--table",
        "-t",
        help="Display output as a formatted table instead of JSON",
    ),
):
    """
    List connections in the environment.

    Connections are authenticated credentials that allow Power Platform
    to access external services on your behalf. Each connection stores
    OAuth tokens, API keys, or other authentication details.

    By default, lists all connections in the environment. Use --connector-id
    to filter to a specific connector.

    Examples:
        copilot connections list --table
        copilot connections list --connector-id shared_asana --table
        copilot connections list -c shared_office365 --table
    """
    try:
        client = get_client()
        connections = client.list_connections(connector_id)

        if not connections:
            if connector_id:
                typer.echo(f"No connections found for connector '{connector_id}'.")
                typer.echo("\nThis could mean:")
                typer.echo("  - No connections have been created for this connector")
                typer.echo("  - The connector ID might be incorrect")
                typer.echo("\nUse 'copilot connectors list --table' to see available connectors.")
            else:
                typer.echo("No connections found in the environment.")
            return

        formatted = [format_connection_for_display(c, connector_id or "") for c in connections]

        if table:
            print_table(
                formatted,
                columns=["name", "connector", "id", "status", "created"],
                headers=["Name", "Connector", "Connection ID", "Status", "Created"],
            )
        else:
            print_json(formatted)

    except Exception as e:
        exit_code = handle_api_error(e)
        raise typer.Exit(exit_code)


@app.command("test")
def connections_test(
    connector_id: str = typer.Option(
        ...,
        "--connector-id",
        "-c",
        help="The connector's unique identifier (e.g., shared_asana, shared_office365)",
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
):
    """
    Test authentication for connector connections.

    This command checks the authentication status of connections for a
    given connector. It reads the status from the connection properties.

    Connection statuses:
      - Connected: Connection is authenticated and ready to use
      - Error: Connection has an authentication or configuration issue
      - Unauthenticated: Connection needs to be authenticated

    Examples:
        copilot connections test --connector-id shared_asana
        copilot connections test -c shared_office365 --table
        copilot connections test -c shared_asana --connection-id abc123
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
            typer.echo("\nUse 'copilot connectors list --table' to see available connectors.")
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
            auth_result = "OK" if is_healthy else current_status

            result = {
                "connection_id": conn_id,
                "display_name": display_name[:40] if len(display_name) > 40 else display_name,
                "status": current_status,
                "auth_result": auth_result,
                "healthy": is_healthy,
                "error": status_error,
            }

            results.append(result)

            # Print progress
            status_icon = "+" if is_healthy else "x"
            typer.echo(f"  {status_icon} {display_name[:50]} ({current_status})")

        typer.echo("")

        # Summary
        healthy_count = sum(1 for r in results if r["healthy"])
        unhealthy_count = len(results) - healthy_count

        if table:
            print_table(
                results,
                columns=["display_name", "status", "auth_result", "error"],
                headers=["Connection Name", "Status", "Auth Check", "Error"],
            )
        else:
            print_json(results)

        if unhealthy_count > 0:
            typer.echo(
                f"\nSummary: {healthy_count} healthy, {unhealthy_count} unhealthy "
                f"out of {len(results)} connection(s)"
            )
            raise typer.Exit(1)
        else:
            typer.echo(f"\nSummary: All {len(results)} connection(s) are healthy")

    except typer.Exit:
        raise
    except Exception as e:
        exit_code = handle_api_error(e)
        raise typer.Exit(exit_code)


@app.command("create")
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
    no_wait: bool = typer.Option(
        False,
        "--no-wait",
        help="Don't wait for OAuth authentication to complete (just open browser and exit).",
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
        copilot connections create -c shared_asana -n "My Asana" --oauth

        # Azure AI Search
        copilot connections create -c shared_azureaisearch -n "My Search" \\
            --parameters '{"endpoint": "https://mysearch.search.windows.net", "api_key": "xxx"}'

        # API key connector
        copilot connections create -c shared_sendgrid -n "SendGrid" \\
            --parameters '{"api_key": "SG.xxx"}'
    """
    import json

    try:
        client = get_client()

        # Get environment ID from config if not provided
        if not environment:
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
            import webbrowser
            import time

            # Show OAuth redirect URL configuration requirement
            typer.echo()
            typer.echo("⚠️  OAuth Redirect URL Configuration Required")
            typer.echo()
            typer.echo("Power Platform will use this redirect URL for OAuth:")
            typer.echo()
            typer.echo(f"  https://global.consent.azure-apim.net/redirect/{connector_id}")
            typer.echo()
            typer.echo("You must register this EXACT URL in your OAuth app settings.")
            typer.echo()
            typer.echo("💡 Tip: If your OAuth provider supports wildcards, register:")
            typer.echo("     https://global.consent.azure-apim.net/redirect/*")
            typer.echo("     This will work for all connectors you create.")
            typer.echo()
            typer.echo("Have you registered the redirect URL? (y/N): ", nl=False)

            response = input().strip().lower()
            if response != 'y':
                typer.echo()
                typer.echo("Connection creation cancelled.")
                typer.echo("Register the redirect URL in your OAuth app and try again.")
                raise typer.Exit(0)

            typer.echo()

            # OAuth flow - create connection and get consent link
            result = client.create_oauth_connection(
                connector_id=connector_id,
                connection_name=name,
                environment_id=environment,
            )

            connection_id = result.get("name", "")

            print_success(f"Connection '{name}' created.")
            typer.echo(f"Connection ID: {connection_id}")
            typer.echo(f"Connector: {connector_id}")
            typer.echo("")

            # Get the consent link and open browser
            typer.echo("Getting OAuth consent link...")
            consent_link = client.get_consent_link(connector_id, connection_id, environment)

            if not consent_link:
                typer.echo("Error: Could not get consent link from API.", err=True)
                typer.echo(f"Complete authentication manually at:")
                typer.echo(f"  https://make.powerapps.com/environments/{environment}/connections")
                raise typer.Exit(1)

            typer.echo("Opening browser for OAuth authentication...")
            webbrowser.open(consent_link)

            if no_wait:
                typer.echo(f"Check connection status: copilot connections list -c {connector_id} --table")
                return

            # Poll for connection status
            typer.echo("")
            typer.echo("Waiting for authentication to complete...")
            typer.echo("(Complete the OAuth flow in your browser, then return here)")
            typer.echo("")

            max_attempts = 60  # 5 minutes at 5-second intervals
            poll_interval = 5

            for attempt in range(max_attempts):
                time.sleep(poll_interval)

                try:
                    connections = client.list_connections(connector_id, environment)
                    conn = next((c for c in connections if c.get("name") == connection_id), None)

                    if conn:
                        statuses = conn.get("properties", {}).get("statuses", [])
                        if statuses:
                            status = statuses[0].get("status", "Unknown")
                            if status.lower() == "connected":
                                typer.echo("")
                                print_success(f"Authentication complete! Connection '{name}' is now connected.")
                                return

                    # Show progress
                    elapsed = (attempt + 1) * poll_interval
                    typer.echo(f"  Still waiting... ({elapsed}s elapsed)", nl=False)
                    typer.echo("\r", nl=False)

                except Exception:
                    # Ignore polling errors, keep trying
                    pass

            typer.echo("")
            typer.echo("Timed out waiting for authentication.")
            typer.echo(f"Check connection status: copilot connections list -c {connector_id} --table")

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

    except typer.Exit:
        raise
    except Exception as e:
        exit_code = handle_api_error(e)
        raise typer.Exit(exit_code)


@app.command("delete")
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
        copilot connections delete <guid> -c shared_asana
        copilot connections delete <guid> -c shared_office365 --force
        copilot connections delete <guid> -c shared_azureaisearch --env Default-xxx
    """
    try:
        client = get_client()

        # Get environment ID from config if not provided
        if not environment:
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
