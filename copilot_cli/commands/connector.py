"""Connector commands for listing available Power Platform connectors."""
import typer
from typing import Optional

from ..client import get_client
from ..output import print_json, print_table, handle_api_error


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
    connector_id: str = typer.Option(
        ...,
        "--connector-id",
        "-c",
        help="The connector's unique identifier (e.g., shared_office365)",
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
    List connections for a specific connector.

    This command lists all authenticated connections (instances) for a
    given connector. Each connection represents a user's authentication
    to the connector's backend service.

    Examples:
        copilot tool connector connections list --connector-id shared_office365
        copilot tool connector connections list -c shared_commondataserviceforapps --table
        copilot tool connector connections list -c shared_podio --connection-id abc123
    """
    try:
        client = get_client()
        connections = client.list_connections(connector_id)

        if not connections:
            typer.echo(f"No connections found for connector '{connector_id}'.")
            typer.echo("\nThis could mean:")
            typer.echo("  - No connections have been created for this connector")
            typer.echo("  - The connector ID might be incorrect")
            typer.echo("\nUse 'copilot tool connector list --table' to see available connectors.")
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
        copilot tool connector connections auth-test --connector-id shared_office365
        copilot tool connector connections auth-test -c shared_commondataserviceforapps --table
        copilot tool connector connections auth-test -c shared_podio --connection-id abc123
        copilot tool connector connections auth-test -c shared_office365 --test-api
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
            typer.echo("\nUse 'copilot tool connector list --table' to see available connectors.")
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
