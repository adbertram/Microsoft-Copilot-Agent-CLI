"""Connection reference commands for managing solution-aware connection references."""
import typer
from typing import Optional

from ..client import get_client
from ..output import print_json, print_table, print_success, handle_api_error


app = typer.Typer(help="Manage connection references (solution-aware pointers to connections)")


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

    connection_id = conn_ref.get("connectionid") or ""

    return {
        "name": display_name,
        "logical_name": logical_name,
        "id": conn_ref.get("connectionreferenceid") or "",
        "connector": connector_name,
        "connection_id": connection_id,
        "state": state_str,
    }


@app.command("list")
def connection_references_list(
    table: bool = typer.Option(
        False,
        "--table",
        "-t",
        help="Display output as a formatted table instead of JSON",
    ),
):
    """
    List all connection references in the environment.

    Connection references are solution-aware pointers to connections. They
    allow flows and agents to reference a connection without being directly
    tied to it, making solutions portable across environments.

    Key concepts:
      - Connection references are environment-level resources
      - They point to actual connections (which hold credentials)
      - Multiple flows/agents can share the same connection reference
      - When moving solutions, you update the reference, not each flow

    Examples:
        copilot connection-references list
        copilot connection-references list --table
    """
    try:
        client = get_client()
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

    except Exception as e:
        exit_code = handle_api_error(e)
        raise typer.Exit(exit_code)


@app.command("get")
def connection_references_get(
    connection_ref_id: str = typer.Argument(
        ...,
        help="The connection reference's unique identifier (GUID)",
    ),
):
    """
    Get details for a specific connection reference.

    Returns the full connection reference record including the connector ID,
    connection ID it points to, and state.

    Examples:
        copilot connection-references get 3562bdae-3fbb-f011-bbd3-000d3a8ba54e
    """
    try:
        client = get_client()
        conn_ref = client.get_connection_reference(connection_ref_id)
        print_json(conn_ref)
    except Exception as e:
        exit_code = handle_api_error(e)
        raise typer.Exit(exit_code)


@app.command("update")
def connection_references_update(
    connection_ref_id: str = typer.Argument(
        ...,
        help="The connection reference's unique identifier (GUID)",
    ),
    connection_id: Optional[str] = typer.Option(
        None,
        "--connection-id",
        "-c",
        help="New connection ID to associate with this reference",
    ),
    name: Optional[str] = typer.Option(
        None,
        "--name",
        "-n",
        help="New display name for the connection reference",
    ),
):
    """
    Update a connection reference.

    You can change which connection the reference points to, or update
    its display name. This is useful when:
      - Moving solutions between environments (point to new connection)
      - Rotating credentials (point to connection with new credentials)
      - Renaming for clarity

    Examples:
        # Update the connection it points to
        copilot connection-references update <ref-id> --connection-id <new-conn-id>

        # Update the display name
        copilot connection-references update <ref-id> --name "Production Asana"

        # Update both
        copilot connection-references update <ref-id> -c <conn-id> -n "New Name"
    """
    if not connection_id and not name:
        typer.echo("Error: At least one of --connection-id or --name must be provided.", err=True)
        raise typer.Exit(1)

    try:
        client = get_client()
        result = client.update_connection_reference(
            connection_reference_id=connection_ref_id,
            connection_id=connection_id,
            display_name=name,
        )

        print_success("Connection reference updated successfully.")
        formatted = format_connection_reference_for_display(result)
        typer.echo(f"Name: {formatted['name']}")
        typer.echo(f"Connector: {formatted['connector']}")
        typer.echo(f"Connection ID: {formatted['connection_id']}")
        typer.echo(f"State: {formatted['state']}")

    except Exception as e:
        exit_code = handle_api_error(e)
        raise typer.Exit(exit_code)


@app.command("remove")
def connection_references_remove(
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

    This deletes the connection reference from Dataverse. It does NOT delete
    the underlying connection (credentials remain intact).

    WARNING: Removing a connection reference may break flows or agents that
    depend on it. Use with caution.

    Note: System-managed connection references (with 'msdyn_' prefix) cannot
    be removed as they are managed by Microsoft.

    Examples:
        copilot connection-references remove 3562bdae-3fbb-f011-bbd3-000d3a8ba54e
        copilot connection-references remove 3562bdae-3fbb-f011-bbd3-000d3a8ba54e --force
    """
    try:
        client = get_client()

        # Try to get the connection reference first to show details
        try:
            conn_ref = client.get_connection_reference(connection_ref_id)
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
        print_success("Connection reference removed successfully.")

    except typer.Exit:
        raise
    except Exception as e:
        exit_code = handle_api_error(e)
        raise typer.Exit(exit_code)
