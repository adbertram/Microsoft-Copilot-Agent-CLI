"""REST API commands for listing custom connectors available as agent tools."""
import typer
from typing import Optional

from ..client import get_client
from ..output import print_json, print_table, handle_api_error


app = typer.Typer(help="Manage REST API tools (custom connectors)")


def format_restapi_for_display(connector: dict) -> dict:
    """Format a REST API connector for display."""
    name = connector.get("displayname") or connector.get("name", "")
    connector_id = connector.get("connectorid", "")

    # Get description (truncate if too long)
    description = connector.get("description") or ""
    if len(description) > 60:
        description = description[:57] + "..."

    # Get state
    state_code = connector.get("statecode", 0)
    state = "Active" if state_code == 0 else "Inactive"

    # Get owner
    owner = connector.get("_ownerid_value@OData.Community.Display.V1.FormattedValue", "")

    # Get created/modified dates
    created = connector.get("createdon", "")
    if created:
        created = created.split("T")[0]

    modified = connector.get("modifiedon", "")
    if modified:
        modified = modified.split("T")[0]

    # Check if managed
    is_managed = connector.get("ismanaged", False)

    return {
        "name": name,
        "id": connector_id,
        "state": state,
        "owner": owner,
        "description": description,
        "created": created,
        "modified": modified,
        "managed": is_managed,
    }


@app.command("list")
def restapi_list(
    filter_text: Optional[str] = typer.Option(
        None,
        "--filter",
        "-f",
        help="Filter by name or description (case-insensitive)",
    ),
    table: bool = typer.Option(
        False,
        "--table",
        "-t",
        help="Display output as a formatted table instead of JSON",
    ),
):
    """
    List all REST API tools (custom connectors) available for agents.

    REST API tools are custom connectors defined with OpenAPI specifications
    that can be attached to Copilot Studio agents as tools. They allow agents
    to call external REST APIs.

    Examples:
        copilot restapi list
        copilot restapi list --table
        copilot restapi list --filter "podio" --table
    """
    try:
        client = get_client()
        connectors = client.list_rest_apis()

        if not connectors:
            typer.echo("No REST API tools found.")
            return

        # Filter by text
        if filter_text:
            filter_lower = filter_text.lower()
            connectors = [
                c for c in connectors
                if filter_lower in (c.get("displayname") or c.get("name", "")).lower()
                or filter_lower in (c.get("description") or "").lower()
            ]

        if not connectors:
            typer.echo("No REST API tools match the filter criteria.")
            return

        formatted = [format_restapi_for_display(c) for c in connectors]

        # Sort by name
        formatted.sort(key=lambda x: x["name"].lower())

        if table:
            print_table(
                formatted,
                columns=["name", "state", "owner", "description", "id"],
                headers=["Name", "State", "Owner", "Description", "ID"],
            )
        else:
            print_json(formatted)
    except Exception as e:
        exit_code = handle_api_error(e)
        raise typer.Exit(exit_code)


@app.command("get")
def restapi_get(
    connector_id: str = typer.Argument(
        ...,
        help="The REST API connector's unique identifier (GUID)",
    ),
):
    """
    Get details for a specific REST API tool.

    Examples:
        copilot restapi get 56c1700d-a317-4472-8bd6-928afa5be754
    """
    try:
        client = get_client()
        connector = client.get_rest_api(connector_id)
        print_json(connector)
    except Exception as e:
        exit_code = handle_api_error(e)
        raise typer.Exit(exit_code)
