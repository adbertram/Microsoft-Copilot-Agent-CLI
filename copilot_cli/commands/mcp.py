"""MCP server commands for listing Model Context Protocol servers available as agent tools."""
import typer
from typing import Optional

from ..client import get_client
from ..output import print_json, print_table, handle_api_error


app = typer.Typer(help="Manage MCP (Model Context Protocol) servers")


def format_mcp_for_display(connector: dict) -> dict:
    """Format an MCP server connector for display."""
    name = connector.get("name", "")
    props = connector.get("properties", {})
    display_name = props.get("displayName", name)

    # Get description (truncate if too long)
    description = props.get("description") or ""
    if len(description) > 60:
        description = description[:57] + "..."

    # Get tier
    tier = props.get("tier", "")

    # Get publisher
    publisher = props.get("publisher", "")

    # Get release tag (Preview, GA, etc.)
    release_tag = props.get("releaseTag", "")

    # Get created/modified dates
    created = props.get("createdTime", "")
    if created:
        created = created.split("T")[0]

    modified = props.get("changedTime", "")
    if modified:
        modified = modified.split("T")[0]

    return {
        "name": display_name,
        "id": name,
        "publisher": publisher,
        "tier": tier,
        "release": release_tag,
        "description": description,
        "created": created,
        "modified": modified,
    }


@app.command("list")
def mcp_list(
    filter_text: Optional[str] = typer.Option(
        None,
        "--filter",
        "-f",
        help="Filter by name, publisher, or description (case-insensitive)",
    ),
    table: bool = typer.Option(
        False,
        "--table",
        "-t",
        help="Display output as a formatted table instead of JSON",
    ),
):
    """
    List all MCP (Model Context Protocol) servers available for agents.

    MCP servers are connectors that implement the Model Context Protocol,
    allowing Copilot Studio agents to connect to external data sources and tools.
    They provide structured access to resources, tools, and prompts.

    Examples:
        copilot mcp list
        copilot mcp list --table
        copilot mcp list --filter "microsoft" --table
    """
    try:
        client = get_client()
        servers = client.list_mcp_servers()

        if not servers:
            typer.echo("No MCP servers found.")
            return

        # Filter by text
        if filter_text:
            filter_lower = filter_text.lower()
            servers = [
                s for s in servers
                if filter_lower in s.get("name", "").lower()
                or filter_lower in s.get("properties", {}).get("displayName", "").lower()
                or filter_lower in s.get("properties", {}).get("publisher", "").lower()
                or filter_lower in (s.get("properties", {}).get("description") or "").lower()
            ]

        if not servers:
            typer.echo("No MCP servers match the filter criteria.")
            return

        formatted = [format_mcp_for_display(s) for s in servers]

        # Sort by name
        formatted.sort(key=lambda x: x["name"].lower())

        if table:
            print_table(
                formatted,
                columns=["name", "publisher", "tier", "release", "description"],
                headers=["Name", "Publisher", "Tier", "Release", "Description"],
            )
        else:
            print_json(formatted)
    except Exception as e:
        exit_code = handle_api_error(e)
        raise typer.Exit(exit_code)


@app.command("get")
def mcp_get(
    connector_id: str = typer.Argument(
        ...,
        help="The MCP server connector's unique identifier (e.g., shared_microsoftlearndocsmcpserver)",
    ),
):
    """
    Get details for a specific MCP server.

    Examples:
        copilot mcp get shared_microsoftlearndocsmcpserver
        copilot mcp get shared_boxmcpserver
    """
    try:
        client = get_client()
        connector = client.get_mcp_server(connector_id)
        print_json(connector)
    except Exception as e:
        exit_code = handle_api_error(e)
        raise typer.Exit(exit_code)
