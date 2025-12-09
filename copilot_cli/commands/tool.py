"""Tool commands for managing all tool types available to Copilot Studio agents."""
import typer
from typing import Optional

from ..client import get_client
from ..output import print_json, print_table, handle_api_error

# Import subcommand modules
from . import prompt, restapi, connector, mcp


app = typer.Typer(help="Manage tools available to Copilot Studio agents")

# Register type-specific subcommands
app.add_typer(prompt.app, name="prompt", help="Manage AI Builder prompts")
app.add_typer(restapi.app, name="restapi", help="Manage REST API tools")
app.add_typer(connector.app, name="connector", help="Manage custom connectors")
app.add_typer(mcp.app, name="mcp", help="Manage MCP servers")


def format_unified_tool(tool: dict, tool_type: str) -> dict:
    """Format a tool for unified display."""
    if tool_type == "prompt":
        return {
            "name": tool.get("msdyn_name", ""),
            "type": "Prompt",
            "subtype": "System" if tool.get("ismanaged", False) else "Custom",
            "owner": tool.get("_ownerid_value@OData.Community.Display.V1.FormattedValue", ""),
            "id": tool.get("msdyn_aimodelid", ""),
        }
    elif tool_type == "restapi":
        return {
            "name": tool.get("displayname") or tool.get("name", ""),
            "type": "REST API",
            "subtype": "Custom",
            "owner": tool.get("_ownerid_value@OData.Community.Display.V1.FormattedValue", ""),
            "id": tool.get("connectorid", ""),
        }
    elif tool_type == "connector":
        props = tool.get("properties", {})
        # Only include custom connectors in unified view
        return {
            "name": props.get("displayName") or tool.get("name", ""),
            "type": "Connector",
            "subtype": "Custom",
            "owner": props.get("publisher", ""),
            "id": tool.get("name", ""),
        }
    elif tool_type == "mcp":
        props = tool.get("properties", {})
        is_custom = props.get("isCustomApi", False)
        return {
            "name": props.get("displayName") or tool.get("name", ""),
            "type": "MCP",
            "subtype": "Custom" if is_custom else "Managed",
            "owner": props.get("publisher", ""),
            "id": tool.get("name", ""),
        }
    return {}


@app.command("list")
def tool_list(
    tool_type: Optional[str] = typer.Option(
        None,
        "--type",
        "-T",
        help="Filter by tool type: prompt, restapi, connector, mcp",
    ),
    custom_only: bool = typer.Option(
        False,
        "--custom",
        "-c",
        help="Show only custom (user-created) tools",
    ),
    filter_text: Optional[str] = typer.Option(
        None,
        "--filter",
        "-f",
        help="Filter by name (case-insensitive)",
    ),
    table: bool = typer.Option(
        False,
        "--table",
        "-t",
        help="Display output as a formatted table instead of JSON",
    ),
):
    """
    List all tools available to Copilot Studio agents.

    Shows a unified view of all tool types: prompts, REST APIs, custom connectors,
    and MCP servers. Use --type to filter by specific tool type.

    Tool Types:
      - prompt: AI Builder prompts for text analysis and generation
      - restapi: REST API tools defined with OpenAPI specs
      - connector: Custom Power Platform connectors
      - mcp: Model Context Protocol servers

    Examples:
        copilot tool list                        # All tools
        copilot tool list --table                # All tools as table
        copilot tool list --type prompt          # Only prompts
        copilot tool list --type restapi --table # Only REST APIs
        copilot tool list --custom --table       # Only custom tools
        copilot tool list --filter "podio"       # Search by name
    """
    valid_types = ["prompt", "restapi", "connector", "mcp"]
    if tool_type and tool_type.lower() not in valid_types:
        typer.echo(f"Error: Invalid tool type '{tool_type}'. Must be one of: {', '.join(valid_types)}", err=True)
        raise typer.Exit(1)

    try:
        client = get_client()
        all_tools = []

        # Determine which types to fetch
        types_to_fetch = [tool_type.lower()] if tool_type else valid_types

        # Fetch each tool type
        if "prompt" in types_to_fetch:
            prompts = client.list_prompts()
            if custom_only:
                prompts = [p for p in prompts if not p.get("ismanaged", False)]
            for p in prompts:
                all_tools.append(format_unified_tool(p, "prompt"))

        if "restapi" in types_to_fetch:
            restapis = client.list_rest_apis()
            for r in restapis:
                all_tools.append(format_unified_tool(r, "restapi"))

        if "connector" in types_to_fetch:
            connectors = client.list_connectors()
            # Only include custom connectors in unified view
            from .connector import is_custom_connector
            connectors = [c for c in connectors if is_custom_connector(c)]
            for c in connectors:
                all_tools.append(format_unified_tool(c, "connector"))

        if "mcp" in types_to_fetch:
            mcps = client.list_mcp_servers()
            if custom_only:
                mcps = [m for m in mcps if m.get("properties", {}).get("isCustomApi", False)]
            for m in mcps:
                all_tools.append(format_unified_tool(m, "mcp"))

        if not all_tools:
            typer.echo("No tools found.")
            return

        # Apply text filter
        if filter_text:
            filter_lower = filter_text.lower()
            all_tools = [t for t in all_tools if filter_lower in t.get("name", "").lower()]

        if not all_tools:
            typer.echo("No tools match the filter criteria.")
            return

        # Sort by type then name
        type_order = {"Prompt": 0, "Connector": 1, "REST API": 2, "MCP": 3}
        all_tools.sort(key=lambda x: (type_order.get(x["type"], 99), x["name"].lower()))

        if table:
            print_table(
                all_tools,
                columns=["name", "type", "subtype", "owner", "id"],
                headers=["Name", "Type", "Subtype", "Owner", "ID"],
            )
        else:
            print_json(all_tools)

    except Exception as e:
        exit_code = handle_api_error(e)
        raise typer.Exit(exit_code)
