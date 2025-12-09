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
app.add_typer(connector.app, name="connector", help="Manage connectors")
app.add_typer(mcp.app, name="mcp", help="Manage MCP servers")


def format_unified_tool(tool: dict, tool_type: str) -> dict:
    """Format a tool for unified display with installed and managed status."""
    if tool_type == "prompt":
        is_managed = tool.get("ismanaged", False)
        return {
            "name": tool.get("msdyn_name", ""),
            "type": "Prompt",
            "publisher": tool.get("_ownerid_value@OData.Community.Display.V1.FormattedValue", ""),
            "installed": not is_managed,
            "managed": is_managed,
            "id": tool.get("msdyn_aimodelid", ""),
        }
    elif tool_type == "restapi":
        # REST APIs are always custom/installed
        return {
            "name": tool.get("displayname") or tool.get("name", ""),
            "type": "REST API",
            "publisher": tool.get("_ownerid_value@OData.Community.Display.V1.FormattedValue", ""),
            "installed": True,
            "managed": False,
            "id": tool.get("connectorid", ""),
        }
    elif tool_type == "connector":
        props = tool.get("properties", {})
        from .connector import is_custom_connector
        is_custom = is_custom_connector(tool)
        return {
            "name": props.get("displayName") or tool.get("name", ""),
            "type": "Connector",
            "publisher": props.get("publisher", ""),
            "installed": is_custom,
            "managed": not is_custom,
            "id": tool.get("name", ""),
        }
    elif tool_type == "mcp":
        props = tool.get("properties", {})
        is_custom = props.get("isCustomApi", False)
        return {
            "name": props.get("displayName") or tool.get("name", ""),
            "type": "MCP",
            "publisher": props.get("publisher", ""),
            "installed": is_custom,
            "managed": not is_custom,
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
    installed_only: bool = typer.Option(
        False,
        "--installed",
        "-i",
        help="Show only tools installed in your environment",
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
    List all tools available to add to Copilot Studio agents.

    Shows the full catalog of available tools: prompts, connectors, REST APIs,
    and MCP servers. The 'Installed' column shows which tools are configured
    in your environment.

    Tool Types:
      - prompt: AI Builder prompts for text analysis and generation
      - connector: Power Platform connectors (1000+ available)
      - restapi: REST API tools defined with OpenAPI specs
      - mcp: Model Context Protocol servers

    Installed Status:
      - Yes: Custom tool created in your environment
      - No: Available from catalog, not yet installed
      - System: Built-in system tool (always available)

    Examples:
        copilot tool list --table                  # All available tools
        copilot tool list --installed --table      # Only installed tools
        copilot tool list --type connector --table # All connectors
        copilot tool list --filter "excel" --table # Search by name
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
            for p in prompts:
                formatted = format_unified_tool(p, "prompt")
                if not installed_only or formatted["installed"]:
                    all_tools.append(formatted)

        if "restapi" in types_to_fetch:
            # REST APIs are always custom/installed
            restapis = client.list_rest_apis()
            for r in restapis:
                all_tools.append(format_unified_tool(r, "restapi"))

        if "connector" in types_to_fetch:
            # Show ALL connectors from catalog
            connectors = client.list_connectors()
            for c in connectors:
                formatted = format_unified_tool(c, "connector")
                if not installed_only or formatted["installed"]:
                    all_tools.append(formatted)

        if "mcp" in types_to_fetch:
            # Show ALL MCP servers from catalog
            mcps = client.list_mcp_servers()
            for m in mcps:
                formatted = format_unified_tool(m, "mcp")
                if not installed_only or formatted["installed"]:
                    all_tools.append(formatted)

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
                columns=["name", "type", "publisher", "installed", "managed", "id"],
                headers=["Name", "Type", "Publisher", "Installed", "Managed", "ID"],
            )
        else:
            print_json(all_tools)

    except Exception as e:
        exit_code = handle_api_error(e)
        raise typer.Exit(exit_code)
