"""Tool commands for managing all tool types available to Copilot Studio agents."""
import typer
from typing import Optional

from ..client import get_client
from ..output import print_json, print_table, print_success, handle_api_error

# Import subcommand modules
from . import prompt, restapi, connector, mcp


app = typer.Typer(help="Manage tools available to Copilot Studio agents")

# Register type-specific subcommands
app.add_typer(prompt.app, name="prompt", help="Manage AI Builder prompts")
app.add_typer(restapi.app, name="restapi", help="Manage REST API tools")
app.add_typer(connector.app, name="connector", help="Manage connectors")
app.add_typer(mcp.app, name="mcp", help="Manage MCP servers")


# Solution component types for dependency lookup
# See: https://learn.microsoft.com/en-us/power-apps/developer/data-platform/reference/entities/solutioncomponent
COMPONENT_TYPE_CONNECTOR = 372  # Custom Connector (connectors table)
COMPONENT_TYPE_AI_MODEL = None  # Will be looked up dynamically for msdyn_aimodel


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
            "_component_type": "msdyn_aimodel",  # Entity name for dynamic lookup
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
            "_component_type": COMPONENT_TYPE_CONNECTOR,  # Hardcoded - entity lookup doesn't work
        }
    elif tool_type == "connector":
        props = tool.get("properties", {})
        from .connector import is_custom_connector
        is_custom = is_custom_connector(tool)
        return {
            "name": props.get("displayName") or tool.get("name", ""),
            "type": "Custom Connector" if is_custom else "Connector",
            "publisher": props.get("publisher", ""),
            "installed": is_custom,
            "managed": not is_custom,
            "id": tool.get("name", ""),
            "_component_type": None,  # Power Apps connectors - no Dataverse dependency lookup
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
            "_component_type": None,  # Power Apps connectors - no Dataverse dependency lookup
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

        # Fetch dependencies for installed tools (only those with component type)
        for tool in all_tools:
            component_type = tool.get("_component_type")
            if tool.get("installed") and component_type is not None:
                try:
                    # Handle both integer component types and entity names
                    if isinstance(component_type, int):
                        deps = client.get_dependencies(tool["id"], component_type)
                    else:
                        deps = client.get_dependencies_for_entity(tool["id"], component_type)
                    tool["deps"] = len(deps)
                except Exception:
                    tool["deps"] = "-"
            else:
                tool["deps"] = "-"
            # Remove internal field
            tool.pop("_component_type", None)

        # Sort by type then name
        type_order = {"Prompt": 0, "Connector": 1, "Custom Connector": 2, "REST API": 3, "MCP": 4}
        all_tools.sort(key=lambda x: (type_order.get(x["type"], 99), x["name"].lower()))

        if table:
            print_table(
                all_tools,
                columns=["name", "type", "publisher", "installed", "deps", "id"],
                headers=["Name", "Type", "Publisher", "Installed", "Deps", "ID"],
            )
        else:
            print_json(all_tools)

    except Exception as e:
        exit_code = handle_api_error(e)
        raise typer.Exit(exit_code)


@app.command("remove")
def tool_remove(
    tool_id: str = typer.Argument(
        ...,
        help="The tool's unique identifier (GUID for prompts/REST APIs)",
    ),
    tool_type: Optional[str] = typer.Option(
        None,
        "--type",
        "-T",
        help="Tool type: prompt, restapi. Required if auto-detection fails.",
    ),
    force: bool = typer.Option(
        False,
        "--force",
        "-f",
        help="Skip confirmation prompt",
    ),
):
    """
    Remove a tool from the environment.

    Permanently deletes a custom tool (prompt or REST API) from the environment.
    This action cannot be undone.

    Supported tool types:
      - prompt: AI Builder prompts (custom only, not system prompts)
      - restapi: REST API tools (custom connectors)

    Note: Managed connectors from Microsoft cannot be deleted.

    Examples:
        copilot tool remove 25583c46-ea44-4e47-8d83-a89bffb4ab27 --type prompt
        copilot tool remove 56c1700d-a317-4472-8bd6-928afa5be754 --type restapi
        copilot tool remove <id> --type prompt --force
    """
    valid_types = ["prompt", "restapi"]
    if tool_type and tool_type.lower() not in valid_types:
        typer.echo(f"Error: Invalid tool type '{tool_type}'. Must be one of: {', '.join(valid_types)}", err=True)
        raise typer.Exit(1)

    try:
        client = get_client()
        tool_name = None
        detected_type = tool_type.lower() if tool_type else None

        # Try to auto-detect tool type if not specified
        if not detected_type:
            # Try prompt first
            try:
                tool_info = client.get_prompt(tool_id)
                detected_type = "prompt"
                tool_name = tool_info.get("msdyn_name", tool_id)
                # Check if managed
                if tool_info.get("ismanaged", False):
                    typer.echo("Error: Cannot delete system/managed prompts.", err=True)
                    raise typer.Exit(1)
            except Exception:
                pass

            # Try REST API
            if not detected_type:
                try:
                    tool_info = client.get_rest_api(tool_id)
                    detected_type = "restapi"
                    tool_name = tool_info.get("displayname") or tool_info.get("name", tool_id)
                except Exception:
                    pass

            if not detected_type:
                typer.echo(
                    f"Error: Could not find tool with ID '{tool_id}'. "
                    "Please specify --type to indicate the tool type.",
                    err=True
                )
                raise typer.Exit(1)
        else:
            # Get tool info for confirmation message
            if detected_type == "prompt":
                tool_info = client.get_prompt(tool_id)
                tool_name = tool_info.get("msdyn_name", tool_id)
                if tool_info.get("ismanaged", False):
                    typer.echo("Error: Cannot delete system/managed prompts.", err=True)
                    raise typer.Exit(1)
            elif detected_type == "restapi":
                tool_info = client.get_rest_api(tool_id)
                tool_name = tool_info.get("displayname") or tool_info.get("name", tool_id)

        # Confirm deletion
        type_display = "Prompt" if detected_type == "prompt" else "REST API"
        if not force:
            typer.confirm(
                f"Are you sure you want to delete {type_display} '{tool_name}'? This cannot be undone.",
                abort=True,
            )

        # Delete the tool
        # Use component type for REST APIs (372), entity name for prompts
        component_type = COMPONENT_TYPE_CONNECTOR if detected_type == "restapi" else "msdyn_aimodel"
        try:
            if detected_type == "prompt":
                client.delete_prompt(tool_id)
            elif detected_type == "restapi":
                client.delete_rest_api(tool_id)
            print_success(f"Deleted {type_display} '{tool_name}'")
        except Exception as delete_error:
            error_msg = str(delete_error)
            # Check if it's a dependency error
            if "referenced by" in error_msg.lower() or "cannot be deleted" in error_msg.lower():
                typer.echo(f"Error: Cannot delete {type_display} '{tool_name}' - it has dependencies.\n", err=True)
                # Fetch and display dependencies
                try:
                    if isinstance(component_type, int):
                        deps = client.get_dependencies(tool_id, component_type)
                    else:
                        deps = client.get_dependencies_for_entity(tool_id, component_type)
                    if deps:
                        typer.echo("Dependent components:", err=True)
                        for dep in deps:
                            dep_type_code = dep.get("dependentcomponenttype")
                            dep_type = dep.get("dependentcomponenttype@OData.Community.Display.V1.FormattedValue")
                            dep_id = dep.get("dependentcomponentobjectid", "")
                            # If formatted value is None, look up the component type name
                            if not dep_type or dep_type == "None":
                                try:
                                    type_info = client.get(f"solutioncomponentdefinitions?$filter=solutioncomponenttype eq {dep_type_code}&$select=name")
                                    type_defs = type_info.get("value", [])
                                    dep_type = type_defs[0].get("name") if type_defs else f"Type {dep_type_code}"
                                except Exception:
                                    dep_type = f"Type {dep_type_code}"
                            typer.echo(f"  - {dep_type}: {dep_id}", err=True)
                        typer.echo("\nRemove this tool from these components before deleting.", err=True)
                except Exception:
                    pass  # If dependency lookup fails, just show original error
                raise typer.Exit(1)
            else:
                raise delete_error

    except typer.Abort:
        typer.echo("Aborted.")
        raise typer.Exit(0)
    except Exception as e:
        exit_code = handle_api_error(e)
        raise typer.Exit(exit_code)
