"""Prompt commands for listing AI Builder prompts available as agent tools."""
import typer
from typing import Optional

from ..client import get_client
from ..output import print_json, print_table, handle_api_error


app = typer.Typer(help="Manage AI Builder prompts (available as agent tools)")


# GptPowerPrompt template ID - identifies AI Builder prompts
GPT_POWER_PROMPT_TEMPLATE_ID = "edfdb190-3791-45d8-9a6c-8f90a37c278a"


def format_prompt_for_display(prompt: dict) -> dict:
    """Format a prompt for display."""
    name = prompt.get("msdyn_name", "")
    prompt_id = prompt.get("msdyn_aimodelid", "")

    # Determine type (Custom vs System based on ismanaged)
    is_managed = prompt.get("ismanaged", False)
    prompt_type = "System" if is_managed else "Custom"

    # Get state
    state_code = prompt.get("statecode", 0)
    state = "Active" if state_code == 1 else "Inactive"

    # Get owner
    owner = prompt.get("_ownerid_value@OData.Community.Display.V1.FormattedValue", "")

    # Get created/modified dates
    created = prompt.get("createdon", "")
    if created:
        created = created.split("T")[0]  # Just the date part

    modified = prompt.get("modifiedon", "")
    if modified:
        modified = modified.split("T")[0]

    return {
        "name": name,
        "type": prompt_type,
        "id": prompt_id,
        "state": state,
        "owner": owner,
        "created": created,
        "modified": modified,
    }


@app.command("list")
def prompt_list(
    custom: bool = typer.Option(
        False,
        "--custom",
        "-c",
        help="Show only custom (user-created) prompts",
    ),
    system: bool = typer.Option(
        False,
        "--system",
        "-s",
        help="Show only system (managed) prompts",
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
    List all AI Builder prompts available as agent tools.

    AI Builder prompts are custom prompts that can be attached to
    Copilot Studio agents as tools. They use GPT models to perform
    specific tasks like classification, extraction, or content generation.

    Prompt Types:
      - Custom: User-created prompts in the environment
      - System: Built-in Microsoft prompts (AI Classify, AI Summarize, etc.)

    Examples:
        copilot prompt list
        copilot prompt list --table
        copilot prompt list --custom --table
        copilot prompt list --system --table
        copilot prompt list --filter "classify" --table
    """
    if custom and system:
        typer.echo("Error: Cannot specify both --custom and --system", err=True)
        raise typer.Exit(1)

    try:
        client = get_client()
        prompts = client.list_prompts()

        if not prompts:
            typer.echo("No prompts found.")
            return

        # Filter by custom/system
        if custom:
            prompts = [p for p in prompts if not p.get("ismanaged", False)]
        elif system:
            prompts = [p for p in prompts if p.get("ismanaged", False)]

        # Filter by text
        if filter_text:
            filter_lower = filter_text.lower()
            prompts = [
                p for p in prompts
                if filter_lower in p.get("msdyn_name", "").lower()
            ]

        if not prompts:
            typer.echo("No prompts match the filter criteria.")
            return

        formatted = [format_prompt_for_display(p) for p in prompts]

        # Sort by name
        formatted.sort(key=lambda x: x["name"].lower())

        if table:
            print_table(
                formatted,
                columns=["name", "type", "state", "owner", "modified", "id"],
                headers=["Name", "Type", "State", "Owner", "Modified", "ID"],
            )
        else:
            print_json(formatted)
    except Exception as e:
        exit_code = handle_api_error(e)
        raise typer.Exit(exit_code)


@app.command("get")
def prompt_get(
    prompt_id: str = typer.Argument(
        ...,
        help="The prompt's unique identifier (GUID)",
    ),
):
    """
    Get details for a specific prompt.

    Examples:
        copilot prompt get 25583c46-ea44-4e47-8d83-a89bffb4ab27
    """
    try:
        client = get_client()
        prompt = client.get_prompt(prompt_id)
        print_json(prompt)
    except Exception as e:
        exit_code = handle_api_error(e)
        raise typer.Exit(exit_code)
