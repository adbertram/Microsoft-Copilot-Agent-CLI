"""Prompt commands for managing AI Builder prompts available as agent tools."""
import typer
from typing import Optional

from ..client import get_client
from ..output import print_json, print_table, print_success, handle_api_error


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
    show_text: bool = typer.Option(
        False,
        "--text",
        "-t",
        help="Show the prompt text content instead of raw metadata",
    ),
):
    """
    Get details for a specific prompt.

    Examples:
        copilot prompt get 25583c46-ea44-4e47-8d83-a89bffb4ab27
        copilot prompt get 25583c46-ea44-4e47-8d83-a89bffb4ab27 --text
    """
    try:
        client = get_client()

        if show_text:
            # Get the prompt configuration with actual prompt text
            config = client.get_prompt_configuration(prompt_id)
            print_json(config)
        else:
            # Get the raw prompt metadata
            prompt = client.get_prompt(prompt_id)
            print_json(prompt)
    except Exception as e:
        exit_code = handle_api_error(e)
        raise typer.Exit(exit_code)


@app.command("update")
def prompt_update(
    prompt_id: str = typer.Argument(
        ...,
        help="The prompt's unique identifier (GUID)",
    ),
    text: Optional[str] = typer.Option(
        None,
        "--text",
        "-t",
        help="New prompt text (replaces existing prompt text)",
    ),
    text_file: Optional[str] = typer.Option(
        None,
        "--file",
        "-f",
        help="Read prompt text from a file",
    ),
    model: Optional[str] = typer.Option(
        None,
        "--model",
        "-m",
        help="Model type (e.g., gpt-41-mini, gpt-4o, gpt-4o-mini)",
    ),
    no_publish: bool = typer.Option(
        False,
        "--no-publish",
        help="Skip republishing after update (changes won't be live)",
    ),
):
    """
    Update an AI Builder prompt's text or model.

    The prompt text can be provided directly via --text or read from a file
    via --file. Input variables from the original prompt are preserved.

    By default, the prompt is automatically republished after updating.
    Use --no-publish to skip republishing.

    Examples:
        copilot prompt update <id> --text "Classify this content into categories..."
        copilot prompt update <id> --file prompt.txt
        copilot prompt update <id> --model gpt-4o
        copilot prompt update <id> --file prompt.txt --model gpt-4o
        copilot prompt update <id> --file prompt.txt --no-publish
    """
    # Validate input
    if text and text_file:
        typer.echo("Error: Cannot specify both --text and --file", err=True)
        raise typer.Exit(1)

    if not text and not text_file and not model:
        typer.echo("Error: Must provide --text, --file, or --model", err=True)
        raise typer.Exit(1)

    # Read prompt text from file if specified
    prompt_text = text
    if text_file:
        try:
            with open(text_file, "r") as f:
                prompt_text = f.read()
        except FileNotFoundError:
            typer.echo(f"Error: File not found: {text_file}", err=True)
            raise typer.Exit(1)
        except IOError as e:
            typer.echo(f"Error reading file: {e}", err=True)
            raise typer.Exit(1)

    try:
        client = get_client()

        # Get prompt name for confirmation message
        prompt_info = client.get_prompt(prompt_id)
        prompt_name = prompt_info.get("msdyn_name", prompt_id)

        typer.echo(f"Updating prompt '{prompt_name}'...")

        # Update the prompt (handles unpublish/update/republish workflow)
        client.update_prompt(
            prompt_id,
            prompt_text=prompt_text,
            model_type=model,
            publish=not no_publish
        )

        # Build update summary
        updates = []
        if prompt_text:
            updates.append("prompt text")
        if model:
            updates.append(f"model type ({model})")

        if no_publish:
            print_success(f"Updated {', '.join(updates)} for prompt '{prompt_name}' (not published)")
            typer.echo("\nUse AI Builder to publish when ready.")
        else:
            print_success(f"Updated and published {', '.join(updates)} for prompt '{prompt_name}'")

    except Exception as e:
        exit_code = handle_api_error(e)
        raise typer.Exit(exit_code)
