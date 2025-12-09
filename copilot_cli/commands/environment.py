"""Environment commands for managing Power Platform environments."""
import typer
from typing import Optional

from ..client import get_client
from ..output import print_json, print_table, handle_api_error


app = typer.Typer(help="Manage Power Platform environments")


def format_environment_for_display(env: dict) -> dict:
    """Format an environment for display."""
    props = env.get("properties", {})

    # Get environment type
    env_type = props.get("environmentSku", "")

    # Get state
    states = props.get("states", {})
    runtime_state = states.get("runtime", {}).get("id", "")

    # Get region
    azure_region = props.get("azureRegion", "")

    # Get created time
    created = props.get("createdTime", "")
    if created:
        created = created.split("T")[0]

    # Get linked environment (Dataverse)
    linked_env = props.get("linkedEnvironmentMetadata", {})
    dataverse_url = linked_env.get("instanceUrl", "")

    # Check if default
    is_default = props.get("isDefault", False)

    return {
        "name": props.get("displayName", ""),
        "id": env.get("name", ""),
        "type": env_type,
        "region": azure_region,
        "state": runtime_state,
        "default": is_default,
        "dataverse_url": dataverse_url,
        "created": created,
    }


@app.command("list")
def environment_list(
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
    List all Power Platform environments accessible to you.

    Shows all environments in your tenant including production, sandbox,
    developer, and trial environments.

    Examples:
        copilot environment list
        copilot environment list --table
        copilot environment list --filter "dev" --table
    """
    try:
        client = get_client()
        environments = client.list_environments()

        if not environments:
            typer.echo("No environments found.")
            return

        # Apply text filter
        if filter_text:
            filter_lower = filter_text.lower()
            environments = [
                e for e in environments
                if filter_lower in e.get("properties", {}).get("displayName", "").lower()
                or filter_lower in e.get("name", "").lower()
            ]

        if not environments:
            typer.echo("No environments match the filter criteria.")
            return

        formatted = [format_environment_for_display(e) for e in environments]

        # Sort by default first, then name
        formatted.sort(key=lambda x: (not x["default"], x["name"].lower()))

        if table:
            print_table(
                formatted,
                columns=["name", "type", "region", "state", "default", "id"],
                headers=["Name", "Type", "Region", "State", "Default", "ID"],
            )
        else:
            print_json(formatted)

    except Exception as e:
        exit_code = handle_api_error(e)
        raise typer.Exit(exit_code)


@app.command("get")
def environment_get(
    environment_id: str = typer.Argument(
        ...,
        help="The environment's unique identifier (e.g., Default-<tenant-id> or GUID)",
    ),
):
    """
    Get details for a specific environment.

    Examples:
        copilot environment get Default-12345678-1234-1234-1234-123456789012
        copilot environment get 11376bd0-c80f-4e99-b86f-05d17b73518d
    """
    try:
        client = get_client()
        environment = client.get_environment(environment_id)
        print_json(environment)

    except Exception as e:
        exit_code = handle_api_error(e)
        raise typer.Exit(exit_code)
