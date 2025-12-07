"""Flow commands for listing available Power Automate flows."""
import typer
from typing import Optional

from ..client import get_client
from ..output import print_json, print_table, handle_api_error


app = typer.Typer(help="Manage Power Automate flows")


# Flow category mappings
FLOW_CATEGORIES = {
    0: "Automated",
    1: "Scheduled",
    2: "Button",
    3: "Approval",
    5: "Instant",
    6: "Business Process",
}


def get_category_name(category: int) -> str:
    """Get human-readable category name."""
    return FLOW_CATEGORIES.get(category, f"Category {category}")


def format_flow_for_display(flow: dict) -> dict:
    """Format a flow for display."""
    description = flow.get("description") or ""
    if len(description) > 80:
        description = description[:77] + "..."

    category = flow.get("category", 0)

    return {
        "name": flow.get("name"),
        "id": flow.get("workflowid"),
        "category": get_category_name(category),
        "description": description,
        "status": flow.get("statecode@OData.Community.Display.V1.FormattedValue", "Active"),
    }


@app.command("list")
def flow_list(
    category: Optional[int] = typer.Option(
        None,
        "--category",
        "-c",
        help="Filter by category: 0=Automated, 5=Instant, 6=Business Process",
    ),
    table: bool = typer.Option(
        False,
        "--table",
        "-t",
        help="Display output as a formatted table instead of JSON",
    ),
):
    """
    List Power Automate cloud flows in the environment.

    This command lists flows stored in Dataverse that can potentially
    be used as tools in Copilot Studio agents.

    Categories:
      - 0: Automated (automated/scheduled flows)
      - 5: Instant (button/HTTP triggered flows)
      - 6: Business Process flows

    Note: Flows that work best as agent tools are typically Instant (5)
    flows with HTTP request triggers.

    Examples:
        copilot flow list
        copilot flow list --table
        copilot flow list --category 5
        copilot flow list --category 5 --table
    """
    try:
        client = get_client()
        flows = client.list_flows(category=category)

        if not flows:
            typer.echo("No flows found.")
            return

        formatted = [format_flow_for_display(f) for f in flows]

        if table:
            print_table(
                formatted,
                columns=["name", "category", "status", "id"],
                headers=["Name", "Category", "Status", "ID"],
            )
        else:
            print_json(formatted)
    except Exception as e:
        exit_code = handle_api_error(e)
        raise typer.Exit(exit_code)


@app.command("get")
def flow_get(
    workflow_id: str = typer.Argument(
        ...,
        help="The flow's unique identifier (GUID)",
    ),
):
    """
    Get details for a specific flow.

    Examples:
        copilot flow get <flow-id>
    """
    try:
        client = get_client()
        flow = client.get_flow(workflow_id)
        print_json(flow)
    except Exception as e:
        exit_code = handle_api_error(e)
        raise typer.Exit(exit_code)
