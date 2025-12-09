"""Connector commands for listing available Power Platform connectors."""
import typer
from typing import Optional

from ..client import get_client
from ..output import print_json, print_table, handle_api_error


app = typer.Typer(help="List and inspect Power Platform connectors")


def is_custom_connector(connector: dict) -> bool:
    """
    Determine if a connector is custom or managed.

    Custom connectors have different properties structure than managed connectors.
    """
    props = connector.get("properties", {})

    # Check for custom connector indicators
    if "environment" in props:
        return True

    # Check publisher - custom connectors often have user/org as publisher
    publisher = (props.get("publisher") or "").lower()
    if publisher and publisher not in ["microsoft", "microsoft corporation", "azure"]:
        # Could be custom, but also third-party managed
        # Check tier - custom connectors typically don't have tier or have "NotSpecified"
        tier = props.get("tier", "")
        if not tier or tier == "NotSpecified":
            return True

    return False


def format_connector_for_display(connector: dict) -> dict:
    """Format a connector for display."""
    props = connector.get("properties", {})

    description = props.get("description") or ""
    if len(description) > 60:
        description = description[:57] + "..."

    is_custom = is_custom_connector(connector)

    return {
        "name": props.get("displayName") or connector.get("name", ""),
        "id": connector.get("name", ""),
        "type": "Custom" if is_custom else "Managed",
        "publisher": props.get("publisher") or "",
        "tier": props.get("tier") or "N/A",
        "description": description,
    }


@app.command("list")
def connectors_list(
    custom: bool = typer.Option(
        False,
        "--custom",
        "-c",
        help="Show only custom connectors",
    ),
    managed: bool = typer.Option(
        False,
        "--managed",
        "-m",
        help="Show only managed (Microsoft) connectors",
    ),
    filter_text: Optional[str] = typer.Option(
        None,
        "--filter",
        "-f",
        help="Filter by name or publisher (case-insensitive)",
    ),
    table: bool = typer.Option(
        False,
        "--table",
        "-t",
        help="Display output as a formatted table instead of JSON",
    ),
):
    """
    List all available connectors in the environment.

    Connectors are proxies/wrappers around APIs that define what actions
    are available (e.g., Asana, SharePoint, SQL Server). They represent
    the "type" of service you can connect to.

    Connector Types:
      - Managed: Built-in connectors published by Microsoft
      - Custom: User-created connectors in the environment

    Examples:
        copilot connectors list
        copilot connectors list --table
        copilot connectors list --custom --table
        copilot connectors list --managed --table
        copilot connectors list --filter "asana" --table
    """
    if custom and managed:
        typer.echo("Error: Cannot specify both --custom and --managed", err=True)
        raise typer.Exit(1)

    try:
        client = get_client()
        connectors = client.list_connectors()

        if not connectors:
            typer.echo("No connectors found.")
            return

        # Filter by custom/managed
        if custom:
            connectors = [c for c in connectors if is_custom_connector(c)]
        elif managed:
            connectors = [c for c in connectors if not is_custom_connector(c)]

        # Filter by text
        if filter_text:
            filter_lower = filter_text.lower()
            connectors = [
                c for c in connectors
                if filter_lower in c.get("properties", {}).get("displayName", "").lower()
                or filter_lower in c.get("properties", {}).get("publisher", "").lower()
                or filter_lower in c.get("name", "").lower()
            ]

        if not connectors:
            typer.echo("No connectors match the filter criteria.")
            return

        formatted = [format_connector_for_display(c) for c in connectors]

        # Sort by type (Custom first) then name
        formatted.sort(key=lambda x: (0 if x["type"] == "Custom" else 1, x["name"].lower()))

        if table:
            print_table(
                formatted,
                columns=["name", "type", "publisher", "tier", "id"],
                headers=["Name", "Type", "Publisher", "Tier", "ID"],
            )
        else:
            print_json(formatted)
    except Exception as e:
        exit_code = handle_api_error(e)
        raise typer.Exit(exit_code)


@app.command("get")
def connectors_get(
    connector_id: str = typer.Argument(
        ...,
        help="The connector's unique identifier (e.g., shared_asana, shared_office365)",
    ),
):
    """
    Get details for a specific connector.

    Returns the full connector definition including available actions,
    triggers, and connection parameters.

    Examples:
        copilot connectors get shared_asana
        copilot connectors get shared_office365
        copilot connectors get shared_sharepointonline
    """
    try:
        client = get_client()
        connector = client.get_connector(connector_id)
        print_json(connector)
    except Exception as e:
        exit_code = handle_api_error(e)
        raise typer.Exit(exit_code)
