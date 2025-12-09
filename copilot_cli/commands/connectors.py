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


def extract_operations(
    connector: dict,
    include_deprecated: bool = False,
    include_internal: bool = False,
) -> list:
    """
    Extract operations (actions/triggers) from connector swagger definition.

    Args:
        connector: The connector definition with swagger
        include_deprecated: If True, include deprecated operations
        include_internal: If True, include internal-visibility operations
                         (internal ops cannot be used as Copilot agent tools)

    Returns:
        List of operation dicts with id, name, description, type, deprecated, visibility
    """
    operations = []
    swagger = connector.get("properties", {}).get("swagger", {})
    paths = swagger.get("paths", {})

    for path, methods in paths.items():
        for method, details in methods.items():
            if method in ["get", "post", "put", "patch", "delete"]:
                op_id = details.get("operationId")
                if not op_id:
                    continue

                is_deprecated = details.get("deprecated", False)
                visibility = details.get("x-ms-visibility", "normal")
                is_internal = visibility == "internal"

                # Skip deprecated unless explicitly requested
                if is_deprecated and not include_deprecated:
                    continue

                # Skip internal unless explicitly requested
                # Internal operations cannot be used as Copilot agent tools
                if is_internal and not include_internal:
                    continue

                # Determine if trigger or action
                is_trigger = details.get("x-ms-trigger") is not None
                op_type = "Trigger" if is_trigger else "Action"

                # Get description, truncate if too long
                description = details.get("description") or details.get("summary") or ""
                if len(description) > 80:
                    description = description[:77] + "..."

                operations.append({
                    "id": op_id,
                    "name": details.get("summary") or op_id,
                    "type": op_type,
                    "method": method.upper(),
                    "deprecated": is_deprecated,
                    "visibility": visibility,
                    "description": description,
                })

    # Sort: Actions first, then Triggers, then by name
    operations.sort(key=lambda x: (0 if x["type"] == "Action" else 1, x["id"].lower()))

    return operations


@app.command("get")
def connectors_get(
    connector_id: str = typer.Argument(
        ...,
        help="The connector's unique identifier (e.g., shared_asana, shared_office365)",
    ),
    table: bool = typer.Option(
        False,
        "--table",
        "-t",
        help="Display operations as a formatted table",
    ),
    include_deprecated: bool = typer.Option(
        False,
        "--include-deprecated",
        "-d",
        help="Include deprecated operations (hidden by default)",
    ),
    include_internal: bool = typer.Option(
        False,
        "--include-internal",
        "-i",
        help="Include internal operations (cannot be used as agent tools)",
    ),
    raw: bool = typer.Option(
        False,
        "--raw",
        "-r",
        help="Output raw JSON connector definition (ignores --table)",
    ),
):
    """
    Get details for a specific connector including available operations.

    By default, shows only operations that can be used as Copilot agent tools.
    Deprecated and internal-visibility operations are hidden by default.

    Examples:
        copilot connectors get shared_asana --table
        copilot connectors get shared_asana --table --include-deprecated
        copilot connectors get shared_asana --table --include-internal
        copilot connectors get shared_office365 --raw
    """
    try:
        client = get_client()
        connector = client.get_connector(connector_id)

        # Raw output - full JSON
        if raw:
            print_json(connector)
            return

        # Extract and display operations
        operations = extract_operations(connector, include_deprecated, include_internal)
        props = connector.get("properties", {})

        # Show connector summary
        typer.echo(f"\nConnector: {props.get('displayName', connector_id)}")
        typer.echo(f"ID: {connector.get('name', connector_id)}")
        typer.echo(f"Publisher: {props.get('publisher', 'N/A')}")

        if not operations:
            typer.echo("\nNo usable operations found.")
            typer.echo("Use --include-deprecated and/or --include-internal to see hidden operations.")
            return

        # Count hidden operations
        all_ops = extract_operations(connector, True, True)
        deprecated_count = len([o for o in all_ops if o["deprecated"]])
        internal_count = len([o for o in all_ops if o["visibility"] == "internal"])

        hidden_parts = []
        if deprecated_count > 0 and not include_deprecated:
            hidden_parts.append(f"{deprecated_count} deprecated")
        if internal_count > 0 and not include_internal:
            hidden_parts.append(f"{internal_count} internal")

        hidden_msg = f" ({', '.join(hidden_parts)} hidden)" if hidden_parts else ""
        typer.echo(f"\nOperations: {len(operations)}{hidden_msg}")

        if table:
            # Table format
            display_ops = []
            for op in operations:
                row = {
                    "id": op["id"],
                    "name": op["name"][:40] + "..." if len(op["name"]) > 40 else op["name"],
                    "type": op["type"],
                    "method": op["method"],
                }
                if include_deprecated:
                    row["deprecated"] = "Yes" if op["deprecated"] else "No"
                if include_internal:
                    row["visibility"] = op["visibility"]
                display_ops.append(row)

            columns = ["id", "name", "type", "method"]
            headers = ["Operation ID", "Name", "Type", "Method"]
            if include_deprecated:
                columns.append("deprecated")
                headers.append("Deprecated")
            if include_internal:
                columns.append("visibility")
                headers.append("Visibility")

            print_table(display_ops, columns=columns, headers=headers)
        else:
            # JSON format - just operations list
            print_json(operations)

    except Exception as e:
        exit_code = handle_api_error(e)
        raise typer.Exit(exit_code)
