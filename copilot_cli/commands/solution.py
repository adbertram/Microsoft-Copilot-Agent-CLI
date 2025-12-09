"""Solution commands for Copilot CLI."""
import typer
from typing import Optional

from ..client import get_client
from ..output import print_json, print_table, print_success, handle_api_error

app = typer.Typer(help="Manage solutions and solution components")


def format_solution_for_display(solution: dict) -> dict:
    """Format a solution record for display."""
    return {
        "friendlyname": solution.get("friendlyname", ""),
        "uniquename": solution.get("uniquename", ""),
        "solutionid": solution.get("solutionid", ""),
        "version": solution.get("version", ""),
        "ismanaged": "Yes" if solution.get("ismanaged") else "No",
    }


@app.command("list")
def list_solutions(
    table: bool = typer.Option(
        False,
        "--table",
        "-t",
        help="Display output as a formatted table instead of JSON",
    ),
    all_solutions: bool = typer.Option(
        False,
        "--all",
        "-a",
        help="Include managed solutions (by default, only unmanaged are shown)",
    ),
):
    """
    List solutions in the environment.

    By default, only unmanaged solutions are shown (these are the solutions you can modify).

    Examples:
        copilot solution list
        copilot solution list --table
        copilot solution list --all
    """
    try:
        client = get_client()

        if all_solutions:
            # Get all solutions including managed
            result = client.get("solutions?$orderby=friendlyname")
            solutions = result.get("value", [])
        else:
            solutions = client.list_solutions()

        if table:
            formatted = [format_solution_for_display(s) for s in solutions]
            print_table(
                formatted,
                columns=["friendlyname", "uniquename", "version", "ismanaged"],
                headers=["Name", "Unique Name", "Version", "Managed"],
            )
        else:
            formatted = [format_solution_for_display(s) for s in solutions]
            print_json(formatted)
    except Exception as e:
        exit_code = handle_api_error(e)
        raise typer.Exit(exit_code)


@app.command("get")
def get_solution(
    solution: str = typer.Argument(
        ...,
        help="The solution's unique name or GUID",
    ),
):
    """
    Get details for a specific solution.

    Examples:
        copilot solution get MySolution
        copilot solution get 12345678-1234-1234-1234-123456789abc
    """
    try:
        client = get_client()
        result = client.get_solution(solution)
        print_json(result)
    except Exception as e:
        exit_code = handle_api_error(e)
        raise typer.Exit(exit_code)


@app.command("create")
def create_solution(
    name: str = typer.Option(
        ...,
        "--name",
        "-n",
        help="The display name for the solution",
    ),
    unique_name: str = typer.Option(
        ...,
        "--unique-name",
        "-u",
        help="The unique name for the solution (no spaces, used for identification)",
    ),
    publisher: str = typer.Option(
        ...,
        "--publisher",
        "-p",
        help="The publisher's unique name or GUID",
    ),
    version: str = typer.Option(
        "1.0.0.0",
        "--version",
        "-v",
        help="The solution version (format: major.minor.build.revision)",
    ),
    description: Optional[str] = typer.Option(
        None,
        "--description",
        "-d",
        help="Optional description for the solution",
    ),
):
    """
    Create a new unmanaged solution.

    A publisher must exist before creating a solution. Use 'copilot solution publisher list'
    to see available publishers.

    Examples:
        copilot solution create --name "My Solution" --unique-name MySolution --publisher MyPublisher
        copilot solution create -n "My Solution" -u MySolution -p MyPublisher -v 1.0.0.0
        copilot solution create -n "My Solution" -u MySolution -p MyPublisher -d "Description here"
    """
    try:
        client = get_client()

        client.create_solution(
            unique_name=unique_name,
            friendly_name=name,
            publisher_id=publisher,
            version=version,
            description=description,
        )

        print_success(f"Solution '{name}' created successfully.")

        # Fetch the created solution to display its details
        created_solution = client.get_solution(unique_name)
        print_json(format_solution_for_display(created_solution))

    except Exception as e:
        exit_code = handle_api_error(e)
        raise typer.Exit(exit_code)


@app.command("delete")
def delete_solution(
    solution: str = typer.Argument(
        ...,
        help="The solution's unique name or GUID",
    ),
    force: bool = typer.Option(
        False,
        "--force",
        "-f",
        help="Skip confirmation prompt",
    ),
):
    """
    Delete an unmanaged solution.

    This deletes the solution container but does NOT delete the components within it.
    Components will remain in the environment.

    Examples:
        copilot solution delete MySolution
        copilot solution delete MySolution --force
    """
    try:
        client = get_client()

        # Get solution details for display
        solution_details = client.get_solution(solution)
        solution_name = solution_details.get("friendlyname", solution)

        if solution_details.get("ismanaged"):
            typer.echo("Error: Cannot delete managed solutions.", err=True)
            raise typer.Exit(1)

        if not force:
            confirm = typer.confirm(
                f"Delete solution '{solution_name}'? Components will remain in the environment."
            )
            if not confirm:
                typer.echo("Aborted.")
                raise typer.Exit(0)

        client.delete_solution(solution)
        print_success(f"Solution '{solution_name}' deleted.")

    except Exception as e:
        exit_code = handle_api_error(e)
        raise typer.Exit(exit_code)


@app.command("add-agent")
def add_agent_to_solution(
    solution: str = typer.Option(
        ...,
        "--solution",
        "-s",
        help="The solution's unique name",
    ),
    agent_id: str = typer.Option(
        ...,
        "--agent",
        "-a",
        help="The agent's unique identifier (GUID)",
    ),
    include_connection: bool = typer.Option(
        True,
        "--include-connection/--no-connection",
        help="Also add the agent's connection reference to the solution (default: True)",
    ),
    add_required: bool = typer.Option(
        True,
        "--add-required/--no-required",
        help="Add required dependent components (default: True)",
    ),
):
    """
    Add a Copilot agent (and optionally its connection reference) to a solution.

    This command adds the specified agent to an unmanaged solution. By default,
    it also adds the agent's connection reference for knowledge sources.

    Examples:
        copilot solution add-agent --solution MySolution --agent <agent-id>
        copilot solution add-agent -s MySolution -a <agent-id> --no-connection
        copilot solution add-agent -s MySolution -a <agent-id> --no-required
    """
    try:
        client = get_client()

        # Get agent name for display
        bot = client.get_bot(agent_id)
        agent_name = bot.get("name", agent_id)

        # Add the agent to the solution
        client.add_bot_to_solution(
            solution_unique_name=solution,
            bot_id=agent_id,
            add_required_components=add_required,
        )
        print_success(f"Agent '{agent_name}' added to solution '{solution}'.")

        # Optionally add connection reference
        if include_connection:
            provider_ref_id = bot.get("_providerconnectionreferenceid_value")
            if provider_ref_id:
                try:
                    client.add_connection_reference_to_solution(
                        solution_unique_name=solution,
                        connection_reference_id=provider_ref_id,
                        add_required_components=False,
                    )
                    print_success(f"Connection reference added to solution '{solution}'.")
                except Exception as conn_error:
                    # Connection reference might already be in the solution
                    error_str = str(conn_error).lower()
                    if "already exists" in error_str or "duplicate" in error_str:
                        typer.echo("Connection reference already exists in solution.", err=True)
                    else:
                        typer.echo(f"Warning: Could not add connection reference: {conn_error}", err=True)
            else:
                typer.echo("Note: Agent has no connection reference configured.", err=True)

    except Exception as e:
        exit_code = handle_api_error(e)
        raise typer.Exit(exit_code)


@app.command("remove-agent")
def remove_agent_from_solution(
    solution: str = typer.Option(
        ...,
        "--solution",
        "-s",
        help="The solution's unique name",
    ),
    agent_id: str = typer.Option(
        ...,
        "--agent",
        "-a",
        help="The agent's unique identifier (GUID)",
    ),
    force: bool = typer.Option(
        False,
        "--force",
        "-f",
        help="Skip confirmation prompt",
    ),
):
    """
    Remove a Copilot agent from a solution.

    This removes the agent from the solution but does NOT delete the agent itself.

    Examples:
        copilot solution remove-agent --solution MySolution --agent <agent-id>
        copilot solution remove-agent -s MySolution -a <agent-id> --force
    """
    try:
        client = get_client()

        # Get agent name for display
        bot = client.get_bot(agent_id)
        agent_name = bot.get("name", agent_id)

        if not force:
            confirm = typer.confirm(
                f"Remove agent '{agent_name}' from solution '{solution}'?"
            )
            if not confirm:
                typer.echo("Aborted.")
                raise typer.Exit(0)

        client.remove_bot_from_solution(
            solution_unique_name=solution,
            bot_id=agent_id,
        )
        print_success(f"Agent '{agent_name}' removed from solution '{solution}'.")

    except Exception as e:
        exit_code = handle_api_error(e)
        raise typer.Exit(exit_code)


@app.command("add-connection")
def add_connection_to_solution(
    solution: str = typer.Option(
        ...,
        "--solution",
        "-s",
        help="The solution's unique name",
    ),
    connection_id: str = typer.Option(
        ...,
        "--connection",
        "-c",
        help="The connection reference's unique identifier (GUID)",
    ),
):
    """
    Add a connection reference to a solution.

    Examples:
        copilot solution add-connection --solution MySolution --connection <connection-id>
        copilot solution add-connection -s MySolution -c <connection-id>
    """
    try:
        client = get_client()

        client.add_connection_reference_to_solution(
            solution_unique_name=solution,
            connection_reference_id=connection_id,
            add_required_components=False,
        )
        print_success(f"Connection reference added to solution '{solution}'.")

    except Exception as e:
        exit_code = handle_api_error(e)
        raise typer.Exit(exit_code)


@app.command("remove-connection")
def remove_connection_from_solution(
    solution: str = typer.Option(
        ...,
        "--solution",
        "-s",
        help="The solution's unique name",
    ),
    connection_id: str = typer.Option(
        ...,
        "--connection",
        "-c",
        help="The connection reference's unique identifier (GUID)",
    ),
    force: bool = typer.Option(
        False,
        "--force",
        "-f",
        help="Skip confirmation prompt",
    ),
):
    """
    Remove a connection reference from a solution.

    This removes the connection reference from the solution but does NOT delete it.

    Examples:
        copilot solution remove-connection --solution MySolution --connection <connection-id>
        copilot solution remove-connection -s MySolution -c <connection-id> --force
    """
    try:
        client = get_client()

        if not force:
            confirm = typer.confirm(
                f"Remove connection reference from solution '{solution}'?"
            )
            if not confirm:
                typer.echo("Aborted.")
                raise typer.Exit(0)

        client.remove_connection_reference_from_solution(
            solution_unique_name=solution,
            connection_reference_id=connection_id,
        )
        print_success(f"Connection reference removed from solution '{solution}'.")

    except Exception as e:
        exit_code = handle_api_error(e)
        raise typer.Exit(exit_code)


# Publisher commands
publisher_app = typer.Typer(help="Manage publishers")


def format_publisher_for_display(publisher: dict) -> dict:
    """Format a publisher record for display."""
    return {
        "friendlyname": publisher.get("friendlyname", ""),
        "uniquename": publisher.get("uniquename", ""),
        "publisherid": publisher.get("publisherid", ""),
        "customizationprefix": publisher.get("customizationprefix", ""),
    }


@publisher_app.command("list")
def list_publishers(
    table: bool = typer.Option(
        False,
        "--table",
        "-t",
        help="Display output as a formatted table instead of JSON",
    ),
):
    """
    List publishers in the environment.

    Publishers are required when creating solutions. Each solution must be linked
    to a publisher.

    Examples:
        copilot solution publisher list
        copilot solution publisher list --table
    """
    try:
        client = get_client()
        publishers = client.list_publishers()

        if not publishers:
            typer.echo("No publishers found.")
            return

        if table:
            formatted = [format_publisher_for_display(p) for p in publishers]
            print_table(
                formatted,
                columns=["friendlyname", "uniquename", "customizationprefix"],
                headers=["Name", "Unique Name", "Prefix"],
            )
        else:
            formatted = [format_publisher_for_display(p) for p in publishers]
            print_json(formatted)

    except Exception as e:
        exit_code = handle_api_error(e)
        raise typer.Exit(exit_code)


@publisher_app.command("get")
def get_publisher(
    publisher: str = typer.Argument(
        ...,
        help="The publisher's unique name or GUID",
    ),
):
    """
    Get details for a specific publisher.

    Examples:
        copilot solution publisher get MyPublisher
        copilot solution publisher get 12345678-1234-1234-1234-123456789abc
    """
    try:
        client = get_client()
        result = client.get_publisher(publisher)
        print_json(result)
    except Exception as e:
        exit_code = handle_api_error(e)
        raise typer.Exit(exit_code)


@publisher_app.command("create")
def create_publisher(
    name: str = typer.Option(
        ...,
        "--name",
        "-n",
        help="The display name for the publisher",
    ),
    unique_name: str = typer.Option(
        ...,
        "--unique-name",
        "-u",
        help="The unique name for the publisher (no spaces)",
    ),
    prefix: str = typer.Option(
        ...,
        "--prefix",
        "-x",
        help="Customization prefix (2-8 lowercase letters, used for solution components)",
    ),
    option_value_prefix: int = typer.Option(
        ...,
        "--option-prefix",
        "-o",
        help="Option value prefix (10000-99999, used for choice option values)",
    ),
    description: Optional[str] = typer.Option(
        None,
        "--description",
        "-d",
        help="Optional description for the publisher",
    ),
):
    """
    Create a new publisher.

    Publishers are required for creating solutions. The customization prefix is used
    to prefix schema names of solution components.

    Examples:
        copilot solution publisher create --name "My Publisher" --unique-name MyPublisher --prefix mypub --option-prefix 10000
        copilot solution publisher create -n "My Publisher" -u MyPublisher -x mypub -o 10000 -d "My description"
    """
    try:
        client = get_client()

        client.create_publisher(
            unique_name=unique_name,
            friendly_name=name,
            customization_prefix=prefix,
            customization_option_value_prefix=option_value_prefix,
            description=description,
        )

        print_success(f"Publisher '{name}' created successfully.")

        # Fetch the created publisher to display its details
        created_publisher = client.get_publisher(unique_name)
        print_json(format_publisher_for_display(created_publisher))

    except Exception as e:
        exit_code = handle_api_error(e)
        raise typer.Exit(exit_code)


@publisher_app.command("delete")
def delete_publisher(
    publisher: str = typer.Argument(
        ...,
        help="The publisher's unique name or GUID",
    ),
    force: bool = typer.Option(
        False,
        "--force",
        "-f",
        help="Skip confirmation prompt",
    ),
):
    """
    Delete a publisher.

    Publishers cannot be deleted if they have solutions associated with them.
    You must delete all solutions using the publisher first.

    Examples:
        copilot solution publisher delete MyPublisher
        copilot solution publisher delete MyPublisher --force
    """
    try:
        client = get_client()

        # Get publisher details for display
        publisher_details = client.get_publisher(publisher)
        publisher_name = publisher_details.get("friendlyname", publisher)

        if not force:
            confirm = typer.confirm(
                f"Delete publisher '{publisher_name}'? This cannot be undone."
            )
            if not confirm:
                typer.echo("Aborted.")
                raise typer.Exit(0)

        client.delete_publisher(publisher)
        print_success(f"Publisher '{publisher_name}' deleted.")

    except Exception as e:
        exit_code = handle_api_error(e)
        raise typer.Exit(exit_code)


# Connection reference listing commands
connection_app = typer.Typer(help="Manage connection references")


@connection_app.command("list")
def list_connections(
    agent_id: Optional[str] = typer.Option(
        None,
        "--agent",
        "-a",
        help="Filter to show only the connection reference for a specific agent",
    ),
    table: bool = typer.Option(
        False,
        "--table",
        "-t",
        help="Display output as a formatted table instead of JSON",
    ),
):
    """
    List connection references in the environment.

    Examples:
        copilot solution connection list
        copilot solution connection list --table
        copilot solution connection list --agent <agent-id>
    """
    try:
        client = get_client()
        connections = client.list_connection_references(bot_id=agent_id)

        if not connections:
            typer.echo("No connection references found.")
            return

        if table:
            formatted = [
                {
                    "name": c.get("connectionreferencedisplayname", ""),
                    "id": c.get("connectionreferenceid", ""),
                    "connector": c.get("connectorid", ""),
                    "status": c.get("statecode@OData.Community.Display.V1.FormattedValue", c.get("statecode", "")),
                }
                for c in connections
            ]
            print_table(
                formatted,
                columns=["name", "id", "connector", "status"],
                headers=["Name", "ID", "Connector", "Status"],
            )
        else:
            print_json(connections)

    except Exception as e:
        exit_code = handle_api_error(e)
        raise typer.Exit(exit_code)


# Register subgroups
app.add_typer(publisher_app, name="publisher")
app.add_typer(connection_app, name="connection")
