"""Main entry point for Copilot CLI."""
import typer
from typing import Optional

from .client import ClientError

# Create main Typer app
app = typer.Typer(
    name="copilot",
    help="CLI interface for Microsoft Copilot Studio agents via Dataverse API",
    add_completion=True,
)


# Import and register command modules
try:
    from .commands import agent
    app.add_typer(agent.app, name="agent", help="Manage Copilot Studio agents")
except ImportError:
    pass

try:
    from .commands import solution
    app.add_typer(solution.app, name="solution", help="Manage solutions and solution components")
except ImportError:
    pass

try:
    from .commands import flow
    app.add_typer(flow.app, name="flow", help="Manage Power Automate flows")
except ImportError:
    pass

try:
    from .commands import tool
    app.add_typer(tool.app, name="tool", help="Manage agent tools (prompts, REST APIs, MCP)")
except ImportError:
    pass

try:
    from .commands import connector
    app.add_typer(connector.app, name="connector", help="List Power Platform connectors")
except ImportError:
    pass

try:
    from .commands import environment
    app.add_typer(environment.app, name="environment", help="Manage Power Platform environments")
except ImportError:
    pass


@app.callback(invoke_without_command=True)
def callback(
    ctx: typer.Context,
    version: Optional[bool] = typer.Option(
        None,
        "--version",
        "-v",
        help="Show version and exit",
        is_eager=True,
    ),
):
    """
    Copilot CLI - Manage Microsoft Copilot Studio agents from the command line.

    Authentication is handled via Azure CLI ('az login') or environment variables:
    - DATAVERSE_URL (required) - Your Dataverse environment URL

    Optional (for service principal auth):
    - AZURE_TENANT_ID
    - AZURE_CLIENT_ID
    - AZURE_CLIENT_SECRET

    Examples:
        copilot agent list
        copilot agent list --table
        copilot agent get <agent_id>
    """
    if version:
        typer.echo("copilot-cli version 0.1.0")
        raise typer.Exit()

    # Show help if no command provided
    if ctx.invoked_subcommand is None:
        typer.echo(ctx.get_help())
        raise typer.Exit()


def main():
    """Main entry point for the CLI application."""
    try:
        app()
    except ClientError as e:
        typer.echo(f"Error: {e}", err=True)
        raise typer.Exit(2)
    except KeyboardInterrupt:
        typer.echo("\nAborted!", err=True)
        raise typer.Exit(130)
    except Exception as e:
        typer.echo(f"Unexpected error: {e}", err=True)
        raise typer.Exit(1)


if __name__ == "__main__":
    main()
