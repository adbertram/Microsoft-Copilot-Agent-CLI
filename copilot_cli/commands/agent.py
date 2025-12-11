"""Agent commands for Copilot CLI."""
import typer
import httpx
import time
import os
import base64
import json
import mimetypes
from pathlib import Path
from typing import Optional

from ..client import get_client
from ..output import (
    print_json,
    print_table,
    print_success,
    handle_api_error,
    format_bot_for_display,
    format_transcript_content,
    format_transcript_for_display,
)

app = typer.Typer(help="Manage Copilot Studio agents")


@app.command("list")
def list_agents(
    table: bool = typer.Option(
        False,
        "--table",
        "-t",
        help="Display output as a formatted table instead of JSON",
    ),
    all_fields: bool = typer.Option(
        False,
        "--all",
        "-a",
        help="Include all fields in the output (JSON mode only)",
    ),
):
    """
    List all Copilot Studio agents in the environment.

    Returns agents with their agent_id, name, schema name, and status.

    Examples:
        copilot agent list
        copilot agent list --table
        copilot agent list --all
    """
    try:
        client = get_client()

        if all_fields:
            bots = client.list_bots()
        else:
            bots = client.list_bots(
                select=["name", "botid", "schemaname", "statecode", "statuscode", "createdon", "modifiedon"]
            )

        if table:
            # Format for table display
            formatted = [format_bot_for_display(bot) for bot in bots]
            print_table(
                formatted,
                columns=["name", "botid", "statecode", "statuscode"],
                headers=["Name", "Agent ID", "State", "Status"],
            )
        else:
            if all_fields:
                print_json(bots)
            else:
                formatted = [format_bot_for_display(bot) for bot in bots]
                print_json(formatted)
    except Exception as e:
        exit_code = handle_api_error(e)
        raise typer.Exit(exit_code)


@app.command("get")
def get_agent(
    agent_id: str = typer.Argument(..., help="The agent's unique identifier (GUID)"),
    include_components: bool = typer.Option(
        False,
        "--components",
        "-c",
        help="Include agent components (topics, triggers, etc.)",
    ),
):
    """
    Get details for a specific Copilot Studio agent.

    Examples:
        copilot agent get fcef595a-30bb-f011-bbd3-000d3a8ba54e
        copilot agent get fcef595a-30bb-f011-bbd3-000d3a8ba54e --components
    """
    try:
        client = get_client()
        bot = client.get_bot(agent_id)

        if include_components:
            components = client.get_bot_components(agent_id)
            bot["components"] = components

        print_json(bot)
    except Exception as e:
        exit_code = handle_api_error(e)
        raise typer.Exit(exit_code)


@app.command("remove")
def remove_agent(
    agent_id: str = typer.Argument(..., help="The agent's unique identifier (GUID)"),
    force: bool = typer.Option(
        False,
        "--force",
        "-f",
        help="Skip confirmation prompt",
    ),
):
    """
    Remove (delete) a Copilot Studio agent.

    Examples:
        copilot agent remove fcef595a-30bb-f011-bbd3-000d3a8ba54e
        copilot agent remove fcef595a-30bb-f011-bbd3-000d3a8ba54e --force
    """
    try:
        client = get_client()

        # Get agent details first to show name in confirmation
        bot = client.get_bot(agent_id)
        agent_name = bot.get("name", agent_id)

        if not force:
            confirm = typer.confirm(f"Are you sure you want to delete agent '{agent_name}'?")
            if not confirm:
                typer.echo("Aborted.")
                raise typer.Exit(0)

        client.delete_bot(agent_id)
        print_success(f"Agent '{agent_name}' deleted successfully.")
    except Exception as e:
        exit_code = handle_api_error(e)
        raise typer.Exit(exit_code)


@app.command("publish")
def publish_agent(
    agent_id: str = typer.Argument(..., help="The agent's unique identifier (GUID)"),
):
    """
    Publish a Copilot Studio agent.

    Publishing makes the latest changes to your agent available to users.
    This includes changes to topics, knowledge sources, tools, and settings.

    Note: Publishing may take a few minutes to complete.

    Examples:
        copilot agent publish fcef595a-30bb-f011-bbd3-000d3a8ba54e
    """
    try:
        client = get_client()

        # Get agent details first to show name
        bot = client.get_bot(agent_id)
        agent_name = bot.get("name", agent_id)

        typer.echo(f"Publishing agent '{agent_name}'...")

        result = client.publish_bot(agent_id)

        if result.get("status") == "success":
            print_success(f"Agent '{agent_name}' published successfully!")
            if result.get("PublishedBotContentId"):
                typer.echo(f"Published Content ID: {result['PublishedBotContentId']}")
        else:
            typer.echo(f"Publish completed with status: {result}")
    except Exception as e:
        exit_code = handle_api_error(e)
        raise typer.Exit(exit_code)


@app.command("update")
def update_agent(
    agent_id: str = typer.Argument(..., help="The agent's unique identifier (GUID)"),
    name: Optional[str] = typer.Option(
        None,
        "--name",
        "-n",
        help="New display name for the agent",
    ),
    description: Optional[str] = typer.Option(
        None,
        "--description",
        "-d",
        help="New description for the agent",
    ),
    instructions: Optional[str] = typer.Option(
        None,
        "--instructions",
        "-i",
        help="New system instructions/prompt for the agent",
    ),
    instructions_file: Optional[str] = typer.Option(
        None,
        "--instructions-file",
        help="Path to file containing new system instructions",
    ),
    orchestration: Optional[bool] = typer.Option(
        None,
        "--orchestration/--no-orchestration",
        help="Enable/disable generative AI orchestration",
    ),
):
    """
    Update an existing Copilot Studio agent.

    Note: Model selection and web search must be configured via the Copilot Studio portal UI.

    Examples:
        copilot agent update <agent-id> --name "New Name"
        copilot agent update <agent-id> --description "New description"
        copilot agent update <agent-id> --instructions "New system prompt"
        copilot agent update <agent-id> --instructions-file ./prompt.txt
        copilot agent update <agent-id> --no-orchestration
    """
    try:
        # Handle instructions from file if provided
        agent_instructions = instructions
        if instructions_file:
            try:
                with open(instructions_file, "r") as f:
                    agent_instructions = f.read()
            except FileNotFoundError:
                typer.echo(f"Error: Instructions file not found: {instructions_file}", err=True)
                raise typer.Exit(1)
            except IOError as e:
                typer.echo(f"Error reading instructions file: {e}", err=True)
                raise typer.Exit(1)

        client = get_client()

        # Get current agent name for success message
        current_bot = client.get_bot(agent_id)
        agent_name = name if name else current_bot.get("name", agent_id)

        client.update_bot(
            bot_id=agent_id,
            name=name,
            instructions=agent_instructions,
            description=description,
            orchestration=orchestration,
        )

        print_success(f"Agent '{agent_name}' updated successfully.")
    except Exception as e:
        exit_code = handle_api_error(e)
        raise typer.Exit(exit_code)


@app.command("create")
def create_agent(
    name: str = typer.Option(
        ...,
        "--name",
        "-n",
        help="Display name for the agent",
    ),
    description: Optional[str] = typer.Option(
        None,
        "--description",
        "-d",
        help="Description for the agent",
    ),
    schema_name: Optional[str] = typer.Option(
        None,
        "--schema-name",
        "-s",
        help="Internal schema name (auto-generated from name if not provided)",
    ),
    instructions: Optional[str] = typer.Option(
        None,
        "--instructions",
        "-i",
        help="System instructions/prompt for the agent",
    ),
    instructions_file: Optional[str] = typer.Option(
        None,
        "--instructions-file",
        help="Path to file containing system instructions",
    ),
    orchestration: bool = typer.Option(
        True,
        "--orchestration/--no-orchestration",
        help="Enable/disable generative AI orchestration (default: enabled)",
    ),
):
    """
    Create a new Copilot Studio agent.

    Note: Model selection and web search must be configured via the Copilot Studio portal UI.

    Examples:
        copilot agent create --name "My Agent"
        copilot agent create --name "My Agent" --description "A helpful assistant"
        copilot agent create --name "My Agent" --instructions "You are a helpful assistant"
        copilot agent create --name "My Agent" --instructions-file ./prompt.txt
        copilot agent create --name "My Agent" --no-orchestration
    """
    try:
        # Handle instructions from file if provided
        agent_instructions = instructions
        if instructions_file:
            try:
                with open(instructions_file, "r") as f:
                    agent_instructions = f.read()
            except FileNotFoundError:
                typer.echo(f"Error: Instructions file not found: {instructions_file}", err=True)
                raise typer.Exit(1)
            except IOError as e:
                typer.echo(f"Error reading instructions file: {e}", err=True)
                raise typer.Exit(1)

        client = get_client()
        client.create_bot(
            name=name,
            schema_name=schema_name,
            instructions=agent_instructions,
            description=description,
            orchestration=orchestration,
        )

        print_success(f"Agent '{name}' created successfully.")
    except Exception as e:
        exit_code = handle_api_error(e)
        raise typer.Exit(exit_code)


# =============================================================================
# Direct Line API Constants
# =============================================================================

DIRECTLINE_URL = "https://directline.botframework.com/v3/directline"


@app.command("prompt")
def prompt_agent(
    agent_id: str = typer.Argument(
        ...,
        help="The agent's unique identifier (GUID)",
    ),
    message: str = typer.Option(
        ...,
        "--message",
        "-m",
        help="The message/prompt to send to the agent",
    ),
    secret: Optional[str] = typer.Option(
        None,
        "--secret",
        "-s",
        help="Direct Line secret (or set DIRECTLINE_SECRET env var)",
    ),
    entra_id: bool = typer.Option(
        False,
        "--entra-id",
        help="Use Entra ID (Azure AD) authentication instead of Direct Line secret",
    ),
    client_id: Optional[str] = typer.Option(
        None,
        "--client-id",
        help="Entra ID application (client) ID (or set ENTRA_CLIENT_ID env var)",
    ),
    tenant_id: Optional[str] = typer.Option(
        None,
        "--tenant-id",
        help="Entra ID tenant ID (or set ENTRA_TENANT_ID env var)",
    ),
    scope: Optional[str] = typer.Option(
        None,
        "--scope",
        help="OAuth scope (default: https://api.powerplatform.com/.default)",
    ),
    token_endpoint: Optional[str] = typer.Option(
        None,
        "--token-endpoint",
        help="Agent token endpoint URL (from Copilot Studio > Channels > Mobile app)",
    ),
    max_polls: int = typer.Option(
        30,
        "--max-polls",
        help="Maximum number of polling attempts for response",
    ),
    poll_interval: int = typer.Option(
        3,
        "--poll-interval",
        help="Seconds between polling attempts",
    ),
    timeout: int = typer.Option(
        120,
        "--timeout",
        help="Total timeout in seconds for the request",
    ),
    verbose: bool = typer.Option(
        False,
        "--verbose",
        "-v",
        help="Show detailed progress and response information",
    ),
    json_output: bool = typer.Option(
        False,
        "--json",
        "-j",
        help="Output response as JSON",
    ),
    file: Optional[str] = typer.Option(
        None,
        "--file",
        "-f",
        help="Path to a file to attach (Word, PDF, text, markdown, etc.)",
    ),
):
    """
    Send a prompt to a Copilot Studio agent and get the response.

    Uses the Direct Line API to communicate with the agent. Supports two authentication modes:

    1. Direct Line Secret (default): Requires a Direct Line secret from Copilot Studio.
    2. Entra ID: Uses device code flow to authenticate with Microsoft Entra ID.

    ENTRA ID SETUP (Azure App Registration):

    1. Create an app registration in Azure Portal
    2. Go to Authentication > Enable "Allow public client flows"
    3. Go to API permissions > Add permission > APIs my organization uses
    4. Search for "Power Platform API" (ID: 8578e004-a5c6-46e7-913e-12f58912df43)
    5. Add delegated permission: CopilotStudio.Copilots.Invoke
    6. Grant admin consent for the permission
    7. Note the Application (client) ID and Directory (tenant) ID

    ENTRA ID SETUP (Copilot Studio):

    1. Go to Settings > Security > Authentication
    2. Select "Authenticate manually"
    3. Set Service provider to "Microsoft Entra ID V2"
    4. Enter your Client ID and Tenant ID
    5. Save and Publish the agent

    GET TOKEN ENDPOINT:

    The token endpoint URL format is:
    https://{ENV}.environment.api.powerplatform.com/powervirtualagents/botsbyschema/{BOT_SCHEMA_NAME}/directline/token?api-version=2022-03-01-preview

    Where:
    - {ENV}: Your environment ID (found in Copilot Studio > Channels > Mobile app)
    - {AGENT_SCHEMA_NAME}: Your agent's schema name (e.g., cr83c_myAgent)

    Examples:
        # Using Direct Line secret
        copilot agent prompt <agent-id> --message "Hello" --secret "your-secret"

        # Using Entra ID authentication
        copilot agent prompt <agent-id> -m "Hello" --entra-id \\
            --client-id <app-client-id> --tenant-id <tenant-id> \\
            --token-endpoint "https://{ENV}.environment.api.powerplatform.com/powervirtualagents/botsbyschema/{AGENT}/directline/token?api-version=2022-03-01-preview"

        # With file attachment
        copilot agent prompt <agent-id> -m "Review this" --file ./draft.docx --secret "xxx"

    Environment Variables:
        DIRECTLINE_SECRET - Direct Line secret (alternative to --secret)
        ENTRA_CLIENT_ID - Entra ID client ID (alternative to --client-id)
        ENTRA_TENANT_ID - Entra ID tenant ID (alternative to --tenant-id)
        ENTRA_SCOPE - OAuth scope (default: https://api.powerplatform.com/.default)
        AGENT_TOKEN_ENDPOINT - Agent token endpoint (alternative to --token-endpoint)
    """
    try:
        # Determine authentication method
        directline_token = None
        user_id = f"copilot-cli-{int(time.time())}"

        if entra_id:
            # Entra ID authentication flow
            entra_client_id = client_id or os.environ.get("ENTRA_CLIENT_ID")
            entra_tenant_id = tenant_id or os.environ.get("ENTRA_TENANT_ID")
            # Default to Power Platform API scope with CopilotStudio.Copilots.Invoke permission
            entra_scope = scope or os.environ.get("ENTRA_SCOPE") or "https://api.powerplatform.com/.default"
            agent_token_endpoint = token_endpoint or os.environ.get("AGENT_TOKEN_ENDPOINT") or os.environ.get("BOT_TOKEN_ENDPOINT")

            if not entra_client_id:
                typer.echo("Error: --client-id or ENTRA_CLIENT_ID env var required for Entra ID auth", err=True)
                raise typer.Exit(1)
            if not entra_tenant_id:
                typer.echo("Error: --tenant-id or ENTRA_TENANT_ID env var required for Entra ID auth", err=True)
                raise typer.Exit(1)
            if not agent_token_endpoint:
                typer.echo("Error: --token-endpoint or AGENT_TOKEN_ENDPOINT env var required for Entra ID auth", err=True)
                typer.echo("Get endpoint from: Copilot Studio > Channels > Mobile app > Token Endpoint", err=True)
                raise typer.Exit(1)

            if verbose:
                typer.echo("Using Entra ID authentication...")
                typer.echo(f"  Client ID: {entra_client_id[:8]}...")
                typer.echo(f"  Tenant ID: {entra_tenant_id[:8]}...")
                typer.echo(f"  Scope: {entra_scope}")

            # Step 1: Acquire access token using MSAL device code flow
            try:
                import msal
            except ImportError:
                typer.echo("Error: msal package required for Entra ID auth. Install with: pip install msal", err=True)
                raise typer.Exit(1)

            # Set up persistent token cache in CLI project directory
            cache_file = Path(__file__).parent.parent.parent / ".token-cache.json"
            cache = msal.SerializableTokenCache()

            if cache_file.exists():
                try:
                    cache.deserialize(cache_file.read_text())
                    if verbose:
                        typer.echo(f"Loaded token cache from {cache_file}")
                except Exception:
                    pass  # Ignore cache load errors

            authority = f"https://login.microsoftonline.com/{entra_tenant_id}"
            app = msal.PublicClientApplication(
                client_id=entra_client_id,
                authority=authority,
                token_cache=cache,
            )

            # Check cache for existing tokens
            accounts = app.get_accounts()
            access_token = None

            if accounts:
                if verbose:
                    typer.echo("Found cached account, attempting silent token acquisition...")
                result = app.acquire_token_silent(scopes=[entra_scope], account=accounts[0])
                if result and "access_token" in result:
                    access_token = result["access_token"]
                    if verbose:
                        typer.echo("Token acquired from cache.")

            if not access_token:
                # Initiate device code flow
                if verbose:
                    typer.echo("Initiating device code flow...")

                flow = app.initiate_device_flow(scopes=[entra_scope])
                if "user_code" not in flow:
                    typer.echo(f"Error: Failed to initiate device flow: {flow.get('error_description', 'Unknown error')}", err=True)
                    raise typer.Exit(1)

                # Display device code message to user
                typer.echo("")
                typer.echo(flow["message"])
                typer.echo("")

                # Wait for user to complete authentication
                result = app.acquire_token_by_device_flow(flow)

                if "error" in result:
                    typer.echo(f"Error: Authentication failed: {result.get('error_description', result.get('error'))}", err=True)
                    raise typer.Exit(1)

                access_token = result["access_token"]
                if verbose:
                    typer.echo("Authentication successful!")

            # Save token cache if it changed
            if cache.has_state_changed:
                try:
                    cache_file.write_text(cache.serialize())
                    if verbose:
                        typer.echo(f"Saved token cache to {cache_file}")
                except Exception as e:
                    if verbose:
                        typer.echo(f"Warning: Could not save token cache: {e}", err=True)

            # Step 2: Exchange Entra ID token for Direct Line token
            # The token endpoint returns a Direct Line token when called with Bearer auth
            if verbose:
                typer.echo("Exchanging Entra ID token for Direct Line token...")

            with httpx.Client(timeout=30.0) as token_client:
                token_response = token_client.get(
                    agent_token_endpoint,
                    headers={"Authorization": f"Bearer {access_token}"},
                )

                if verbose:
                    typer.echo(f"Token endpoint response: HTTP {token_response.status_code}")

                if token_response.status_code != 200:
                    typer.echo(f"Error: Failed to get Direct Line token (HTTP {token_response.status_code})", err=True)
                    if verbose:
                        typer.echo(f"Response: {token_response.text}", err=True)
                    raise typer.Exit(1)

                token_data = token_response.json()
                directline_token = token_data.get("token")

                if not directline_token:
                    typer.echo("Error: No token in response", err=True)
                    if verbose:
                        typer.echo(f"Response: {token_data}", err=True)
                    raise typer.Exit(1)

                if verbose:
                    typer.echo("Direct Line token obtained successfully!")

        else:
            # Direct Line secret authentication (original flow)
            directline_secret = secret or os.environ.get("DIRECTLINE_SECRET")
            if not directline_secret:
                typer.echo(
                    "Error: Direct Line secret required. Provide via --secret or DIRECTLINE_SECRET env var.",
                    err=True,
                )
                typer.echo(
                    "Get secret from: Copilot Studio > Settings > Channels > Direct Line",
                    err=True,
                )
                typer.echo(
                    "Or use --entra-id for Entra ID authentication.",
                    err=True,
                )
                raise typer.Exit(1)
            directline_token = directline_secret

        # Handle file attachment (upload via Direct Line upload endpoint)
        file_to_upload = None
        if file:
            file_path = Path(file)
            if not file_path.exists():
                typer.echo(f"Error: File not found: {file}", err=True)
                raise typer.Exit(1)

            file_name = file_path.name
            ext = file_path.suffix.lower()

            # Map file extensions to MIME types
            mime_types = {
                ".txt": "text/plain",
                ".md": "text/markdown",
                ".json": "application/json",
                ".xml": "application/xml",
                ".html": "text/html",
                ".csv": "text/csv",
                ".yaml": "application/x-yaml",
                ".yml": "application/x-yaml",
                ".pdf": "application/pdf",
                ".docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                ".doc": "application/msword",
                ".png": "image/png",
                ".jpg": "image/jpeg",
                ".jpeg": "image/jpeg",
                ".gif": "image/gif",
            }

            content_type = mime_types.get(ext)
            if not content_type:
                typer.echo(f"Error: Unsupported file type: {ext}", err=True)
                typer.echo(f"Supported types: {', '.join(mime_types.keys())}", err=True)
                raise typer.Exit(1)

            # Read file content
            try:
                with open(file_path, "rb") as f:
                    file_content = f.read()
                file_to_upload = {
                    "name": file_name,
                    "content": file_content,
                    "content_type": content_type,
                }
                if verbose:
                    typer.echo(f"Prepared file for upload: {file_name} ({len(file_content)} bytes, {content_type})")
            except IOError as e:
                typer.echo(f"Error reading file: {e}", err=True)
                raise typer.Exit(1)

        # Start conversation via Direct Line API
        if verbose:
            typer.echo(f"Starting conversation with agent {agent_id}...")

        with httpx.Client(timeout=30.0) as client:
            conv_response = client.post(
                f"{DIRECTLINE_URL}/conversations",
                headers={
                    "Authorization": f"Bearer {directline_token}",
                    "Content-Type": "application/json",
                },
            )

            if conv_response.status_code == 403:
                typer.echo("Error: Authentication failed (HTTP 403)", err=True)
                if entra_id:
                    typer.echo("Check that the Entra ID token exchange was successful", err=True)
                else:
                    typer.echo("Check that the Direct Line secret is valid and not expired", err=True)
                raise typer.Exit(1)

            if conv_response.status_code != 201:
                typer.echo(f"Error: Failed to start conversation (HTTP {conv_response.status_code})", err=True)
                if verbose:
                    typer.echo(f"Response: {conv_response.text}", err=True)
                raise typer.Exit(1)

            conv_data = conv_response.json()
            conv_id = conv_data.get("conversationId")

            if not conv_id:
                typer.echo("Error: No conversation ID in response", err=True)
                raise typer.Exit(1)

            if verbose:
                typer.echo(f"Conversation started: {conv_id}")

            # Step 4: Send message (with file upload if applicable)
            if verbose:
                typer.echo(f"Sending message: \"{message}\"")

            if file_to_upload:
                # Use Direct Line upload endpoint for file attachments
                # This uses multipart/form-data with the activity and file
                import json as json_module

                activity_json = json_module.dumps({
                    "type": "message",
                    "from": {"id": user_id, "name": "Copilot CLI"},
                    "text": message,
                })

                # Build multipart form data
                files = {
                    "activity": (None, activity_json, "application/vnd.microsoft.activity"),
                    "file": (file_to_upload["name"], file_to_upload["content"], file_to_upload["content_type"]),
                }

                if verbose:
                    typer.echo(f"Uploading file via Direct Line: {file_to_upload['name']}")

                send_response = client.post(
                    f"{DIRECTLINE_URL}/conversations/{conv_id}/upload?userId={user_id}",
                    headers={
                        "Authorization": f"Bearer {directline_token}",
                    },
                    files=files,
                )
            else:
                # Standard message without file
                send_payload = {
                    "type": "message",
                    "from": {"id": user_id, "name": "Copilot CLI"},
                    "text": message,
                }

                send_response = client.post(
                    f"{DIRECTLINE_URL}/conversations/{conv_id}/activities",
                    headers={
                        "Authorization": f"Bearer {directline_token}",
                        "Content-Type": "application/json",
                    },
                    json=send_payload,
                )

            if send_response.status_code not in (200, 201, 204):
                typer.echo(f"Error: Failed to send message (HTTP {send_response.status_code})", err=True)
                if verbose:
                    typer.echo(f"Response: {send_response.text}", err=True)
                raise typer.Exit(1)

            activity_id = send_response.json().get("id") if send_response.text else None
            if verbose:
                typer.echo(f"Message sent (Activity ID: {activity_id})")

            # Step 5: Poll for response
            if verbose:
                typer.echo(f"Polling for response (max {max_polls} attempts, {poll_interval}s interval)...")

            bot_response = None
            bot_from = None
            watermark = None
            poll_count = 0
            start_time = time.time()

            while bot_response is None and poll_count < max_polls:
                # Check timeout
                if time.time() - start_time > timeout:
                    typer.echo(f"Error: Timeout after {timeout} seconds", err=True)
                    raise typer.Exit(1)

                poll_count += 1
                time.sleep(poll_interval)

                # Build URL with watermark
                activities_url = f"{DIRECTLINE_URL}/conversations/{conv_id}/activities"
                if watermark:
                    activities_url = f"{activities_url}?watermark={watermark}"

                activities_response = client.get(
                    activities_url,
                    headers={"Authorization": f"Bearer {directline_token}"},
                )

                if activities_response.status_code != 200:
                    if verbose:
                        typer.echo(f"Warning: Poll failed (HTTP {activities_response.status_code})", err=True)
                    continue

                activities_data = activities_response.json()
                watermark = activities_data.get("watermark")

                # Find bot messages (exclude our user messages)
                activities = activities_data.get("activities", [])
                bot_messages = [
                    a for a in activities
                    if a.get("type") == "message" and a.get("from", {}).get("id") != user_id
                ]

                if bot_messages:
                    # Get the last bot message
                    last_message = bot_messages[-1]
                    bot_response = last_message.get("text", "")
                    bot_from = last_message.get("from", {}).get("name") or last_message.get("from", {}).get("id")

                if verbose and not bot_response:
                    typer.echo(f"  Polling... attempt {poll_count}/{max_polls}", nl=False)
                    typer.echo("\r", nl=False)

            if verbose:
                typer.echo("")  # Clear the polling line

            if not bot_response:
                typer.echo(f"Error: No response received after {poll_count} polling attempts", err=True)
                typer.echo("Possible causes:", err=True)
                typer.echo("  - Agent is not published", err=True)
                typer.echo("  - Agent is experiencing errors (check Copilot Studio)", err=True)
                typer.echo("  - Direct Line channel is not enabled", err=True)
                raise typer.Exit(1)

            # Check for error responses
            is_error = any(phrase in bot_response for phrase in [
                "something unexpected happened",
                "Error code:",
                "InvalidContent",
                "We're looking into it",
            ])

            # Output the response
            if json_output:
                result = {
                    "success": not is_error,
                    "response": bot_response,
                    "conversationId": conv_id,
                    "pollCount": poll_count,
                    "respondent": bot_from,
                }
                if is_error:
                    result["error"] = True
                print_json(result)
            else:
                if verbose:
                    typer.echo(f"Response from {bot_from} (after {poll_count} poll(s)):")
                    typer.echo("")

                typer.echo(bot_response)

                if is_error:
                    typer.echo("")
                    typer.echo("Warning: Agent returned an error response", err=True)
                    raise typer.Exit(1)

    except httpx.TimeoutException:
        typer.echo("Error: Request timed out", err=True)
        raise typer.Exit(1)
    except httpx.RequestError as e:
        typer.echo(f"Error: Request failed: {e}", err=True)
        raise typer.Exit(1)
    except Exception as e:
        if isinstance(e, typer.Exit):
            raise
        exit_code = handle_api_error(e)
        raise typer.Exit(exit_code)


# Knowledge source commands as a subgroup
# Usage: copilot agent knowledge list --agent <agent-id>
#        copilot agent knowledge file add --agent <agent-id> ...
#        copilot agent knowledge azure-ai-search add --agent <agent-id> ...

knowledge_app = typer.Typer(help="Manage knowledge sources for an agent")

# Component type mapping
COMPONENT_TYPE_NAMES = {
    14: "file",
    16: "azure-ai-search",
}


def format_knowledge_source(source: dict) -> dict:
    """Format a knowledge source for display."""
    component_type = source.get("componenttype", 14)
    type_name = COMPONENT_TYPE_NAMES.get(component_type, f"unknown({component_type})")
    return {
        "name": source.get("name"),
        "type": type_name,
        "component_id": source.get("botcomponentid"),
        "description": source.get("description"),
    }


@knowledge_app.command("list")
def knowledge_list(
    agent_id: str = typer.Option(
        ...,
        "--agent",
        "-a",
        help="The agent's unique identifier (GUID)",
    ),
    source_type: Optional[str] = typer.Option(
        None,
        "--type",
        help="Filter by type: 'file' or 'connector'",
    ),
    table: bool = typer.Option(
        False,
        "--table",
        "-t",
        help="Display output as a formatted table instead of JSON",
    ),
):
    """
    List knowledge sources for an agent.

    Examples:
        copilot agent knowledge list --agent <agent-id>
        copilot agent knowledge list --agent <agent-id> --table
        copilot agent knowledge list --agent <agent-id> --type file
    """
    try:
        client = get_client()
        sources = client.list_knowledge_sources(agent_id, source_type=source_type)

        if not sources:
            typer.echo("No knowledge sources found for this agent.")
            return

        formatted = [format_knowledge_source(s) for s in sources]

        if table:
            print_table(
                formatted,
                columns=["name", "type", "component_id", "description"],
                headers=["Name", "Type", "Component ID", "Description"],
            )
        else:
            print_json(formatted)
    except Exception as e:
        exit_code = handle_api_error(e)
        raise typer.Exit(exit_code)


@knowledge_app.command("remove")
def knowledge_remove(
    agent_id: str = typer.Option(
        ...,
        "--agent",
        "-a",
        help="The agent's unique identifier (GUID)",
    ),
    component_id: str = typer.Argument(..., help="The knowledge source component's unique identifier (GUID)"),
    force: bool = typer.Option(
        False,
        "--force",
        "-f",
        help="Skip confirmation prompt",
    ),
):
    """
    Remove a knowledge source from an agent.

    Examples:
        copilot agent knowledge remove --agent <agent-id> <component-id>
        copilot agent knowledge remove --agent <agent-id> <component-id> --force
    """
    try:
        if not force:
            confirm = typer.confirm("Are you sure you want to delete this knowledge source?")
            if not confirm:
                typer.echo("Aborted.")
                raise typer.Exit(0)

        client = get_client()
        client.remove_knowledge_source(component_id)
        print_success("Knowledge source deleted successfully.")
    except Exception as e:
        exit_code = handle_api_error(e)
        raise typer.Exit(exit_code)


# File knowledge source subgroup
file_app = typer.Typer(help="Manage file-based knowledge sources")


@file_app.command("add")
def file_add(
    agent_id: str = typer.Option(
        ...,
        "--agent",
        "-a",
        help="The agent's unique identifier (GUID)",
    ),
    name: str = typer.Option(
        ...,
        "--name",
        "-n",
        help="Display name for the knowledge source",
    ),
    content: Optional[str] = typer.Option(
        None,
        "--content",
        "-c",
        help="Text content for the knowledge source",
    ),
    file: Optional[str] = typer.Option(
        None,
        "--file",
        "-f",
        help="Path to file containing the knowledge content",
    ),
    description: Optional[str] = typer.Option(
        None,
        "--description",
        "-d",
        help="Description for the knowledge source (auto-generated if not provided)",
    ),
):
    """
    Add a file-based knowledge source to an agent.

    Provide content either via --content or --file.

    Examples:
        copilot agent knowledge file add --agent <agent-id> --name "FAQ" --content "Q: What? A: Test."
        copilot agent knowledge file add --agent <agent-id> --name "Guide" --file ./guide.md
    """
    try:
        # Validate input
        if not content and not file:
            typer.echo("Error: Provide content via --content or --file", err=True)
            raise typer.Exit(1)

        if content and file:
            typer.echo("Error: Provide either --content or --file, not both", err=True)
            raise typer.Exit(1)

        # Read content from file if provided
        knowledge_content = content
        if file:
            try:
                with open(file, "r") as f:
                    knowledge_content = f.read()
            except FileNotFoundError:
                typer.echo(f"Error: File not found: {file}", err=True)
                raise typer.Exit(1)
            except IOError as e:
                typer.echo(f"Error reading file: {e}", err=True)
                raise typer.Exit(1)

        client = get_client()
        component_id = client.add_file_knowledge_source(
            bot_id=agent_id,
            name=name,
            content=knowledge_content,
            description=description,
        )

        print_success(f"File knowledge source '{name}' added successfully.")
        if component_id:
            typer.echo(f"Component ID: {component_id}")
    except Exception as e:
        exit_code = handle_api_error(e)
        raise typer.Exit(exit_code)


knowledge_app.add_typer(file_app, name="file")


# Azure AI Search knowledge source subgroup
azure_search_app = typer.Typer(help="Manage Azure AI Search knowledge sources")


@azure_search_app.command("add")
def azure_search_add(
    agent_id: str = typer.Option(
        ...,
        "--agent",
        "-a",
        help="The agent's unique identifier (GUID)",
    ),
    name: str = typer.Option(
        ...,
        "--name",
        "-n",
        help="Display name for the knowledge source",
    ),
    endpoint: str = typer.Option(
        ...,
        "--endpoint",
        "-e",
        help="Azure AI Search endpoint URL (e.g., https://mysearch.search.windows.net)",
    ),
    index: str = typer.Option(
        ...,
        "--index",
        "-i",
        help="Name of the Azure AI Search index",
    ),
    api_key: str = typer.Option(
        ...,
        "--api-key",
        "-k",
        help="Azure AI Search API key (admin or query key)",
    ),
    description: Optional[str] = typer.Option(
        None,
        "--description",
        "-d",
        help="Description for the knowledge source",
    ),
):
    """
    Add an Azure AI Search knowledge source to an agent (EXPERIMENTAL).

    WARNING: This command creates an agent component record but the knowledge source
    may not appear in Copilot Studio UI. Copilot Studio requires a Power Platform
    connection to be properly linked, which involves internal configuration not
    exposed via public APIs.

    RECOMMENDED APPROACH:
    1. Use 'copilot connection create' to create the Power Platform connection
    2. Link the connection to your agent via the Copilot Studio UI

    Examples:
        copilot agent knowledge azure-ai-search add --agent <agent-id> \\
            --name "Product Docs" \\
            --endpoint https://mysearch.search.windows.net \\
            --index products-index \\
            --api-key <your-api-key>
    """
    try:
        client = get_client()
        component_id = client.add_azure_ai_search_knowledge_source(
            bot_id=agent_id,
            name=name,
            search_endpoint=endpoint,
            search_index=index,
            api_key=api_key,
            description=description,
        )

        print_success(f"Azure AI Search knowledge source '{name}' added successfully.")
        if component_id:
            typer.echo(f"Component ID: {component_id}")
    except Exception as e:
        exit_code = handle_api_error(e)
        raise typer.Exit(exit_code)


knowledge_app.add_typer(azure_search_app, name="azure-ai-search")


# Register knowledge subgroup
app.add_typer(knowledge_app, name="knowledge")




# =============================================================================
# Transcript Commands
# =============================================================================

transcript_app = typer.Typer(help="View conversation transcripts for troubleshooting")


def _is_guid(value: str) -> bool:
    """Check if a string looks like a GUID."""
    import re
    guid_pattern = r'^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$'
    return bool(re.match(guid_pattern, value))


@transcript_app.command("list")
def transcript_list(
    agent: Optional[str] = typer.Option(
        None,
        "--agent",
        "-a",
        help="Filter by agent name or ID",
    ),
    limit: int = typer.Option(
        20,
        "--limit",
        "-l",
        help="Maximum number of transcripts to return",
    ),
    table: bool = typer.Option(
        False,
        "--table",
        "-t",
        help="Display output as a formatted table instead of JSON",
    ),
):
    """
    List conversation transcripts.

    Shows recent conversation transcripts, optionally filtered by agent name or ID.

    Examples:
        copilot agent transcript list
        copilot agent transcript list --table
        copilot agent transcript list --agent "Writer Draft Reviewer" --limit 10
        copilot agent transcript list --agent d2735b5c-aecb-f011-bbd3-000d3a8ba54e
    """
    try:
        client = get_client()

        # Determine if agent is an ID or name
        agent_id = None
        agent_name = None
        if agent:
            if _is_guid(agent):
                agent_id = agent
            else:
                agent_name = agent

        transcripts = client.list_transcripts(bot_id=agent_id, bot_name=agent_name, limit=limit)

        if not transcripts:
            typer.echo("No transcripts found.")
            return

        if table:
            formatted = [format_transcript_for_display(t) for t in transcripts]
            print_table(
                formatted,
                columns=["id", "agent_name", "start_time"],
                headers=["ID", "Agent", "Start Time"],
            )
        else:
            formatted = [format_transcript_for_display(t) for t in transcripts]
            print_json(formatted)
    except Exception as e:
        exit_code = handle_api_error(e)
        raise typer.Exit(exit_code)


@transcript_app.command("get")
def transcript_get(
    transcript_id: str = typer.Argument(
        ...,
        help="The transcript's unique identifier (GUID)",
    ),
    pretty: bool = typer.Option(
        False,
        "--pretty",
        "-p",
        help="Output as formatted conversation instead of JSON",
    ),
):
    """
    Get full transcript content.

    By default, outputs JSON. Use --pretty for a formatted conversation view.

    Examples:
        copilot agent transcript get <transcript-id>
        copilot agent transcript get <transcript-id> --pretty
    """
    try:
        client = get_client()
        transcript = client.get_transcript(transcript_id)

        if not pretty:
            print_json(transcript)
            return

        # Pretty format the transcript
        name = transcript.get("name", "Unknown")
        # Get bot name from OData annotation, fall back to ID
        bot_name = transcript.get(
            "_bot_conversationtranscriptid_value@OData.Community.Display.V1.FormattedValue",
            transcript.get("_bot_conversationtranscriptid_value", "Unknown"),
        )
        start_time = transcript.get("conversationstarttime", "Unknown")
        if start_time:
            start_time = start_time.replace("T", " ").replace("Z", "")
        content = transcript.get("content", "")

        typer.echo(f"Transcript: {name}")
        typer.echo(f"Agent: {bot_name}")
        typer.echo(f"Started: {start_time}")
        typer.echo("")
        typer.echo("--- Conversation ---")
        typer.echo(format_transcript_content(content))

    except Exception as e:
        exit_code = handle_api_error(e)
        raise typer.Exit(exit_code)


# Register transcript subgroup
app.add_typer(transcript_app, name="transcript")


# =============================================================================
# Topic Commands
# =============================================================================

topic_app = typer.Typer(help="Manage agent topics")

# Topic component type mapping
TOPIC_COMPONENT_TYPE_NAMES = {
    0: "Topic",
    9: "Topic (V2)",
}


def format_topic_for_display(topic: dict) -> dict:
    """Format a topic for display."""
    component_type = topic.get("componenttype", 0)
    type_name = TOPIC_COMPONENT_TYPE_NAMES.get(component_type, f"unknown({component_type})")

    return {
        "name": topic.get("name"),
        "component_type": type_name,
        "component_id": topic.get("botcomponentid"),
        "schema_name": topic.get("schemaname"),
        "status": topic.get("statecode@OData.Community.Display.V1.FormattedValue", "Active"),
    }


@topic_app.command("list")
def topic_list(
    agent_id: str = typer.Option(
        ...,
        "--agentId",
        "-a",
        help="The agent's unique identifier (GUID)",
    ),
    table: bool = typer.Option(
        False,
        "--table",
        "-t",
        help="Display output as a formatted table instead of JSON",
    ),
    system: bool = typer.Option(
        False,
        "--system",
        "-s",
        help="List only system topics (built-in, managed)",
    ),
    custom: bool = typer.Option(
        False,
        "--custom",
        "-c",
        help="List only custom topics (user-created)",
    ),
):
    """
    List topics for an agent.

    Examples:
        copilot agent topic list --agentId <agent-id>
        copilot agent topic list --agentId <agent-id> --table
        copilot agent topic list --agentId <agent-id> --system --table
        copilot agent topic list --agentId <agent-id> --custom --table
    """
    try:
        if system and custom:
            print_error("Cannot specify both --system and --custom")
            raise typer.Exit(1)

        client = get_client()
        topics = client.list_topics(agent_id, system_only=system, custom_only=custom)

        if not topics:
            filter_type = "system " if system else "custom " if custom else ""
            typer.echo(f"No {filter_type}topics found for this agent.")
            return

        formatted = [format_topic_for_display(t) for t in topics]

        if table:
            print_table(
                formatted,
                columns=["name", "component_type", "status", "component_id"],
                headers=["Name", "Component Type", "Status", "Component ID"],
            )
        else:
            print_json(formatted)
    except Exception as e:
        exit_code = handle_api_error(e)
        raise typer.Exit(exit_code)


@topic_app.command("enable")
def topic_enable(
    topic_id: str = typer.Argument(
        ...,
        help="The topic's component ID (GUID)",
    ),
):
    """
    Enable a topic.

    Sets the topic state to Active so it will be triggered during conversations.

    Examples:
        copilot agent topic enable <topic-id>
    """
    try:
        client = get_client()

        # Get topic name for confirmation message
        topic = client.get_topic(topic_id)
        topic_name = topic.get("name", topic_id)

        client.set_topic_state(topic_id, enabled=True)
        print_success(f"Topic '{topic_name}' enabled successfully.")
    except Exception as e:
        exit_code = handle_api_error(e)
        raise typer.Exit(exit_code)


@topic_app.command("delete")
@topic_app.command("remove")
def topic_delete(
    topic_id: str = typer.Argument(
        ...,
        help="The topic's component ID (GUID)",
    ),
    force: bool = typer.Option(
        False,
        "--force",
        "-f",
        help="Skip confirmation prompt",
    ),
):
    """
    Delete a topic.

    Permanently removes the topic from the agent. This action cannot be undone.

    Examples:
        copilot agent topic delete <topic-id>
        copilot agent topic delete <topic-id> --force
    """
    try:
        client = get_client()

        # Get topic name for confirmation message
        topic = client.get_topic(topic_id)
        topic_name = topic.get("name", topic_id)

        if not force:
            confirm = typer.confirm(f"Are you sure you want to delete topic '{topic_name}'? This cannot be undone.")
            if not confirm:
                typer.echo("Aborted.")
                raise typer.Exit(0)

        client.delete(f"botcomponents({topic_id})")
        print_success(f"Topic '{topic_name}' deleted successfully.")
    except Exception as e:
        exit_code = handle_api_error(e)
        raise typer.Exit(exit_code)


@topic_app.command("disable")
def topic_disable(
    topic_id: str = typer.Argument(
        ...,
        help="The topic's component ID (GUID)",
    ),
    force: bool = typer.Option(
        False,
        "--force",
        "-f",
        help="Skip confirmation prompt",
    ),
):
    """
    Disable a topic.

    Sets the topic state to Inactive so it will not be triggered during conversations.

    Examples:
        copilot agent topic disable <topic-id>
        copilot agent topic disable <topic-id> --force
    """
    try:
        client = get_client()

        # Get topic name for confirmation message
        topic = client.get_topic(topic_id)
        topic_name = topic.get("name", topic_id)

        if not force:
            confirm = typer.confirm(f"Are you sure you want to disable topic '{topic_name}'?")
            if not confirm:
                typer.echo("Aborted.")
                raise typer.Exit(0)

        client.set_topic_state(topic_id, enabled=False)
        print_success(f"Topic '{topic_name}' disabled successfully.")
    except Exception as e:
        exit_code = handle_api_error(e)
        raise typer.Exit(exit_code)


@topic_app.command("get")
def topic_get(
    topic_id: str = typer.Argument(
        ...,
        help="The topic's component ID (GUID)",
    ),
    yaml_output: bool = typer.Option(
        False,
        "--yaml",
        "-y",
        help="Output topic content as YAML",
    ),
    output: str = typer.Option(
        None,
        "--output",
        "-o",
        help="Write YAML content to a file",
    ),
):
    """
    Get a topic by ID.

    Retrieves topic details including the YAML content that defines the conversation flow.

    Examples:
        copilot agent topic get <topic-id>
        copilot agent topic get <topic-id> --yaml
        copilot agent topic get <topic-id> --output my-topic.yaml
    """
    try:
        client = get_client()
        topic = client.get_topic(topic_id)

        content = topic.get("data", "")

        if output:
            # Write content to file
            with open(output, "w") as f:
                f.write(content)
            print_success(f"Topic content written to {output}")
        elif yaml_output:
            # Print just the YAML content
            if content:
                typer.echo(content)
            else:
                typer.echo("# No YAML content found for this topic")
        else:
            # Print full topic info as JSON
            print_json({
                "name": topic.get("name"),
                "component_id": topic.get("botcomponentid"),
                "schema_name": topic.get("schemaname"),
                "component_type": TOPIC_COMPONENT_TYPE_NAMES.get(topic.get("componenttype", 0), "unknown"),
                "status": topic.get("statecode@OData.Community.Display.V1.FormattedValue", "Active"),
                "is_managed": topic.get("ismanaged", False),
                "description": topic.get("description", ""),
                "content": content,
            })
    except Exception as e:
        exit_code = handle_api_error(e)
        raise typer.Exit(exit_code)


@topic_app.command("create")
def topic_create(
    agent_id: str = typer.Option(
        ...,
        "--agentId",
        "-a",
        help="The agent's unique identifier (GUID)",
    ),
    name: str = typer.Option(
        ...,
        "--name",
        "-n",
        help="Display name for the topic",
    ),
    file: str = typer.Option(
        None,
        "--file",
        "-f",
        help="Path to YAML file containing topic content",
    ),
    triggers: str = typer.Option(
        None,
        "--triggers",
        "-t",
        help="Comma-separated trigger phrases (for simple topics)",
    ),
    message: str = typer.Option(
        None,
        "--message",
        "-m",
        help="Response message (for simple topics)",
    ),
    description: str = typer.Option(
        None,
        "--description",
        "-d",
        help="Optional description for the topic",
    ),
):
    """
    Create a new topic for an agent.

    Topics can be created in two ways:
    1. From a YAML file with --file
    2. Using simple parameters (--triggers and --message) for basic topics

    Examples:
        # Create from YAML file
        copilot agent topic create --agentId <agent-id> --name "My Topic" --file topic.yaml

        # Create simple topic with triggers and message
        copilot agent topic create --agentId <agent-id> --name "Greeting" \\
            --triggers "hello,hi,hey there" --message "Hello! How can I help?"
    """
    try:
        client = get_client()

        # Determine topic content
        if file:
            # Read content from file
            try:
                with open(file, "r") as f:
                    content = f.read()
            except FileNotFoundError:
                print_error(f"File not found: {file}")
                raise typer.Exit(1)
            except Exception as e:
                print_error(f"Error reading file: {e}")
                raise typer.Exit(1)
        elif triggers and message:
            # Generate simple topic YAML
            trigger_list = [t.strip() for t in triggers.split(",")]
            content = client.generate_simple_topic_yaml(name, trigger_list, message)
        else:
            print_error("Must provide either --file or both --triggers and --message")
            raise typer.Exit(1)

        # Create the topic
        component_id = client.create_topic(
            bot_id=agent_id,
            name=name,
            content=content,
            description=description,
        )

        if component_id:
            print_success(f"Topic '{name}' created successfully.")
            typer.echo(f"Component ID: {component_id}")
        else:
            print_success(f"Topic '{name}' created successfully.")
    except Exception as e:
        exit_code = handle_api_error(e)
        raise typer.Exit(exit_code)


@topic_app.command("update")
def topic_update(
    topic_id: str = typer.Argument(
        ...,
        help="The topic's component ID (GUID)",
    ),
    name: str = typer.Option(
        None,
        "--name",
        "-n",
        help="New display name for the topic",
    ),
    file: str = typer.Option(
        None,
        "--file",
        "-f",
        help="Path to YAML file containing updated topic content",
    ),
    triggers: str = typer.Option(
        None,
        "--triggers",
        "-t",
        help="New comma-separated trigger phrases (replaces existing triggers)",
    ),
    message: str = typer.Option(
        None,
        "--message",
        "-m",
        help="New response message (replaces existing message)",
    ),
    description: str = typer.Option(
        None,
        "--description",
        "-d",
        help="New description for the topic",
    ),
):
    """
    Update an existing topic.

    You can update a topic's name, content, or description.
    Content can be updated from a YAML file or using simple parameters.

    Examples:
        # Update from YAML file
        copilot agent topic update <topic-id> --file updated-topic.yaml

        # Update topic name
        copilot agent topic update <topic-id> --name "New Name"

        # Update triggers and message
        copilot agent topic update <topic-id> --triggers "new phrase,another" --message "New response"

        # Update multiple fields
        copilot agent topic update <topic-id> --name "New Name" --description "Updated description"
    """
    try:
        client = get_client()

        # Get current topic for name and validation
        current_topic = client.get_topic(topic_id)
        topic_name = current_topic.get("name", topic_id)

        # Check if this is a system topic
        if current_topic.get("ismanaged", False):
            print_error(f"Cannot update system topic '{topic_name}'. System topics are read-only.")
            raise typer.Exit(1)

        # Determine content update
        content = None
        if file:
            # Read content from file
            try:
                with open(file, "r") as f:
                    content = f.read()
            except FileNotFoundError:
                print_error(f"File not found: {file}")
                raise typer.Exit(1)
            except Exception as e:
                print_error(f"Error reading file: {e}")
                raise typer.Exit(1)
        elif triggers or message:
            if not (triggers and message):
                print_error("When updating triggers/message, both --triggers and --message must be provided")
                raise typer.Exit(1)
            # Generate new simple topic YAML
            display_name = name or topic_name
            trigger_list = [t.strip() for t in triggers.split(",")]
            content = client.generate_simple_topic_yaml(display_name, trigger_list, message)

        # Check if any updates provided
        if not any([name, content, description]):
            print_error("No updates provided. Specify at least one field to update.")
            raise typer.Exit(1)

        # Update the topic
        client.update_topic(
            component_id=topic_id,
            name=name,
            content=content,
            description=description,
        )

        print_success(f"Topic '{topic_name}' updated successfully.")
    except Exception as e:
        exit_code = handle_api_error(e)
        raise typer.Exit(exit_code)


# Register topic subgroup
app.add_typer(topic_app, name="topic")


# =============================================================================
# Tool Commands (Agent Tools / Connected Agents)
# =============================================================================

tool_app = typer.Typer(help="Manage agent tools (connected agents)")


def get_tool_category(schema_name: str, data: str = "") -> str:
    """Determine the tool category from the schema name or data field."""
    # Combine schema_name and data for pattern matching
    # UI-created tools have action kind in data, API-created in schema name
    search_text = (schema_name or "") + " " + (data or "")

    if not search_text.strip():
        return "Unknown"

    # Check for known patterns
    if "InvokeConnectedAgentTaskAction" in search_text:
        return "Agent"
    elif "InvokeFlowTaskAction" in search_text:
        return "Flow"
    elif "InvokePromptTaskAction" in search_text:
        return "Prompt"
    elif "InvokeConnectorTaskAction" in search_text:
        return "Connector"
    elif "InvokeHttpTaskAction" in search_text:
        return "HTTP"
    elif "TaskAction" in search_text:
        # Generic task action - extract the type
        import re
        match = re.search(r'Invoke(\w+)TaskAction', search_text)
        if match:
            return match.group(1)
        return "Action"
    elif ".action." in (schema_name or "").lower():
        # UI-created action without clear type - mark as Action
        return "Action"
    else:
        return "Unknown"


def format_tool_for_display(tool: dict) -> dict:
    """Format an agent tool for display."""
    schema_name = tool.get("schemaname", "") or ""
    data = tool.get("data", "") or ""

    # Determine category from schema and data
    category = get_tool_category(schema_name, data)

    # Extract description and display name from data if available
    data = tool.get("data", "") or ""
    description = ""
    display_name = ""
    if data:
        # Extract the description and display name from YAML-like data
        lines = data.split("\n")
        for line in lines:
            if line.startswith("modelDescription:"):
                description = line.replace("modelDescription:", "").strip().strip('"')
                # Truncate long descriptions
                if len(description) > 80:
                    description = description[:77] + "..."
            elif line.startswith("modelDisplayName:"):
                display_name = line.replace("modelDisplayName:", "").strip().strip('"')

    return {
        "name": tool.get("name"),
        "display_name": display_name,
        "category": category,
        "component_id": tool.get("botcomponentid"),
        "description": description,
        "status": tool.get("statecode@OData.Community.Display.V1.FormattedValue", "Active"),
    }


@tool_app.command("list")
def tool_list(
    agent_id: str = typer.Option(
        ...,
        "--agentId",
        "-a",
        help="The agent's unique identifier (GUID)",
    ),
    category: Optional[str] = typer.Option(
        None,
        "--category",
        "-c",
        help="Filter by category: agent, flow, prompt, connector, http",
    ),
    table: bool = typer.Option(
        False,
        "--table",
        "-t",
        help="Display output as a formatted table instead of JSON",
    ),
):
    """
    List tools for an agent.

    Tools include connected agents, flows, prompts, connectors, and HTTP actions
    that the agent can invoke.

    Categories:
      - agent: Connected sub-agents (InvokeConnectedAgentTaskAction)
      - flow: Power Automate flows (InvokeFlowTaskAction)
      - prompt: AI prompts (InvokePromptTaskAction)
      - connector: Connector actions (InvokeConnectorTaskAction)
      - http: HTTP requests (InvokeHttpTaskAction)

    Examples:
        copilot agent tool list --agentId <agent-id>
        copilot agent tool list --agentId <agent-id> --table
        copilot agent tool list --agentId <agent-id> --category agent
    """
    try:
        client = get_client()
        tools = client.list_tools(agent_id, category=category)

        if not tools:
            typer.echo("No agent tools found for this agent.")
            return

        formatted = [format_tool_for_display(t) for t in tools]

        if table:
            print_table(
                formatted,
                columns=["name", "display_name", "category", "status", "component_id"],
                headers=["Name", "Display Name", "Category", "Status", "Component ID"],
            )
        else:
            print_json(formatted)
    except Exception as e:
        exit_code = handle_api_error(e)
        raise typer.Exit(exit_code)


@tool_app.command("get")
def tool_get(
    component_id: str = typer.Argument(
        ...,
        help="The tool component's unique identifier (GUID)",
    ),
    raw: bool = typer.Option(
        False,
        "--raw",
        "-r",
        help="Show raw component data without parsing",
    ),
    yaml_output: bool = typer.Option(
        False,
        "--yaml",
        "-y",
        help="Show the tool's YAML definition",
    ),
):
    """
    Get details of a specific tool.

    Shows comprehensive information about an agent tool including its
    configuration, inputs, outputs, and YAML definition.

    Examples:
        copilot agent tool get <component-id>
        copilot agent tool get <component-id> --yaml
        copilot agent tool get <component-id> --raw
    """
    import yaml as yaml_lib

    try:
        client = get_client()
        tool = client.get_tool(component_id)

        if raw:
            print_json(tool)
            return

        if yaml_output:
            data = tool.get("data", "")
            if data:
                typer.echo(data)
            else:
                typer.echo("No YAML definition found for this tool.")
            return

        # Parse the tool data for formatted output
        schema_name = tool.get("schemaname", "") or ""
        data = tool.get("data", "") or ""
        category = get_tool_category(schema_name, data)

        # Extract details from YAML data
        parsed_data = {}
        yaml_parse_error = None
        if data:
            try:
                parsed_data = yaml_lib.safe_load(data) or {}
            except Exception as e:
                yaml_parse_error = str(e)

        # Build display output - Basic Info Section
        typer.echo("=" * 60)
        typer.echo(f"Tool: {tool.get('name', 'Unknown')}")
        typer.echo("=" * 60)
        typer.echo(f"Component ID: {tool.get('botcomponentid', '')}")
        typer.echo(f"Category: {category}")
        typer.echo(f"Schema Name: {schema_name}")
        typer.echo(f"Status: {tool.get('statecode@OData.Community.Display.V1.FormattedValue', 'Active')}")

        # Show entity-level description if present
        entity_description = tool.get("description", "")
        if entity_description:
            typer.echo(f"Entity Description: {entity_description}")
        typer.echo("")

        # Display parsed YAML fields
        if parsed_data:
            typer.echo("--- Configuration ---")
            if parsed_data.get("modelDisplayName"):
                typer.echo(f"Display Name: {parsed_data.get('modelDisplayName')}")
            if parsed_data.get("modelDescription"):
                typer.echo(f"Description: {parsed_data.get('modelDescription')}")

            # Show availability settings
            availability = parsed_data.get("isAvailableForAgentInvocation")
            if availability is not None:
                typer.echo(f"Available for Agent: {availability}")

            # Show confirmation settings
            user_confirm = parsed_data.get("requiresUserConfirmation")
            if user_confirm is not None:
                typer.echo(f"Requires Confirmation: {user_confirm}")
            confirm_msg = parsed_data.get("userConfirmationText")
            if confirm_msg:
                typer.echo(f"Confirmation Message: {confirm_msg}")

            # Show inputs
            inputs = parsed_data.get("inputs") or []
            if inputs:
                typer.echo("")
                typer.echo("--- Inputs ---")
                for inp in inputs:
                    inp_name = inp.get("name", "unknown")
                    inp_type = inp.get("dataType", "unknown")
                    inp_required = inp.get("isRequired", False)
                    inp_desc = inp.get("description", "")
                    default_val = inp.get("defaultValue")
                    visible = inp.get("isVisible", True)
                    req_marker = " [required]" if inp_required else ""
                    vis_marker = " [hidden]" if not visible else ""
                    typer.echo(f"  {inp_name} ({inp_type}){req_marker}{vis_marker}")
                    if inp_desc:
                        typer.echo(f"    Description: {inp_desc}")
                    if default_val is not None:
                        typer.echo(f"    Default: {default_val}")

            # Show outputs (supports both 'name' and 'propertyName' formats)
            outputs = parsed_data.get("outputs") or []
            if outputs:
                typer.echo("")
                typer.echo("--- Outputs ---")
                for out in outputs:
                    out_name = out.get("name") or out.get("propertyName", "unknown")
                    out_type = out.get("dataType", "")
                    out_desc = out.get("description", "")
                    type_suffix = f" ({out_type})" if out_type else ""
                    typer.echo(f"  {out_name}{type_suffix}")
                    if out_desc:
                        typer.echo(f"    Description: {out_desc}")

        elif yaml_parse_error and data:
            # YAML couldn't be parsed, but try to extract key fields with regex
            import re
            typer.echo("--- Configuration ---")
            typer.echo("(Note: YAML data contains formatting issues)")
            # Try to extract modelDisplayName
            display_match = re.search(r'modelDisplayName:\s*(.+?)(?:\n|$)', data)
            if display_match:
                typer.echo(f"Display Name: {display_match.group(1).strip()}")
            # Try to extract modelDescription
            desc_match = re.search(r'modelDescription:\s*(.+?)(?:\noutputs:|$)', data, re.DOTALL)
            if desc_match:
                desc = desc_match.group(1).strip()
                if len(desc) > 200:
                    desc = desc[:200] + "..."
                typer.echo(f"Description: {desc}")

            # Try to extract outputs from raw YAML
            typer.echo("")
            typer.echo("--- Outputs ---")
            output_matches = re.findall(r'propertyName:\s*(\S+)', data)
            for out_name in output_matches:
                typer.echo(f"  {out_name}")

            # Try to extract action details
            typer.echo("")
            typer.echo("--- Action Details ---")
            kind_match = re.search(r'kind:\s*(\S+)', data)
            if kind_match and "TaskDialog" not in kind_match.group(1):
                typer.echo(f"Action Type: {kind_match.group(1)}")
            # Look for action kind specifically
            action_kind_match = re.search(r'action:\s*\n\s*kind:\s*(\S+)', data)
            if action_kind_match:
                typer.echo(f"Action Type: {action_kind_match.group(1)}")

            conn_ref_match = re.search(r'connectionReference:\s*(\S+)', data)
            if conn_ref_match:
                typer.echo(f"Connection Ref: {conn_ref_match.group(1)}")

            op_id_match = re.search(r'operationId:\s*(\S+)', data)
            if op_id_match:
                typer.echo(f"Operation ID: {op_id_match.group(1)}")

        # Show action-specific details - outside of if/elif for parsed data
        if parsed_data:
            # Support both 'actions' (list) and 'action' (single object) formats
            actions = parsed_data.get("actions") or []
            single_action = parsed_data.get("action")
            if single_action:
                actions = [single_action]
            if actions and len(actions) > 0:
                action = actions[0]  # Usually there's one main action
                action_kind = action.get("kind", "")
                typer.echo("")
                typer.echo("--- Action Details ---")
                typer.echo(f"Action Type: {action_kind}")

                # Connector-specific details
                if "Connector" in action_kind:
                    connector_id = action.get("connectorId", "")
                    operation_id = action.get("operationId", "")
                    # Support both connectionReferenceLogicalName and connectionReference
                    conn_ref = action.get("connectionReferenceLogicalName") or action.get("connectionReference", "")
                    if connector_id:
                        typer.echo(f"Connector ID: {connector_id}")
                    if operation_id:
                        typer.echo(f"Operation ID: {operation_id}")
                    if conn_ref:
                        typer.echo(f"Connection Ref: {conn_ref}")

                    # Show connection properties if present
                    conn_props = action.get("connectionProperties") or {}
                    if conn_props:
                        mode = conn_props.get("mode", "")
                        if mode:
                            typer.echo(f"Connection Mode: {mode}")

                    # Show input mappings if present
                    input_params = action.get("inputParameters") or {}
                    if input_params:
                        typer.echo("Input Mappings:")
                        for param_name, param_value in input_params.items():
                            typer.echo(f"  {param_name}: {param_value}")

                    # Show output mappings if present
                    output_params = action.get("outputParameters") or {}
                    if output_params:
                        typer.echo("Output Mappings:")
                        for param_name, param_value in output_params.items():
                            typer.echo(f"  {param_name}: {param_value}")

                # Agent-specific details
                elif "ConnectedAgent" in action_kind:
                    target_id = action.get("agentId", "")
                    if target_id:
                        typer.echo(f"Target Agent ID: {target_id}")
                    include_history = action.get("includeConversationHistory", False)
                    typer.echo(f"Include History: {include_history}")

                # Flow-specific details
                elif "Flow" in action_kind:
                    flow_id = action.get("flowId", "")
                    if flow_id:
                        typer.echo(f"Flow ID: {flow_id}")

                # HTTP-specific details
                elif "Http" in action_kind:
                    url = action.get("url", "")
                    method = action.get("method", "")
                    if url:
                        typer.echo(f"URL: {url}")
                    if method:
                        typer.echo(f"Method: {method}")

        # Show timestamps
        typer.echo("")
        typer.echo("--- Metadata ---")
        created = tool.get("createdon", "")
        modified = tool.get("modifiedon", "")
        if created:
            typer.echo(f"Created: {created}")
        if modified:
            typer.echo(f"Modified: {modified}")

        # Show parent bot info
        parent_bot = tool.get("_parentbotid_value", "")
        if parent_bot:
            typer.echo(f"Parent Bot: {parent_bot}")

    except Exception as e:
        exit_code = handle_api_error(e)
        raise typer.Exit(exit_code)


@tool_app.command("add")
def tool_add(
    agent_id: str = typer.Option(
        ...,
        "--agentId",
        "-a",
        help="The parent agent's unique identifier (GUID)",
    ),
    tool_type: str = typer.Option(
        ...,
        "--toolType",
        "-T",
        help="Tool type: connector, prompt, flow, http, agent",
    ),
    tool_id: str = typer.Option(
        ...,
        "--id",
        help="Tool identifier (format depends on tool type)",
    ),
    name: Optional[str] = typer.Option(
        None,
        "--name",
        "-n",
        help="Display name for the tool (auto-generated if not provided)",
    ),
    description: Optional[str] = typer.Option(
        None,
        "--description",
        "-d",
        help="Description of when to use this tool (for AI orchestration)",
    ),
    inputs: Optional[str] = typer.Option(
        None,
        "--inputs",
        help="Input parameters as JSON string",
    ),
    outputs: Optional[str] = typer.Option(
        None,
        "--outputs",
        help="Output parameters as JSON string",
    ),
    # Type-specific parameters
    connection_reference_id: Optional[str] = typer.Option(
        None,
        "--connection-reference-id",
        help="Connection reference ID (GUID) from 'copilot connection-references list' (required for connector tools)",
    ),
    no_history: bool = typer.Option(
        False,
        "--no-history",
        help="Don't pass conversation history (for agent tools)",
    ),
    method: str = typer.Option(
        "GET",
        "--method",
        help="HTTP method (for http tools)",
    ),
    headers_json: Optional[str] = typer.Option(
        None,
        "--headers",
        help="HTTP headers as JSON string (for http tools)",
    ),
    body: Optional[str] = typer.Option(
        None,
        "--body",
        help="Request body template (for http tools)",
    ),
    connection_mode: str = typer.Option(
        "Maker",
        "--connection-mode",
        help="Connection mode for connector tools: 'Invoker' (user auth) or 'Maker' (maker auth)",
    ),
    force: bool = typer.Option(
        False,
        "--force",
        "-f",
        help="Force adding tool even if operation has internal visibility (may not work correctly)",
    ),
):
    """
    Add a tool to an agent.

    Tool Types:
      - connector: Power Platform connector operation
      - prompt: AI Builder prompt
      - flow: Power Automate flow
      - http: Direct HTTP request
      - agent: Connected Copilot Studio agent

    Tool ID Format (--id):
      - connector: "connector_id:operation_id" (e.g., "shared_asana:GetTask")
      - prompt: Prompt GUID
      - flow: Flow GUID
      - http: URL
      - agent: Target agent GUID

    Examples:
        # Connector tool (requires connection reference ID)
        copilot agent tool add -a <agent-id> --toolType connector \\
            --id "shared_asana:GetTask" --connection-reference-id <conn-ref-id> --name "Get Task"

        # Prompt tool
        copilot agent tool add -a <agent-id> --toolType prompt \\
            --id <prompt-guid> --name "Summarize"

        # Flow tool
        copilot agent tool add -a <agent-id> --toolType flow \\
            --id <flow-guid> --name "Process Order"

        # HTTP tool
        copilot agent tool add -a <agent-id> --toolType http \\
            --id "https://api.example.com/data" --method POST

        # Connected agent tool
        copilot agent tool add -a <agent-id> --toolType agent \\
            --id <target-agent-id> --name "Expert Reviewer"
    """
    import json

    # Validate tool type
    valid_types = ['connector', 'prompt', 'flow', 'http', 'agent']
    if tool_type.lower() not in valid_types:
        typer.echo(f"Error: Invalid tool type '{tool_type}'. Must be one of: {', '.join(valid_types)}", err=True)
        raise typer.Exit(1)

    # Validate connection mode
    valid_modes = ['Invoker', 'Maker']
    if connection_mode not in valid_modes:
        typer.echo(f"Error: Invalid connection mode '{connection_mode}'. Must be one of: {', '.join(valid_modes)}", err=True)
        raise typer.Exit(1)

    # Parse JSON parameters
    inputs_dict = None
    if inputs:
        try:
            inputs_dict = json.loads(inputs)
        except json.JSONDecodeError as e:
            typer.echo(f"Error: Invalid JSON for --inputs: {e}", err=True)
            raise typer.Exit(1)

    outputs_dict = None
    if outputs:
        try:
            outputs_dict = json.loads(outputs)
        except json.JSONDecodeError as e:
            typer.echo(f"Error: Invalid JSON for --outputs: {e}", err=True)
            raise typer.Exit(1)

    headers_dict = None
    if headers_json:
        try:
            headers_dict = json.loads(headers_json)
        except json.JSONDecodeError as e:
            typer.echo(f"Error: Invalid JSON for --headers: {e}", err=True)
            raise typer.Exit(1)

    try:
        client = get_client()

        component_id = client.add_tool(
            bot_id=agent_id,
            tool_type=tool_type,
            tool_id=tool_id,
            name=name,
            description=description,
            inputs=inputs_dict,
            outputs=outputs_dict,
            connection_reference_id=connection_reference_id,
            connection_mode=connection_mode,
            no_history=no_history,
            method=method,
            headers=headers_dict,
            body=body,
            force=force,
        )

        if component_id:
            print_success(f"{tool_type.capitalize()} tool created successfully!")
            typer.echo(f"Component ID: {component_id}")
            typer.echo("")
            typer.echo("Note: You may need to publish the agent for changes to take effect.")
        else:
            typer.echo("Tool created but component ID could not be extracted.")
    except Exception as e:
        exit_code = handle_api_error(e)
        raise typer.Exit(exit_code)


@tool_app.command("remove")
def tool_remove(
    component_id: str = typer.Argument(
        ...,
        help="The tool component's unique identifier (GUID)",
    ),
    force: bool = typer.Option(
        False,
        "--force",
        "-f",
        help="Skip confirmation prompt",
    ),
):
    """
    Remove a tool from an agent.

    Examples:
        copilot agent tool remove <component-id>
        copilot agent tool remove <component-id> --force
    """
    if not force:
        confirm = typer.confirm(f"Are you sure you want to remove tool {component_id}?")
        if not confirm:
            typer.echo("Operation cancelled.")
            raise typer.Exit(0)

    try:
        client = get_client()
        client.remove_tool(component_id)
        print_success(f"Tool {component_id} removed successfully.")
    except Exception as e:
        exit_code = handle_api_error(e)
        raise typer.Exit(exit_code)


@tool_app.command("update")
def tool_update(
    component_id: str = typer.Argument(
        ...,
        help="The tool component's unique identifier (GUID)",
    ),
    name: Optional[str] = typer.Option(
        None,
        "--name",
        "-n",
        help="New display name for the tool",
    ),
    description: Optional[str] = typer.Option(
        None,
        "--description",
        "-d",
        help="New description for the tool (used by AI for orchestration, max 1024 chars)",
    ),
    availability: Optional[bool] = typer.Option(
        None,
        "--available/--not-available",
        help="Allow agent to use tool dynamically (--available) or only from topics (--not-available)",
    ),
    confirmation: Optional[bool] = typer.Option(
        None,
        "--confirm/--no-confirm",
        help="Ask user for confirmation before running the tool",
    ),
    confirmation_message: Optional[str] = typer.Option(
        None,
        "--confirm-message",
        "-m",
        help="Custom message to show when asking for confirmation",
    ),
    inputs: Optional[str] = typer.Option(
        None,
        "--inputs",
        "-i",
        help='Input default values as JSON, e.g., \'{"workspace": "123", "projects": "456"}\'',
    ),
):
    """
    Update a tool's attributes.

    The description field is especially important as it's used by the AI agent
    to determine when to use this tool. Make it descriptive and explicit about
    when the tool should be used.

    Tool Availability:
      --available      Agent may use this tool at any time (generative orchestration)
      --not-available  Only use when explicitly referenced by topics or agents

    User Confirmation:
      --confirm        Ask end user for approval before running
      --no-confirm     Run without asking (default)
      --confirm-message  Custom confirmation prompt text

    Input Defaults:
      --inputs         Set default values for tool inputs as JSON

    Examples:
        # Update name and description
        copilot agent tool update <component-id> --name "New Tool Name"
        copilot agent tool update <component-id> --description "Use this tool when..."

        # Configure availability
        copilot agent tool update <component-id> --available      # Allow dynamic use
        copilot agent tool update <component-id> --not-available  # Only from topics

        # Configure user confirmation
        copilot agent tool update <component-id> --confirm        # Enable confirmation
        copilot agent tool update <component-id> --no-confirm     # Disable confirmation
        copilot agent tool update <component-id> --confirm --confirm-message "Proceed with action?"

        # Set input default values
        copilot agent tool update <component-id> --inputs '{"workspace": "123456", "projects": "789012"}'

        # Combined update
        copilot agent tool update <component-id> -n "Name" -d "Description" --available --confirm
    """
    if not any([name, description, availability is not None, confirmation is not None, confirmation_message, inputs]):
        typer.echo("Error: At least one option must be provided.", err=True)
        typer.echo("Options: --name, --description, --available/--not-available, --confirm/--no-confirm, --confirm-message, --inputs")
        raise typer.Exit(1)

    # Validate description length
    if description and len(description) > 1024:
        typer.echo(f"Error: Description exceeds 1024 character limit ({len(description)} chars).", err=True)
        raise typer.Exit(1)

    # Parse inputs JSON if provided
    inputs_dict = None
    if inputs:
        try:
            inputs_dict = json.loads(inputs)
            if not isinstance(inputs_dict, dict):
                typer.echo("Error: --inputs must be a JSON object (dict)", err=True)
                raise typer.Exit(1)
        except json.JSONDecodeError as e:
            typer.echo(f"Error: Invalid JSON for --inputs: {e}", err=True)
            raise typer.Exit(1)

    try:
        client = get_client()
        result = client.update_tool(
            component_id=component_id,
            name=name,
            description=description,
            availability=availability,
            confirmation=confirmation,
            confirmation_message=confirmation_message,
            inputs=inputs_dict,
        )
        print_success(f"Tool updated successfully!")
        typer.echo(f"Name: {result.get('name', 'N/A')}")
        if result.get('description'):
            desc = result['description']
            if len(desc) > 100:
                desc = desc[:100] + "..."
            typer.echo(f"Description: {desc}")

        # Show additional settings if they were updated
        data = result.get('data', '')
        if availability is not None:
            status = "Available for dynamic use" if availability else "Only available from topics"
            typer.echo(f"Availability: {status}")
        if confirmation is not None or confirmation_message:
            if 'confirmation:' in data:
                typer.echo(f"User Confirmation: Enabled")
            else:
                typer.echo(f"User Confirmation: Disabled")
        if inputs_dict:
            typer.echo(f"Input defaults updated: {', '.join(inputs_dict.keys())}")
    except Exception as e:
        exit_code = handle_api_error(e)
        raise typer.Exit(exit_code)


# Register tool subgroup
app.add_typer(tool_app, name="tool")


# =============================================================================
# Analytics (Application Insights) Commands
# =============================================================================

analytics_app = typer.Typer(help="Manage Application Insights telemetry for agents")


@analytics_app.command("get")
def analytics_get(
    agent_id: str = typer.Argument(..., help="The agent's unique identifier (GUID)"),
):
    """
    Get Application Insights configuration for an agent.

    Shows the current App Insights connection string and logging settings.

    Examples:
        copilot agent analytics get fcef595a-30bb-f011-bbd3-000d3a8ba54e
    """
    try:
        client = get_client()

        # Get agent name for display
        bot = client.get_bot(agent_id)
        agent_name = bot.get("name", agent_id)

        config = client.get_bot_app_insights(agent_id)

        typer.echo(f"\nApplication Insights for '{agent_name}':\n")

        if config["enabled"]:
            typer.echo(f"  Status:                   Enabled")
            # Mask connection string for security (show only first 20 chars)
            conn_str = config["connectionString"]
            masked = conn_str[:40] + "..." if len(conn_str) > 40 else conn_str
            typer.echo(f"  Connection String:        {masked}")
        else:
            typer.echo(f"  Status:                   Not configured")

        typer.echo(f"  Log Activities:           {config['logActivities']}")
        typer.echo(f"  Log Sensitive Properties: {config['logSensitiveProperties']}")
        typer.echo("")

    except Exception as e:
        exit_code = handle_api_error(e)
        raise typer.Exit(exit_code)


@analytics_app.command("enable")
def analytics_enable(
    agent_id: str = typer.Argument(..., help="The agent's unique identifier (GUID)"),
    connection_string: str = typer.Option(
        ...,
        "--connection-string",
        "-c",
        help="App Insights connection string (from Azure portal)",
    ),
    log_activities: bool = typer.Option(
        False,
        "--log-activities",
        "-l",
        help="Enable logging of incoming/outgoing messages and events",
    ),
    log_sensitive: bool = typer.Option(
        False,
        "--log-sensitive",
        "-s",
        help="Enable logging of sensitive properties (userid, name, text, speak)",
    ),
):
    """
    Enable Application Insights telemetry for an agent.

    Configures the agent to send telemetry to an existing App Insights instance.
    Multiple agents can share the same App Insights instance.

    The connection string can be found in your Azure Application Insights
    resource under Settings > Properties or in the Overview section.

    Examples:
        copilot agent analytics enable <agent-id> -c "InstrumentationKey=xxx;..."
        copilot agent analytics enable <agent-id> -c "..." --log-activities
        copilot agent analytics enable <agent-id> -c "..." --log-activities --log-sensitive
    """
    try:
        client = get_client()

        # Get agent name for display
        bot = client.get_bot(agent_id)
        agent_name = bot.get("name", agent_id)

        typer.echo(f"Enabling Application Insights for '{agent_name}'...")

        client.update_bot_app_insights(
            bot_id=agent_id,
            connection_string=connection_string,
            log_activities=log_activities,
            log_sensitive_properties=log_sensitive,
        )

        print_success(f"Application Insights enabled for '{agent_name}'!")
        typer.echo("")
        typer.echo("Settings applied:")
        typer.echo(f"  Log Activities:           {log_activities}")
        typer.echo(f"  Log Sensitive Properties: {log_sensitive}")
        typer.echo("")
        typer.echo("Note: Telemetry data will appear in your App Insights Logs section.")
        typer.echo("      You may need to republish the agent for changes to take effect.")

    except Exception as e:
        exit_code = handle_api_error(e)
        raise typer.Exit(exit_code)


@analytics_app.command("disable")
def analytics_disable(
    agent_id: str = typer.Argument(..., help="The agent's unique identifier (GUID)"),
    force: bool = typer.Option(
        False,
        "--force",
        "-f",
        help="Skip confirmation prompt",
    ),
):
    """
    Disable Application Insights telemetry for an agent.

    Removes the App Insights connection string and disables all logging.

    Examples:
        copilot agent analytics disable <agent-id>
        copilot agent analytics disable <agent-id> --force
    """
    try:
        client = get_client()

        # Get agent name for display
        bot = client.get_bot(agent_id)
        agent_name = bot.get("name", agent_id)

        if not force:
            confirm = typer.confirm(
                f"Are you sure you want to disable Application Insights for '{agent_name}'?"
            )
            if not confirm:
                typer.echo("Operation cancelled.")
                raise typer.Exit(0)

        typer.echo(f"Disabling Application Insights for '{agent_name}'...")

        client.update_bot_app_insights(bot_id=agent_id, disable=True)

        print_success(f"Application Insights disabled for '{agent_name}'.")
        typer.echo("")
        typer.echo("Note: You may need to republish the agent for changes to take effect.")

    except Exception as e:
        exit_code = handle_api_error(e)
        raise typer.Exit(exit_code)


@analytics_app.command("update")
def analytics_update(
    agent_id: str = typer.Argument(..., help="The agent's unique identifier (GUID)"),
    log_activities: Optional[bool] = typer.Option(
        None,
        "--log-activities/--no-log-activities",
        help="Enable/disable logging of messages and events",
    ),
    log_sensitive: Optional[bool] = typer.Option(
        None,
        "--log-sensitive/--no-log-sensitive",
        help="Enable/disable logging of sensitive properties",
    ),
):
    """
    Update Application Insights logging options for an agent.

    Use this to change logging settings without modifying the connection string.

    Examples:
        copilot agent analytics update <agent-id> --log-activities
        copilot agent analytics update <agent-id> --no-log-activities
        copilot agent analytics update <agent-id> --log-sensitive
        copilot agent analytics update <agent-id> --log-activities --log-sensitive
    """
    if log_activities is None and log_sensitive is None:
        typer.echo("Error: Please specify at least one option to update.")
        typer.echo("Use --log-activities/--no-log-activities or --log-sensitive/--no-log-sensitive")
        raise typer.Exit(1)

    try:
        client = get_client()

        # Get agent name for display
        bot = client.get_bot(agent_id)
        agent_name = bot.get("name", agent_id)

        typer.echo(f"Updating Application Insights settings for '{agent_name}'...")

        client.update_bot_app_insights(
            bot_id=agent_id,
            log_activities=log_activities,
            log_sensitive_properties=log_sensitive,
        )

        print_success(f"Application Insights settings updated for '{agent_name}'!")

        # Show what was updated
        updates = []
        if log_activities is not None:
            updates.append(f"Log Activities: {log_activities}")
        if log_sensitive is not None:
            updates.append(f"Log Sensitive Properties: {log_sensitive}")

        if updates:
            typer.echo("")
            typer.echo("Updated settings:")
            for update in updates:
                typer.echo(f"  {update}")

    except Exception as e:
        exit_code = handle_api_error(e)
        raise typer.Exit(exit_code)


def _convert_timespan(timespan: str) -> str:
    """
    Convert user-friendly timespan to ISO 8601 duration.

    Examples:
        1h → PT1H
        24h → PT24H
        7d → P7D
        30d → P30D
    """
    timespan = timespan.lower().strip()

    # Already ISO 8601 format
    if timespan.startswith("p"):
        return timespan.upper()

    # Parse number and unit
    import re
    match = re.match(r"^(\d+)([hd])$", timespan)
    if not match:
        raise ValueError(f"Invalid timespan format: {timespan}. Use format like '24h' or '7d'")

    value = match.group(1)
    unit = match.group(2)

    if unit == "h":
        return f"PT{value}H"
    elif unit == "d":
        return f"P{value}D"

    raise ValueError(f"Unknown time unit: {unit}")


@analytics_app.command("query")
def analytics_query(
    agent_id: str = typer.Argument(..., help="The agent's unique identifier (GUID)"),
    timespan: str = typer.Option(
        "24h",
        "--timespan",
        "-t",
        help="Time range to query (e.g., 1h, 24h, 7d, 30d)",
    ),
    events_only: bool = typer.Option(
        False,
        "--events",
        "-e",
        help="Query only customEvents table (faster)",
    ),
    json_output: bool = typer.Option(
        False,
        "--json",
        help="Output raw JSON response",
    ),
    limit: int = typer.Option(
        100,
        "--limit",
        "-l",
        help="Maximum number of rows to display",
    ),
):
    """
    Query Application Insights telemetry for an agent.

    Retrieves telemetry data from the Application Insights instance
    configured for this agent. Requires App Insights to be enabled.

    Examples:
        copilot agent analytics query <agent-id>
        copilot agent analytics query <agent-id> --timespan 7d
        copilot agent analytics query <agent-id> --events --json
        copilot agent analytics query <agent-id> -t 1h -l 50
    """
    try:
        # Convert timespan to ISO 8601
        try:
            iso_timespan = _convert_timespan(timespan)
        except ValueError as e:
            typer.echo(f"Error: {e}")
            raise typer.Exit(1)

        client = get_client()

        # Get agent name for display
        bot = client.get_bot(agent_id)
        agent_name = bot.get("name", agent_id)

        typer.echo(f"Querying Application Insights for '{agent_name}'...")
        typer.echo(f"Time range: {timespan}")
        typer.echo("")

        # Execute query
        result = client.get_bot_telemetry(
            bot_id=agent_id,
            timespan=iso_timespan,
            events_only=events_only,
        )

        # Handle JSON output
        if json_output:
            print_json(result)
            return

        # Parse and display results
        tables = result.get("tables", [])
        if not tables:
            typer.echo("No telemetry data found for the specified time range.")
            return

        table = tables[0]
        columns = [col["name"] for col in table.get("columns", [])]
        rows = table.get("rows", [])

        if not rows:
            typer.echo("No telemetry data found for the specified time range.")
            return

        typer.echo(f"Found {len(rows)} records (showing up to {limit}):")
        typer.echo("")

        # Display as formatted output
        displayed = 0
        for row in rows:
            if displayed >= limit:
                typer.echo(f"\n... and {len(rows) - limit} more records. Use --limit to see more.")
                break

            # Create a dict for this row
            row_data = dict(zip(columns, row))

            timestamp = row_data.get("timestamp", "")
            if timestamp:
                # Format timestamp for display
                timestamp = timestamp.replace("T", " ").split(".")[0]

            table_name = row_data.get("_table", "event")
            name = row_data.get("name", "")
            message = row_data.get("message", "")

            # Format the line
            line = f"[{timestamp}] [{table_name}]"
            if name:
                line += f" {name}"
            if message:
                line += f": {message}"

            typer.echo(line)

            # Show custom dimensions if present (condensed)
            custom_dims = row_data.get("customDimensions")
            if custom_dims and isinstance(custom_dims, dict):
                # Show key fields from customDimensions
                key_fields = ["TopicName", "Kind", "text", "channelId", "fromName"]
                dim_parts = []
                for field in key_fields:
                    if field in custom_dims and custom_dims[field]:
                        dim_parts.append(f"{field}={custom_dims[field]}")
                if dim_parts:
                    typer.echo(f"    {', '.join(dim_parts)}")

            displayed += 1

        typer.echo("")
        print_success(f"Query complete. Retrieved {len(rows)} records.")

    except Exception as e:
        exit_code = handle_api_error(e)
        raise typer.Exit(exit_code)


# Register analytics subgroup
app.add_typer(analytics_app, name="analytics")


# =============================================================================
# Authentication Commands
# =============================================================================

auth_app = typer.Typer(help="Manage agent authentication configuration")

# Authentication mode mapping for display
AUTH_MODE_NAMES = {
    1: "None",
    2: "Integrated",
    3: "Custom Azure AD",
}


@auth_app.command("get")
def auth_get(
    agent_id: str = typer.Argument(..., help="The agent's unique identifier (GUID)"),
):
    """
    Get authentication configuration for an agent.

    Shows the current authentication mode and settings.

    Authentication Modes:
      - 1 = None (no authentication required)
      - 2 = Integrated (Microsoft Entra ID integrated)
      - 3 = Custom Azure AD (manual configuration)

    Examples:
        copilot agent auth get fcef595a-30bb-f011-bbd3-000d3a8ba54e
    """
    try:
        client = get_client()

        # Get agent name for display
        bot = client.get_bot(agent_id)
        agent_name = bot.get("name", agent_id)

        auth_config = client.get_bot_auth(agent_id)

        typer.echo(f"\nAuthentication for '{agent_name}':\n")
        typer.echo(f"  Mode:    {auth_config['mode']} ({auth_config['mode_name']})")
        typer.echo(f"  Trigger: {auth_config['trigger']} ({auth_config['trigger_name']})")

        if auth_config.get("configuration"):
            typer.echo(f"  Config:  {auth_config['configuration']}")

        typer.echo("")

    except Exception as e:
        exit_code = handle_api_error(e)
        raise typer.Exit(exit_code)


@auth_app.command("set")
def auth_set(
    agent_id: str = typer.Argument(..., help="The agent's unique identifier (GUID)"),
    mode: Optional[int] = typer.Option(
        None,
        "--mode",
        "-m",
        help="Authentication mode: 1=None, 2=Integrated, 3=Custom Azure AD",
    ),
    trigger: Optional[int] = typer.Option(
        None,
        "--trigger",
        help="Authentication trigger: 0=As Needed, 1=Always",
    ),
):
    """
    Set authentication mode and/or trigger for an agent.

    Authentication Modes:
      - 1 = None (no authentication required)
      - 2 = Integrated (Microsoft Entra ID integrated - default for new agents)
      - 3 = Custom Azure AD (manual Microsoft Entra ID configuration)

    Authentication Triggers:
      - 0 = As Needed (authenticate only when required)
      - 1 = Always (require authentication for all conversations)

    Examples:
        copilot agent auth set <agent-id> --mode 1
        copilot agent auth set <agent-id> --mode 1 --trigger 0
        copilot agent auth set <agent-id> --trigger 0
    """
    try:
        if mode is None and trigger is None:
            typer.echo("Error: Must specify at least --mode or --trigger", err=True)
            raise typer.Exit(1)

        if mode is not None and mode not in AUTH_MODE_NAMES:
            typer.echo(f"Error: Invalid mode {mode}. Valid modes: 1=None, 2=Integrated, 3=Custom Azure AD", err=True)
            raise typer.Exit(1)

        if trigger is not None and trigger not in (0, 1):
            typer.echo(f"Error: Invalid trigger {trigger}. Valid triggers: 0=As Needed, 1=Always", err=True)
            raise typer.Exit(1)

        client = get_client()

        # Get agent name for display
        bot = client.get_bot(agent_id)
        agent_name = bot.get("name", agent_id)

        updates = []
        if mode is not None:
            updates.append(f"mode to {mode} ({AUTH_MODE_NAMES[mode]})")
        if trigger is not None:
            trigger_name = "As Needed" if trigger == 0 else "Always"
            updates.append(f"trigger to {trigger} ({trigger_name})")

        typer.echo(f"Setting authentication for '{agent_name}': {', '.join(updates)}...")

        client.update_bot_auth(bot_id=agent_id, mode=mode, trigger=trigger)

        print_success(f"Authentication updated for '{agent_name}'!")
        typer.echo("")
        typer.echo("Note: You may need to republish the agent for changes to take effect.")

    except Exception as e:
        exit_code = handle_api_error(e)
        raise typer.Exit(exit_code)


@auth_app.command("list")
def auth_list(
    table: bool = typer.Option(
        False,
        "--table",
        "-t",
        help="Display output as a formatted table instead of JSON",
    ),
):
    """
    List authentication modes for all agents.

    Shows the authentication mode for each agent in the environment.

    Examples:
        copilot agent auth list
        copilot agent auth list --table
    """
    try:
        client = get_client()
        bots = client.list_bots(
            select=["name", "botid", "authenticationmode", "statecode"]
        )

        if not bots:
            typer.echo("No agents found.")
            return

        # Format for display
        formatted = []
        for bot in bots:
            auth_mode = bot.get("authenticationmode", 2)
            formatted.append({
                "name": bot.get("name"),
                "bot_id": bot.get("botid"),
                "auth_mode": auth_mode,
                "auth_mode_name": AUTH_MODE_NAMES.get(auth_mode, f"Unknown({auth_mode})"),
            })

        if table:
            print_table(
                formatted,
                columns=["name", "auth_mode", "auth_mode_name", "bot_id"],
                headers=["Name", "Mode", "Mode Name", "Bot ID"],
            )
        else:
            print_json(formatted)

    except Exception as e:
        exit_code = handle_api_error(e)
        raise typer.Exit(exit_code)


# Register auth subgroup
app.add_typer(auth_app, name="auth")
