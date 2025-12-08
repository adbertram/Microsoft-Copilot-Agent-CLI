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

    Returns agents with their bot_id, name, schema name, and status.

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
                headers=["Name", "Bot ID", "State", "Status"],
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
    bot_id: str = typer.Argument(..., help="The bot's unique identifier (GUID)"),
    include_components: bool = typer.Option(
        False,
        "--components",
        "-c",
        help="Include bot components (topics, triggers, etc.)",
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
        bot = client.get_bot(bot_id)

        if include_components:
            components = client.get_bot_components(bot_id)
            bot["components"] = components

        print_json(bot)
    except Exception as e:
        exit_code = handle_api_error(e)
        raise typer.Exit(exit_code)


@app.command("remove")
def remove_agent(
    bot_id: str = typer.Argument(..., help="The bot's unique identifier (GUID)"),
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

        # Get bot details first to show name in confirmation
        bot = client.get_bot(bot_id)
        bot_name = bot.get("name", bot_id)

        if not force:
            confirm = typer.confirm(f"Are you sure you want to delete agent '{bot_name}'?")
            if not confirm:
                typer.echo("Aborted.")
                raise typer.Exit(0)

        client.delete_bot(bot_id)
        print_success(f"Agent '{bot_name}' deleted successfully.")
    except Exception as e:
        exit_code = handle_api_error(e)
        raise typer.Exit(exit_code)


@app.command("publish")
def publish_agent(
    bot_id: str = typer.Argument(..., help="The bot's unique identifier (GUID)"),
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

        # Get bot details first to show name
        bot = client.get_bot(bot_id)
        bot_name = bot.get("name", bot_id)

        typer.echo(f"Publishing agent '{bot_name}'...")

        result = client.publish_bot(bot_id)

        if result.get("status") == "success":
            print_success(f"Agent '{bot_name}' published successfully!")
            if result.get("PublishedBotContentId"):
                typer.echo(f"Published Content ID: {result['PublishedBotContentId']}")
        else:
            typer.echo(f"Publish completed with status: {result}")
    except Exception as e:
        exit_code = handle_api_error(e)
        raise typer.Exit(exit_code)


@app.command("update")
def update_agent(
    bot_id: str = typer.Argument(..., help="The bot's unique identifier (GUID)"),
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
        copilot agent update <bot-id> --name "New Name"
        copilot agent update <bot-id> --description "New description"
        copilot agent update <bot-id> --instructions "New system prompt"
        copilot agent update <bot-id> --instructions-file ./prompt.txt
        copilot agent update <bot-id> --no-orchestration
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

        # Get current bot name for success message
        current_bot = client.get_bot(bot_id)
        bot_name = name if name else current_bot.get("name", bot_id)

        client.update_bot(
            bot_id=bot_id,
            name=name,
            instructions=agent_instructions,
            description=description,
            orchestration=orchestration,
        )

        print_success(f"Agent '{bot_name}' updated successfully.")
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
    bot_id: str = typer.Argument(
        ...,
        help="The bot's unique identifier (GUID)",
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
        help="Bot token endpoint URL (from Copilot Studio > Channels > Mobile app)",
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
    - {BOT_SCHEMA_NAME}: Your agent's schema name (e.g., cr83c_myAgent)

    Examples:
        # Using Direct Line secret
        copilot agent prompt <bot-id> --message "Hello" --secret "your-secret"

        # Using Entra ID authentication
        copilot agent prompt <bot-id> -m "Hello" --entra-id \\
            --client-id <app-client-id> --tenant-id <tenant-id> \\
            --token-endpoint "https://{ENV}.environment.api.powerplatform.com/powervirtualagents/botsbyschema/{BOT}/directline/token?api-version=2022-03-01-preview"

        # With file attachment
        copilot agent prompt <bot-id> -m "Review this" --file ./draft.docx --secret "xxx"

    Environment Variables:
        DIRECTLINE_SECRET - Direct Line secret (alternative to --secret)
        ENTRA_CLIENT_ID - Entra ID client ID (alternative to --client-id)
        ENTRA_TENANT_ID - Entra ID tenant ID (alternative to --tenant-id)
        ENTRA_SCOPE - OAuth scope (default: https://api.powerplatform.com/.default)
        BOT_TOKEN_ENDPOINT - Bot token endpoint (alternative to --token-endpoint)
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
            bot_token_endpoint = token_endpoint or os.environ.get("BOT_TOKEN_ENDPOINT")

            if not entra_client_id:
                typer.echo("Error: --client-id or ENTRA_CLIENT_ID env var required for Entra ID auth", err=True)
                raise typer.Exit(1)
            if not entra_tenant_id:
                typer.echo("Error: --tenant-id or ENTRA_TENANT_ID env var required for Entra ID auth", err=True)
                raise typer.Exit(1)
            if not bot_token_endpoint:
                typer.echo("Error: --token-endpoint or BOT_TOKEN_ENDPOINT env var required for Entra ID auth", err=True)
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
                    bot_token_endpoint,
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
            typer.echo(f"Starting conversation with agent {bot_id}...")

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
# Usage: copilot agent knowledge list --bot <bot-id>
#        copilot agent knowledge file add --bot <bot-id> ...
#        copilot agent knowledge azure-ai-search add --bot <bot-id> ...

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
    bot_id: str = typer.Option(
        ...,
        "--bot",
        "-b",
        help="The bot's unique identifier (GUID)",
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
        copilot agent knowledge list --bot <bot-id>
        copilot agent knowledge list --bot <bot-id> --table
        copilot agent knowledge list --bot <bot-id> --type file
    """
    try:
        client = get_client()
        sources = client.list_knowledge_sources(bot_id, source_type=source_type)

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
    bot_id: str = typer.Option(
        ...,
        "--bot",
        "-b",
        help="The bot's unique identifier (GUID)",
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
        copilot agent knowledge remove --bot <bot-id> <component-id>
        copilot agent knowledge remove --bot <bot-id> <component-id> --force
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
    bot_id: str = typer.Option(
        ...,
        "--bot",
        "-b",
        help="The bot's unique identifier (GUID)",
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
        copilot agent knowledge file add --bot <bot-id> --name "FAQ" --content "Q: What? A: Test."
        copilot agent knowledge file add --bot <bot-id> --name "Guide" --file ./guide.md
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
            bot_id=bot_id,
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
    bot_id: str = typer.Option(
        ...,
        "--bot",
        "-b",
        help="The bot's unique identifier (GUID)",
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

    WARNING: This command creates a bot component record but the knowledge source
    may not appear in Copilot Studio UI. Copilot Studio requires a Power Platform
    connection to be properly linked, which involves internal configuration not
    exposed via public APIs.

    RECOMMENDED APPROACH:
    1. Use 'copilot connection create' to create the Power Platform connection
    2. Link the connection to your agent via the Copilot Studio UI

    Examples:
        copilot agent knowledge azure-ai-search add --bot <bot-id> \\
            --name "Product Docs" \\
            --endpoint https://mysearch.search.windows.net \\
            --index products-index \\
            --api-key <your-api-key>
    """
    try:
        client = get_client()
        component_id = client.add_azure_ai_search_knowledge_source(
            bot_id=bot_id,
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
# Connection Commands
# =============================================================================

connection_app = typer.Typer(help="Manage Power Platform connections for Copilot Studio")


@connection_app.command("create")
def connection_create(
    name: str = typer.Option(
        ...,
        "--name",
        "-n",
        help="Display name for the connection",
    ),
    endpoint: str = typer.Option(
        ...,
        "--endpoint",
        "-e",
        help="Azure AI Search endpoint URL (e.g., https://mysearch.search.windows.net)",
    ),
    api_key: str = typer.Option(
        ...,
        "--api-key",
        "-k",
        help="Azure AI Search API key (admin or query key)",
    ),
    environment: str = typer.Option(
        ...,
        "--environment",
        "--env",
        help="Power Platform environment ID (e.g., Default-<tenant-id>)",
    ),
):
    """
    Create a Power Platform connection for Azure AI Search.

    This creates a connection that can be used by Copilot Studio agents to access
    Azure AI Search indexes as knowledge sources. After creating the connection,
    you must link it to your agent through the Copilot Studio UI.

    To find your environment ID:
    - Go to make.powerapps.com
    - Select your environment
    - The ID is in the URL or can be found via 'az rest' commands

    Examples:
        copilot connection create \\
            --name "My Search Connection" \\
            --endpoint https://mysearch.search.windows.net \\
            --api-key <your-api-key> \\
            --environment Default-<tenant-id>

    After creation:
    1. Open your agent in Copilot Studio
    2. Go to Knowledge > Add knowledge
    3. Select Azure AI Search
    4. Select the connection you just created
    5. Specify the index name
    """
    try:
        client = get_client()
        result = client.create_azure_ai_search_connection(
            connection_name=name,
            search_endpoint=endpoint,
            api_key=api_key,
            environment_id=environment,
        )

        connection_id = result.get("name", "")
        display_name = result.get("properties", {}).get("displayName", name)
        statuses = result.get("properties", {}).get("statuses", [])
        status = statuses[0].get("status", "Unknown") if statuses else "Unknown"

        print_success(f"Connection '{display_name}' created successfully.")
        typer.echo(f"Connection ID: {connection_id}")
        typer.echo(f"Status: {status}")
        typer.echo("")
        typer.echo("Next steps:")
        typer.echo("  1. Open your agent in Copilot Studio")
        typer.echo("  2. Go to Knowledge > Add knowledge")
        typer.echo("  3. Select Azure AI Search")
        typer.echo("  4. Select this connection and specify the index name")
    except Exception as e:
        exit_code = handle_api_error(e)
        raise typer.Exit(exit_code)


@connection_app.command("list")
def connection_list(
    environment: str = typer.Option(
        ...,
        "--environment",
        "--env",
        help="Power Platform environment ID (e.g., Default-<tenant-id>)",
    ),
    table: bool = typer.Option(
        False,
        "--table",
        "-t",
        help="Display output as a formatted table instead of JSON",
    ),
):
    """
    List Azure AI Search connections in a Power Platform environment.

    Examples:
        copilot connection list --environment Default-<tenant-id>
        copilot connection list --env Default-<tenant-id> --table
    """
    try:
        client = get_client()
        connections = client.list_azure_ai_search_connections(environment)

        if table:
            # Format for table display
            formatted = []
            for conn in connections:
                props = conn.get("properties", {})
                statuses = props.get("statuses", [])
                status = statuses[0].get("status", "Unknown") if statuses else "Unknown"
                formatted.append({
                    "Name": props.get("displayName", ""),
                    "Connection ID": conn.get("name", ""),
                    "Status": status,
                    "Created": props.get("createdTime", "")[:10] if props.get("createdTime") else "",
                })
            print_table(
                formatted,
                columns=["Name", "Connection ID", "Status", "Created"],
            )
        else:
            print_json(connections)
    except Exception as e:
        exit_code = handle_api_error(e)
        raise typer.Exit(exit_code)


@connection_app.command("delete")
def connection_delete(
    connection_id: str = typer.Argument(
        ...,
        help="The connection's unique identifier (GUID)",
    ),
    environment: str = typer.Option(
        ...,
        "--environment",
        "--env",
        help="Power Platform environment ID (e.g., Default-<tenant-id>)",
    ),
    force: bool = typer.Option(
        False,
        "--force",
        "-f",
        help="Skip confirmation prompt",
    ),
):
    """
    Delete a Power Platform connection.

    Examples:
        copilot connection delete <connection-id> --environment Default-<tenant-id>
        copilot connection delete <connection-id> --env Default-<tenant-id> --force
    """
    if not force:
        confirm = typer.confirm(f"Are you sure you want to delete connection {connection_id}?")
        if not confirm:
            raise typer.Abort()

    try:
        client = get_client()
        client.delete_connection(connection_id, environment)
        print_success(f"Connection {connection_id} deleted successfully.")
    except Exception as e:
        exit_code = handle_api_error(e)
        raise typer.Exit(exit_code)


# Register connection subgroup
app.add_typer(connection_app, name="connection")


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
    bot: Optional[str] = typer.Option(
        None,
        "--bot",
        "-b",
        help="Filter by bot name or ID",
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

    Shows recent conversation transcripts, optionally filtered by bot name or ID.

    Examples:
        copilot agent transcript list
        copilot agent transcript list --table
        copilot agent transcript list --bot "Writer Draft Reviewer" --limit 10
        copilot agent transcript list --bot d2735b5c-aecb-f011-bbd3-000d3a8ba54e
    """
    try:
        client = get_client()

        # Determine if bot is an ID or name
        bot_id = None
        bot_name = None
        if bot:
            if _is_guid(bot):
                bot_id = bot
            else:
                bot_name = bot

        transcripts = client.list_transcripts(bot_id=bot_id, bot_name=bot_name, limit=limit)

        if not transcripts:
            typer.echo("No transcripts found.")
            return

        if table:
            formatted = [format_transcript_for_display(t) for t in transcripts]
            print_table(
                formatted,
                columns=["id", "bot_name", "start_time"],
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


def get_tool_category(schema_name: str) -> str:
    """Determine the tool category from the schema name."""
    if not schema_name:
        return "Unknown"

    # Check for known patterns in schema name
    if "InvokeConnectedAgentTaskAction" in schema_name:
        return "Agent"
    elif "InvokeFlowTaskAction" in schema_name:
        return "Flow"
    elif "InvokePromptTaskAction" in schema_name:
        return "Prompt"
    elif "InvokeConnectorTaskAction" in schema_name or ".connector." in schema_name.lower():
        return "Connector"
    elif "InvokeHttpTaskAction" in schema_name:
        return "HTTP"
    elif "TaskAction" in schema_name:
        # Generic task action - extract the type
        import re
        match = re.search(r'Invoke(\w+)TaskAction', schema_name)
        if match:
            return match.group(1)
        return "Action"
    else:
        return "Unknown"


def format_tool_for_display(tool: dict) -> dict:
    """Format an agent tool for display."""
    schema_name = tool.get("schemaname", "") or ""

    # Determine category from schema
    category = get_tool_category(schema_name)

    # Extract description from data if available
    data = tool.get("data", "") or ""
    description = ""
    if "modelDescription:" in data:
        # Extract the description from YAML-like data
        lines = data.split("\n")
        for line in lines:
            if line.startswith("modelDescription:"):
                description = line.replace("modelDescription:", "").strip().strip('"')
                # Truncate long descriptions
                if len(description) > 80:
                    description = description[:77] + "..."
                break

    return {
        "name": tool.get("name"),
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
                columns=["name", "category", "status", "component_id"],
                headers=["Name", "Category", "Status", "Component ID"],
            )
        else:
            print_json(formatted)
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
    target_agent_id: str = typer.Option(
        ...,
        "--target",
        "-t",
        help="The target agent's unique identifier (GUID) to connect as a tool",
    ),
    name: Optional[str] = typer.Option(
        None,
        "--name",
        "-n",
        help="Display name for the tool (defaults to target agent's name)",
    ),
    description: Optional[str] = typer.Option(
        None,
        "--description",
        "-d",
        help="Description of when to use this tool (for AI orchestration)",
    ),
    no_history: bool = typer.Option(
        False,
        "--no-history",
        help="Don't pass conversation history to the connected agent",
    ),
):
    """
    Add a connected agent as a tool.

    Creates an InvokeConnectedAgentTaskAction that allows this agent to
    invoke another Copilot Studio agent as a sub-agent.

    The target agent must:
      - Be in the same environment
      - Be published
      - Have "Let other agents connect" enabled in settings

    Examples:
        copilot agent tool add --agentId <parent-id> --target <child-id>
        copilot agent tool add -a <parent-id> -t <child-id> --name "Expert Reviewer"
        copilot agent tool add -a <parent-id> -t <child-id> --no-history
    """
    try:
        client = get_client()
        component_id = client.add_connected_agent_tool(
            bot_id=agent_id,
            target_bot_id=target_agent_id,
            name=name,
            description=description,
            pass_conversation_history=not no_history,
        )

        if component_id:
            print_success(f"Connected agent tool created successfully!")
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


# Register tool subgroup
app.add_typer(tool_app, name="tool")


# =============================================================================
# Analytics (Application Insights) Commands
# =============================================================================

analytics_app = typer.Typer(help="Manage Application Insights telemetry for agents")


@analytics_app.command("get")
def analytics_get(
    bot_id: str = typer.Argument(..., help="The bot's unique identifier (GUID)"),
):
    """
    Get Application Insights configuration for an agent.

    Shows the current App Insights connection string and logging settings.

    Examples:
        copilot agent analytics get fcef595a-30bb-f011-bbd3-000d3a8ba54e
    """
    try:
        client = get_client()

        # Get bot name for display
        bot = client.get_bot(bot_id)
        bot_name = bot.get("name", bot_id)

        config = client.get_bot_app_insights(bot_id)

        typer.echo(f"\nApplication Insights for '{bot_name}':\n")

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
    bot_id: str = typer.Argument(..., help="The bot's unique identifier (GUID)"),
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
        copilot agent analytics enable <bot-id> -c "InstrumentationKey=xxx;..."
        copilot agent analytics enable <bot-id> -c "..." --log-activities
        copilot agent analytics enable <bot-id> -c "..." --log-activities --log-sensitive
    """
    try:
        client = get_client()

        # Get bot name for display
        bot = client.get_bot(bot_id)
        bot_name = bot.get("name", bot_id)

        typer.echo(f"Enabling Application Insights for '{bot_name}'...")

        client.update_bot_app_insights(
            bot_id=bot_id,
            connection_string=connection_string,
            log_activities=log_activities,
            log_sensitive_properties=log_sensitive,
        )

        print_success(f"Application Insights enabled for '{bot_name}'!")
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
    bot_id: str = typer.Argument(..., help="The bot's unique identifier (GUID)"),
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
        copilot agent analytics disable <bot-id>
        copilot agent analytics disable <bot-id> --force
    """
    try:
        client = get_client()

        # Get bot name for display
        bot = client.get_bot(bot_id)
        bot_name = bot.get("name", bot_id)

        if not force:
            confirm = typer.confirm(
                f"Are you sure you want to disable Application Insights for '{bot_name}'?"
            )
            if not confirm:
                typer.echo("Operation cancelled.")
                raise typer.Exit(0)

        typer.echo(f"Disabling Application Insights for '{bot_name}'...")

        client.update_bot_app_insights(bot_id=bot_id, disable=True)

        print_success(f"Application Insights disabled for '{bot_name}'.")
        typer.echo("")
        typer.echo("Note: You may need to republish the agent for changes to take effect.")

    except Exception as e:
        exit_code = handle_api_error(e)
        raise typer.Exit(exit_code)


@analytics_app.command("update")
def analytics_update(
    bot_id: str = typer.Argument(..., help="The bot's unique identifier (GUID)"),
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
        copilot agent analytics update <bot-id> --log-activities
        copilot agent analytics update <bot-id> --no-log-activities
        copilot agent analytics update <bot-id> --log-sensitive
        copilot agent analytics update <bot-id> --log-activities --log-sensitive
    """
    if log_activities is None and log_sensitive is None:
        typer.echo("Error: Please specify at least one option to update.")
        typer.echo("Use --log-activities/--no-log-activities or --log-sensitive/--no-log-sensitive")
        raise typer.Exit(1)

    try:
        client = get_client()

        # Get bot name for display
        bot = client.get_bot(bot_id)
        bot_name = bot.get("name", bot_id)

        typer.echo(f"Updating Application Insights settings for '{bot_name}'...")

        client.update_bot_app_insights(
            bot_id=bot_id,
            log_activities=log_activities,
            log_sensitive_properties=log_sensitive,
        )

        print_success(f"Application Insights settings updated for '{bot_name}'!")

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
    bot_id: str = typer.Argument(..., help="The bot's unique identifier (GUID)"),
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
        copilot agent analytics query <bot-id>
        copilot agent analytics query <bot-id> --timespan 7d
        copilot agent analytics query <bot-id> --events --json
        copilot agent analytics query <bot-id> -t 1h -l 50
    """
    try:
        # Convert timespan to ISO 8601
        try:
            iso_timespan = _convert_timespan(timespan)
        except ValueError as e:
            typer.echo(f"Error: {e}")
            raise typer.Exit(1)

        client = get_client()

        # Get bot name for display
        bot = client.get_bot(bot_id)
        bot_name = bot.get("name", bot_id)

        typer.echo(f"Querying Application Insights for '{bot_name}'...")
        typer.echo(f"Time range: {timespan}")
        typer.echo("")

        # Execute query
        result = client.get_bot_telemetry(
            bot_id=bot_id,
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
    bot_id: str = typer.Argument(..., help="The bot's unique identifier (GUID)"),
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

        # Get bot name for display
        bot = client.get_bot(bot_id)
        bot_name = bot.get("name", bot_id)

        auth_config = client.get_bot_auth(bot_id)

        typer.echo(f"\nAuthentication for '{bot_name}':\n")
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
    bot_id: str = typer.Argument(..., help="The bot's unique identifier (GUID)"),
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
        copilot agent auth set <bot-id> --mode 1
        copilot agent auth set <bot-id> --mode 1 --trigger 0
        copilot agent auth set <bot-id> --trigger 0
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

        # Get bot name for display
        bot = client.get_bot(bot_id)
        bot_name = bot.get("name", bot_id)

        updates = []
        if mode is not None:
            updates.append(f"mode to {mode} ({AUTH_MODE_NAMES[mode]})")
        if trigger is not None:
            trigger_name = "As Needed" if trigger == 0 else "Always"
            updates.append(f"trigger to {trigger} ({trigger_name})")

        typer.echo(f"Setting authentication for '{bot_name}': {', '.join(updates)}...")

        client.update_bot_auth(bot_id=bot_id, mode=mode, trigger=trigger)

        print_success(f"Authentication updated for '{bot_name}'!")
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
