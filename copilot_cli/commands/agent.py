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

        # Handle file content extraction (same for both API flows)
        combined_message = message
        if file:
            file_path = Path(file)
            if not file_path.exists():
                typer.echo(f"Error: File not found: {file}", err=True)
                raise typer.Exit(1)

            file_name = file_path.name
            ext = file_path.suffix.lower()

            # Extract text content from the file
            extracted_text = None

            # Text-based files - read directly
            if ext in (".txt", ".md", ".json", ".xml", ".html", ".csv", ".yaml", ".yml"):
                try:
                    with open(file_path, "r", encoding="utf-8") as f:
                        extracted_text = f.read()
                    if verbose:
                        typer.echo(f"Read text file: {file_name} ({len(extracted_text)} characters)")
                except UnicodeDecodeError:
                    typer.echo(f"Warning: Could not decode {file_name} as UTF-8", err=True)
                except IOError as e:
                    typer.echo(f"Error reading file: {e}", err=True)
                    raise typer.Exit(1)

            # Word documents - extract text using python-docx
            elif ext == ".docx":
                try:
                    from docx import Document
                    doc = Document(file_path)
                    paragraphs = [p.text for p in doc.paragraphs if p.text.strip()]
                    extracted_text = "\n\n".join(paragraphs)
                    if verbose:
                        typer.echo(f"Extracted text from Word doc: {file_name} ({len(extracted_text)} characters)")
                except ImportError:
                    typer.echo("Error: python-docx is required for .docx files. Install with: pip install python-docx", err=True)
                    raise typer.Exit(1)
                except Exception as e:
                    typer.echo(f"Error reading Word document: {e}", err=True)
                    raise typer.Exit(1)

            # PDF files - extract text using PyPDF2 or pdfplumber
            elif ext == ".pdf":
                try:
                    try:
                        import pdfplumber
                        with pdfplumber.open(file_path) as pdf:
                            pages_text = [page.extract_text() or "" for page in pdf.pages]
                            extracted_text = "\n\n".join(pages_text)
                        if verbose:
                            typer.echo(f"Extracted text from PDF: {file_name} ({len(extracted_text)} characters)")
                    except ImportError:
                        try:
                            from PyPDF2 import PdfReader
                            reader = PdfReader(file_path)
                            pages_text = [page.extract_text() or "" for page in reader.pages]
                            extracted_text = "\n\n".join(pages_text)
                            if verbose:
                                typer.echo(f"Extracted text from PDF: {file_name} ({len(extracted_text)} characters)")
                        except ImportError:
                            typer.echo("Error: pdfplumber or PyPDF2 is required for .pdf files.", err=True)
                            typer.echo("Install with: pip install pdfplumber", err=True)
                            raise typer.Exit(1)
                except Exception as e:
                    typer.echo(f"Error reading PDF: {e}", err=True)
                    raise typer.Exit(1)

            else:
                typer.echo(f"Error: Unsupported file type: {ext}", err=True)
                typer.echo("Supported types: .txt, .md, .json, .xml, .html, .csv, .yaml, .docx, .pdf", err=True)
                raise typer.Exit(1)

            if not extracted_text or not extracted_text.strip():
                typer.echo(f"Error: No text content could be extracted from {file_name}", err=True)
                raise typer.Exit(1)

            # Combine message with file content
            combined_message = f"{message}\n\n---\n\n**File: {file_name}**\n\n{extracted_text}"
            if verbose:
                typer.echo(f"Combined message ({len(combined_message)} characters)")

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

            # Step 4: Send message
            if verbose:
                typer.echo(f"Sending message: \"{message}\"")

            send_payload = {
                "type": "message",
                "from": {"id": user_id, "name": "Copilot CLI"},
                "text": combined_message,
            }

            send_response = client.post(
                f"{DIRECTLINE_URL}/conversations/{conv_id}/activities",
                headers={
                    "Authorization": f"Bearer {directline_token}",
                    "Content-Type": "application/json",
                },
                json=send_payload,
            )

            if send_response.status_code != 200:
                typer.echo(f"Error: Failed to send message (HTTP {send_response.status_code})", err=True)
                if verbose:
                    typer.echo(f"Response: {send_response.text}", err=True)
                raise typer.Exit(1)

            activity_id = send_response.json().get("id")
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
    is_system = topic.get("_is_system", False)

    return {
        "name": topic.get("name"),
        "type": "System" if is_system else "Custom",
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
    system_only: bool = typer.Option(
        False,
        "--system",
        "-s",
        help="Show only system topics",
    ),
    custom_only: bool = typer.Option(
        False,
        "--custom",
        "-c",
        help="Show only custom topics",
    ),
    table: bool = typer.Option(
        False,
        "--table",
        "-t",
        help="Display output as a formatted table instead of JSON",
    ),
):
    """
    List topics for an agent.

    Returns both system and custom topics by default. Use --system or --custom
    to filter to a specific type.

    Examples:
        copilot agent topic list --agentId <agent-id>
        copilot agent topic list --agentId <agent-id> --table
        copilot agent topic list --agentId <agent-id> --system
        copilot agent topic list --agentId <agent-id> --custom --table
    """
    try:
        # Determine filter settings
        include_system = True
        include_custom = True

        if system_only and custom_only:
            typer.echo("Error: Cannot specify both --system and --custom", err=True)
            raise typer.Exit(1)

        if system_only:
            include_custom = False
        elif custom_only:
            include_system = False

        client = get_client()
        topics = client.list_topics(
            agent_id,
            include_system=include_system,
            include_custom=include_custom,
        )

        if not topics:
            filter_desc = ""
            if system_only:
                filter_desc = " system"
            elif custom_only:
                filter_desc = " custom"
            typer.echo(f"No{filter_desc} topics found for this agent.")
            return

        formatted = [format_topic_for_display(t) for t in topics]

        if table:
            print_table(
                formatted,
                columns=["name", "type", "component_type", "status", "component_id"],
                headers=["Name", "Type", "Component Type", "Status", "Component ID"],
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


# Register topic subgroup
app.add_typer(topic_app, name="topic")
