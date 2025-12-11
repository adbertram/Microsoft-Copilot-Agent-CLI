"""Output formatting and error handling for CopilotAgent CLI."""
import json
import sys
from typing import Any


def print_json(data: Any, indent: int = 2):
    """
    Print data as formatted JSON to stdout.

    Args:
        data: Data to output as JSON
        indent: JSON indentation level (default: 2)
    """
    try:
        json_str = json.dumps(data, indent=indent, ensure_ascii=False)
        print(json_str)
    except (TypeError, ValueError) as e:
        print_error(f"Failed to serialize data to JSON: {e}")
        sys.exit(1)


def print_table(data: list[dict], columns: list[str], headers: list[str] = None):
    """
    Print data as a formatted table.

    Args:
        data: List of dictionaries to display
        columns: List of column keys to display
        headers: Optional list of header names (defaults to column keys)
    """
    if not data:
        print("No results found.")
        return

    if headers is None:
        headers = columns

    # Calculate column widths
    widths = []
    for i, col in enumerate(columns):
        header_width = len(headers[i])
        max_data_width = max(len(str(row.get(col, ""))) for row in data)
        widths.append(max(header_width, max_data_width))

    # Print header
    header_row = "  ".join(h.ljust(w) for h, w in zip(headers, widths))
    print(header_row)
    print("-" * len(header_row))

    # Print data rows
    for row in data:
        values = [str(row.get(col, "")).ljust(w) for col, w in zip(columns, widths)]
        print("  ".join(values))


def print_error(message: str):
    """
    Print error message to stderr.

    Args:
        message: Error message to print
    """
    print(f"Error: {message}", file=sys.stderr)


def print_warning(message: str):
    """
    Print warning message to stderr in yellow.

    Args:
        message: Warning message to print
    """
    # ANSI escape code for yellow text
    yellow = "\033[93m"
    reset = "\033[0m"
    print(f"{yellow}Warning: {message}{reset}", file=sys.stderr)


def print_success(message: str):
    """
    Print success message to stderr (keeps stdout clean for JSON).

    Args:
        message: Success message to print
    """
    print(f"✓ {message}", file=sys.stderr)


def handle_api_error(error: Exception) -> int:
    """
    Handle API errors and return appropriate exit code.

    Args:
        error: Exception from API call

    Returns:
        int: Exit code (1 for general errors, 2 for auth errors)
    """
    from .client import ClientError

    error_str = str(error)

    # For ClientError, always show the full message as it's intentional
    if isinstance(error, ClientError):
        print_error(error_str)
        return 1

    # Check for authentication errors
    if "401" in error_str or "unauthorized" in error_str.lower():
        print_error(
            "Authentication failed. Please check your credentials or run 'az login'."
        )
        return 2

    # Check for not found errors (only for HTTP 404, not general "not found" text)
    if "404" in error_str:
        print_error("Resource not found.")
        return 1

    # Check for permission errors
    if "403" in error_str or "forbidden" in error_str.lower():
        print_error("Permission denied. Check that you have access to this resource.")
        return 1

    # Check for rate limiting
    if "429" in error_str or "rate limit" in error_str.lower():
        print_error("Rate limit exceeded. Please try again later.")
        return 1

    # Check for validation errors
    if "400" in error_str or "bad request" in error_str.lower():
        print_error(f"Invalid request: {error_str}")
        return 1

    # Generic error
    print_error(f"API error: {error_str}")
    return 1


def format_bot_for_display(bot: dict) -> dict:
    """
    Format an agent record for display.

    Args:
        bot: Raw agent record from Dataverse

    Returns:
        Simplified agent record for display
    """
    return {
        "name": bot.get("name", ""),
        "botid": bot.get("botid", ""),  # Keep API field name for compatibility
        "schemaname": bot.get("schemaname", ""),
        "statecode": bot.get("statecode@OData.Community.Display.V1.FormattedValue", bot.get("statecode", "")),
        "statuscode": bot.get("statuscode@OData.Community.Display.V1.FormattedValue", bot.get("statuscode", "")),
        "createdon": bot.get("createdon", ""),
        "modifiedon": bot.get("modifiedon", ""),
    }


def format_transcript_content(content: str) -> str:
    """
    Parse and format transcript JSON content for human readability.

    The transcript content is a JSON string containing conversation activities.
    This function extracts the messages and formats them as a readable conversation.

    Args:
        content: JSON string containing transcript activities

    Returns:
        Formatted conversation string
    """
    from datetime import datetime

    if not content:
        return "(No content)"

    try:
        data = json.loads(content)
    except json.JSONDecodeError:
        return f"(Unable to parse content: {content[:200]}...)"

    lines = []

    # Handle different transcript formats
    activities = []
    if isinstance(data, list):
        activities = data
    elif isinstance(data, dict):
        activities = data.get("activities", data.get("value", []))
        if not activities and "text" in data:
            # Single message format
            activities = [data]

    for activity in activities:
        if not isinstance(activity, dict):
            continue

        # Get activity type
        activity_type = activity.get("type", "")

        # Skip non-message activities (trace, event, etc.)
        if activity_type != "message":
            continue

        # Get the message text
        text = activity.get("text", "")
        if not text:
            # Try to extract from attachments or other fields
            attachments = activity.get("attachments", [])
            if attachments and isinstance(attachments, list):
                for att in attachments:
                    if isinstance(att, dict) and att.get("content"):
                        content_obj = att.get("content", {})
                        if isinstance(content_obj, dict):
                            text = content_obj.get("text", str(content_obj))
                        else:
                            text = str(content_obj)
                        break

        if not text:
            continue

        # Get sender info - role can be integer (0=bot, 1=user) or string
        from_info = activity.get("from", {})
        sender_role = from_info.get("role", "unknown")

        # Map role to display name (handle both int and string formats)
        if sender_role == 1 or sender_role == "user":
            display_sender = "User"
        elif sender_role == 0 or sender_role == "bot":
            display_sender = "Agent"
        else:
            display_sender = str(sender_role) if sender_role else "Unknown"

        # Get timestamp - can be Unix timestamp (int) or ISO string
        timestamp = activity.get("timestamp", "")
        time_display = ""
        if timestamp:
            if isinstance(timestamp, int):
                # Unix timestamp - convert to readable time
                try:
                    dt = datetime.fromtimestamp(timestamp)
                    time_display = dt.strftime("%H:%M:%S")
                except (ValueError, OSError):
                    time_display = str(timestamp)
            elif isinstance(timestamp, str) and "T" in timestamp:
                # ISO timestamp - extract time portion
                time_part = timestamp.split("T")[1]
                if "." in time_part:
                    time_part = time_part.split(".")[0]
                elif "Z" in time_part:
                    time_part = time_part.replace("Z", "")
                time_display = time_part

        # Format the message line
        if time_display:
            lines.append(f"[{time_display}] {display_sender}: {text}")
        else:
            lines.append(f"{display_sender}: {text}")

    if not lines:
        return "(No messages found in transcript)"

    return "\n".join(lines)


def format_transcript_for_display(transcript: dict) -> dict:
    """
    Format a transcript record for table display.

    Args:
        transcript: Raw transcript record from Dataverse

    Returns:
        Simplified transcript record for display
    """
    # Get agent name from OData formatted value annotation, fall back to ID
    agent_name = transcript.get(
        "_bot_conversationtranscriptid_value@OData.Community.Display.V1.FormattedValue",
        transcript.get("_bot_conversationtranscriptid_value", ""),
    )

    # Format start time for readability (remove T and Z if present)
    start_time = transcript.get("conversationstarttime", "")
    if start_time:
        start_time = start_time.replace("T", " ").replace("Z", "")

    return {
        "id": transcript.get("conversationtranscriptid", ""),
        "name": transcript.get("name", ""),
        "agent_name": agent_name,
        "agent_id": transcript.get("_bot_conversationtranscriptid_value", ""),
        "start_time": start_time,
        "schema_type": transcript.get("schematype", ""),
    }
