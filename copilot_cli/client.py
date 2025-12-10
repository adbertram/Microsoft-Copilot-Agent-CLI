"""Dataverse API client for Copilot Studio agents."""
import subprocess
import json
import re
import random
import string
from typing import Optional, Any
import httpx
from .config import get_config


def parse_connection_string(connection_string: str) -> dict[str, str]:
    """
    Parse an Application Insights connection string into a dictionary.

    Connection strings have the format:
    InstrumentationKey=xxx;IngestionEndpoint=xxx;LiveEndpoint=xxx;ApplicationId=xxx

    Args:
        connection_string: The App Insights connection string

    Returns:
        Dict with keys like 'InstrumentationKey', 'ApplicationId', etc.
    """
    result = {}
    if not connection_string:
        return result

    for part in connection_string.split(";"):
        if "=" in part:
            key, value = part.split("=", 1)
            result[key.strip()] = value.strip()

    return result


class ClientError(Exception):
    """Exception raised for client initialization or API errors."""
    pass


class DataverseClient:
    """Client for interacting with Dataverse Web API for Copilot Studio agents."""

    def __init__(self, base_url: str, access_token: str):
        """
        Initialize the Dataverse client.

        Args:
            base_url: Dataverse environment URL (e.g., https://org1cb52429.crm.dynamics.com)
            access_token: OAuth access token for authentication
        """
        self.base_url = base_url.rstrip("/")
        self.api_url = f"{self.base_url}/api/data/v9.2"
        self.access_token = access_token
        self._http_client = httpx.Client(timeout=30.0)

    def _get_headers(self) -> dict[str, str]:
        """Get HTTP headers for API requests."""
        return {
            "Authorization": f"Bearer {self.access_token}",
            "OData-MaxVersion": "4.0",
            "OData-Version": "4.0",
            "Accept": "application/json",
            "Content-Type": "application/json",
            "Prefer": "odata.include-annotations=*",
        }

    def _request(self, method: str, endpoint: str, **kwargs) -> Any:
        """
        Make an HTTP request to the Dataverse API.

        Args:
            method: HTTP method (GET, POST, PATCH, DELETE)
            endpoint: API endpoint (relative to api_url)
            **kwargs: Additional arguments to pass to httpx

        Returns:
            Response data as dict/list

        Raises:
            ClientError: If the request fails
        """
        url = f"{self.api_url}/{endpoint.lstrip('/')}"
        headers = self._get_headers()
        headers.update(kwargs.pop("headers", {}))

        try:
            response = self._http_client.request(method, url, headers=headers, **kwargs)
            response.raise_for_status()

            if response.status_code == 204:
                return None

            return response.json()
        except httpx.HTTPStatusError as e:
            error_detail = ""
            try:
                error_body = e.response.json()
                if "error" in error_body:
                    error_detail = error_body["error"].get("message", str(error_body))
            except Exception:
                error_detail = e.response.text[:500] if e.response.text else str(e)
            raise ClientError(f"HTTP {e.response.status_code}: {error_detail}")
        except httpx.RequestError as e:
            raise ClientError(f"Request failed: {e}")

    def get(self, endpoint: str, params: Optional[dict] = None) -> Any:
        """Make a GET request."""
        return self._request("GET", endpoint, params=params)

    def post(self, endpoint: str, data: dict) -> Any:
        """Make a POST request."""
        return self._request("POST", endpoint, json=data)

    def patch(self, endpoint: str, data: dict) -> Any:
        """Make a PATCH request."""
        return self._request("PATCH", endpoint, json=data)

    def delete(self, endpoint: str, verify: bool = True) -> None:
        """
        Make a DELETE request.

        Args:
            endpoint: API endpoint to delete
            verify: If True, verify the resource was actually deleted by
                    attempting to GET it after deletion. Raises ClientError
                    if resource still exists.
        """
        self._request("DELETE", endpoint)

        if verify:
            try:
                self._request("GET", endpoint)
                # If we get here, resource still exists - deletion failed
                raise ClientError(f"Delete failed: resource still exists at {endpoint}")
            except ClientError as e:
                # 404 means successfully deleted, re-raise other errors
                if "404" not in str(e):
                    raise

    def list_bots(self, select: Optional[list[str]] = None) -> list[dict]:
        """
        List all Copilot Studio agents (bots) in the environment.

        Args:
            select: Optional list of fields to select

        Returns:
            List of bot records
        """
        endpoint = "bots"
        if select:
            endpoint += f"?$select={','.join(select)}"
        result = self.get(endpoint)
        return result.get("value", [])

    def get_bot(self, bot_id: str) -> dict:
        """
        Get a specific bot by ID.

        Args:
            bot_id: The bot's unique identifier

        Returns:
            Bot record
        """
        return self.get(f"bots({bot_id})")

    def get_bot_by_name(self, name: str) -> Optional[dict]:
        """
        Get a bot by its display name.

        Args:
            name: The bot's display name (case-insensitive search)

        Returns:
            Bot record if found, None otherwise
        """
        # Use contains for flexible matching
        result = self.get(f"bots?$filter=contains(name,'{name}')&$select=botid,name")
        bots = result.get("value", [])

        # Try exact match first
        for bot in bots:
            if bot.get("name", "").lower() == name.lower():
                return bot

        # Return first match if no exact match
        return bots[0] if bots else None

    def get_bot_components(self, bot_id: str) -> list[dict]:
        """
        Get components for a specific bot.

        Args:
            bot_id: The bot's unique identifier

        Returns:
            List of bot component records
        """
        result = self.get(f"botcomponents?$filter=_parentbotid_value eq {bot_id}")
        return result.get("value", [])

    def list_topics(
        self,
        bot_id: str,
        include_tools: bool = False,
        system_only: bool = False,
        custom_only: bool = False,
    ) -> list[dict]:
        """
        List topics for a specific bot.

        Args:
            bot_id: The bot's unique identifier
            include_tools: If False (default), filters out agent tools (InvokeConnectedAgentTaskAction)
            system_only: If True, only return system topics (ismanaged=true)
            custom_only: If True, only return custom topics (ismanaged=false)

        Returns:
            List of topic component records

        Note:
            Topic component types:
            - 0 = Topic (legacy)
            - 9 = Topic (V2)

            System vs Custom topics:
            - System topics (ismanaged=true): Built-in topics from managed solutions
            - Custom topics (ismanaged=false): User-created topics

            Agent tools have schema names containing 'InvokeConnectedAgentTaskAction'
            and data starting with 'kind: TaskDialog'. These are filtered out by default.
        """
        # Build filter
        filters = [
            f"_parentbotid_value eq {bot_id}",
            "(componenttype eq 0 or componenttype eq 9)"
        ]

        # Add managed status filter
        if system_only:
            filters.append("ismanaged eq true")
        elif custom_only:
            filters.append("ismanaged eq false")

        filter_str = " and ".join(filters)
        result = self.get(f"botcomponents?$filter={filter_str}&$orderby=name")
        topics = result.get("value", [])

        if not include_tools:
            # Filter out agent tools (InvokeConnectedAgentTaskAction components)
            topics = [
                t for t in topics
                if "InvokeConnectedAgentTaskAction" not in (t.get("schemaname") or "")
            ]

        return topics

    def list_tools(self, bot_id: str = None, category: str = None) -> list[dict]:
        """
        List tools, optionally filtered by bot.

        Args:
            bot_id: Optional bot's unique identifier. If None, lists all tools across all agents.
            category: Optional filter by category ('agent', 'flow', 'prompt', 'connector', 'http')

        Returns:
            List of tool component records

        Note:
            Tools are Topic (V2) components with schema names containing 'TaskAction'.
            Categories are determined by the TaskAction type:
            - Agent: InvokeConnectedAgentTaskAction
            - Flow: InvokeFlowTaskAction
            - Prompt: InvokePromptTaskAction
            - Connector: InvokeConnectorTaskAction
            - HTTP: InvokeHttpTaskAction
        """
        # Build filter - always require componenttype eq 9 (Topic V2)
        if bot_id:
            filter_clause = f"_parentbotid_value eq {bot_id} and componenttype eq 9"
        else:
            filter_clause = "componenttype eq 9"

        result = self.get(
            f"botcomponents?$filter={filter_clause}&$orderby=name"
        )
        components = result.get("value", [])

        # Filter to only tools
        # Tools can be identified by:
        # 1. Schema name containing "TaskAction" (API-created tools)
        # 2. Schema name containing ".action." (UI-created tools)
        # 3. Data containing "kind: TaskDialog" (all tools)
        tools = []
        for t in components:
            schema_name = t.get("schemaname") or ""
            data = t.get("data") or ""
            if ("TaskAction" in schema_name or
                ".action." in schema_name or
                "kind: TaskDialog" in data):
                tools.append(t)

        # Apply category filter if specified
        if category:
            category_patterns = {
                "agent": "InvokeConnectedAgentTaskAction",
                "flow": "InvokeFlowTaskAction",
                "prompt": "InvokePromptTaskAction",
                "connector": "InvokeConnectorTaskAction",
                "http": "InvokeHttpTaskAction",
            }
            pattern = category_patterns.get(category.lower())
            if pattern:
                # Check both schema name and data field for the pattern
                tools = [
                    t for t in tools
                    if pattern in (t.get("schemaname") or "") or
                       pattern in (t.get("data") or "")
                ]

        return tools

    def get_topic(self, component_id: str) -> dict:
        """
        Get a specific topic by component ID.

        Args:
            component_id: The topic component's unique identifier

        Returns:
            Topic component record
        """
        return self.get(f"botcomponents({component_id})")

    def get_tool(self, component_id: str) -> dict:
        """
        Get a specific tool by component ID.

        Tools are stored as botcomponents with componenttype 9 (Topic V2).
        This method fetches the component and validates it's a tool.

        Args:
            component_id: The tool component's unique identifier

        Returns:
            Tool component record with full details including data field

        Raises:
            Exception: If component is not found or is not a tool
        """
        component = self.get(f"botcomponents({component_id})")

        # Validate this is actually a tool
        schema_name = component.get("schemaname") or ""
        data = component.get("data") or ""

        is_tool = (
            "TaskAction" in schema_name or
            ".action." in schema_name or
            "kind: TaskDialog" in data
        )

        if not is_tool:
            raise Exception(f"Component {component_id} is not a tool (schema: {schema_name})")

        return component

    def set_topic_state(self, component_id: str, enabled: bool) -> None:
        """
        Enable or disable a topic.

        Args:
            component_id: The topic component's unique identifier
            enabled: True to enable (Active), False to disable (Inactive)

        Note:
            statecode values:
            - 0 = Active (enabled)
            - 1 = Inactive (disabled)
        """
        state_data = {
            "statecode": 0 if enabled else 1,
        }
        self.patch(f"botcomponents({component_id})", state_data)

    def create_topic(
        self,
        bot_id: str,
        name: str,
        content: str,
        description: Optional[str] = None,
        language: int = 1033,
    ) -> str:
        """
        Create a new topic for a bot.

        Args:
            bot_id: The bot's unique identifier
            name: Display name for the topic
            content: YAML content defining the topic's conversation flow
            description: Optional description for the topic
            language: Language code (default: 1033 for English)

        Returns:
            The created component ID

        Note:
            Topic content must be valid AdaptiveDialog YAML format.
            Use generate_simple_topic_yaml() to create basic topic content.

            Component types:
            - 0 = Topic (legacy)
            - 9 = Topic (V2) - used for new topics
        """
        # Get bot schema name for generating component schema name
        bot = self.get_bot(bot_id)
        bot_schema = bot.get("schemaname", f"cr83c_bot{bot_id[:8]}")

        # Generate schema name from display name
        clean_name = re.sub(r'[^a-zA-Z0-9]', '', name)
        schema_name = f"{bot_schema}.topic.{clean_name}"

        component_data = {
            "componenttype": 9,  # Topic (V2)
            "name": name,
            "schemaname": schema_name,
            "data": content,  # Topic YAML content is stored in 'data' field
            "language": language,
            "parentbotid@odata.bind": f"/bots({bot_id})"
        }

        if description:
            component_data["description"] = description

        # Use longer timeout for topic creation
        url = f"{self.api_url}/botcomponents"
        headers = self._get_headers()
        response = self._http_client.post(url, headers=headers, json=component_data, timeout=120.0)
        response.raise_for_status()

        # Extract component ID from OData-EntityId header
        entity_id = response.headers.get("OData-EntityId", "")
        if entity_id:
            match = re.search(r'botcomponents\(([^)]+)\)', entity_id)
            if match:
                return match.group(1)
        return ""

    def update_topic(
        self,
        component_id: str,
        name: Optional[str] = None,
        content: Optional[str] = None,
        description: Optional[str] = None,
    ) -> None:
        """
        Update an existing topic.

        Args:
            component_id: The topic component's unique identifier
            name: New display name for the topic
            content: New YAML content for the topic
            description: New description for the topic

        Raises:
            ClientError: If no updates provided or update fails
        """
        data = {}

        if name is not None:
            data["name"] = name
        if content is not None:
            data["data"] = content  # Topic YAML content is stored in 'data' field
        if description is not None:
            data["description"] = description

        if not data:
            raise ClientError("No updates provided. Specify at least one field to update.")

        self.patch(f"botcomponents({component_id})", data)

    def delete_topic(self, component_id: str) -> None:
        """
        Delete a topic from a bot.

        Args:
            component_id: The topic component's unique identifier

        Note:
            System topics (ismanaged=true) cannot be deleted.
            Only custom topics (ismanaged=false) can be deleted.
        """
        self.delete(f"botcomponents({component_id})")

    @staticmethod
    def generate_simple_topic_yaml(
        display_name: str,
        trigger_phrases: list[str],
        message: str,
    ) -> str:
        """
        Generate basic topic YAML from simple inputs.

        Args:
            display_name: Display name shown in topic list
            trigger_phrases: List of phrases that trigger this topic
            message: Response message to show when topic is triggered

        Returns:
            YAML string for a simple message-response topic

        Example:
            yaml = DataverseClient.generate_simple_topic_yaml(
                display_name="Greeting",
                trigger_phrases=["hello", "hi", "hey there"],
                message="Hello! How can I help you today?"
            )
        """
        import uuid

        # Generate unique IDs for nodes
        msg_id = f"sendMessage_{uuid.uuid4().hex[:8]}"

        # Build trigger phrases YAML
        triggers_yaml = "\n".join(f"      - {phrase}" for phrase in trigger_phrases)

        return f"""kind: AdaptiveDialog
beginDialog:
  kind: OnRecognizedIntent
  id: main
  intent:
    displayName: {display_name}
    triggerQueries:
{triggers_yaml}

  actions:
    - kind: SendMessage
      id: {msg_id}
      message: {message}
"""

    @staticmethod
    def generate_question_topic_yaml(
        display_name: str,
        trigger_phrases: list[str],
        question_prompt: str,
        variable_name: str,
        confirmation_message: str,
        entity_type: str = "StringPrebuiltEntity",
    ) -> str:
        """
        Generate topic YAML with a question node.

        Args:
            display_name: Display name shown in topic list
            trigger_phrases: List of phrases that trigger this topic
            question_prompt: The question to ask the user
            variable_name: Name for the variable to store the answer
            confirmation_message: Message shown after user answers
            entity_type: Entity type for validation (default: StringPrebuiltEntity)

        Returns:
            YAML string for a question-response topic

        Entity Types:
            - StringPrebuiltEntity: Free text input
            - BooleanPrebuiltEntity: Yes/No response
            - NumberPrebuiltEntity: Numeric input
            - DateTimePrebuiltEntity: Date/time input
            - EmailPrebuiltEntity: Email address
            - PersonNamePrebuiltEntity: Person's name
            - StatePrebuiltEntity: US state
            - CityPrebuiltEntity: City name
            - PhoneNumberPrebuiltEntity: Phone number
        """
        import uuid

        # Generate unique IDs for nodes
        question_id = f"question_{uuid.uuid4().hex[:8]}"
        msg_id = f"sendMessage_{uuid.uuid4().hex[:8]}"

        # Build trigger phrases YAML
        triggers_yaml = "\n".join(f"      - {phrase}" for phrase in trigger_phrases)

        return f"""kind: AdaptiveDialog
beginDialog:
  kind: OnRecognizedIntent
  id: main
  intent:
    displayName: {display_name}
    triggerQueries:
{triggers_yaml}

  actions:
    - kind: Question
      id: {question_id}
      alwaysPrompt: false
      variable: init:Topic.{variable_name}
      prompt: {question_prompt}
      entity: {entity_type}

    - kind: SendMessage
      id: {msg_id}
      message: {confirmation_message}
"""

    def delete_bot(self, bot_id: str) -> None:
        """
        Delete a bot by ID.

        Args:
            bot_id: The bot's unique identifier
        """
        self.delete(f"bots({bot_id})")

    def publish_bot(self, bot_id: str) -> dict:
        """
        Publish a Copilot Studio agent.

        This triggers the PvaPublish action which publishes the agent,
        making the latest changes available to users.

        Args:
            bot_id: The bot's unique identifier

        Returns:
            dict containing:
                - PublishedBotContentId: ID of the published content
                - status: "success" or error details

        Note:
            Publishing may take a few minutes to complete.
            The agent must be in a valid state to publish successfully.
        """
        url = f"{self.api_url}/bots({bot_id})/Microsoft.Dynamics.CRM.PvaPublish"
        headers = self._get_headers()

        try:
            response = self._http_client.post(url, headers=headers, json={}, timeout=120.0)
            response.raise_for_status()
            result = response.json()
            return {
                "status": "success",
                "PublishedBotContentId": result.get("PublishedBotContentId", ""),
            }
        except httpx.HTTPStatusError as e:
            error_detail = ""
            try:
                error_body = e.response.json()
                if "error" in error_body:
                    error_detail = error_body["error"].get("message", str(error_body))
            except Exception:
                error_detail = e.response.text[:500] if e.response.text else str(e)
            raise ClientError(f"Failed to publish agent: HTTP {e.response.status_code}: {error_detail}")

    def create_bot(
        self,
        name: str,
        schema_name: Optional[str] = None,
        language: int = 1033,
        instructions: Optional[str] = None,
        description: Optional[str] = None,
        orchestration: bool = True,
    ) -> dict:
        """
        Create a new Copilot Studio agent (bot).

        Args:
            name: Display name for the agent
            schema_name: Internal schema name (auto-generated from name if not provided)
            language: Language code (default: 1033 for English)
            instructions: Optional system instructions for the agent
            description: Optional description for the agent
            orchestration: Enable generative AI orchestration (default: True)

        Returns:
            Created bot record

        Note:
            Model selection and web search must be configured via the Copilot Studio portal UI.
            The API does not support these settings.
        """
        # Publisher prefix required by Dataverse for custom entities
        publisher_prefix = "cr83c_"

        # Generate schema name from display name if not provided
        if not schema_name:
            # Convert to camelCase and remove special characters
            clean_name = re.sub(r'[^a-zA-Z0-9\s]', '', name)
            words = clean_name.split()
            schema_name = words[0].lower() + ''.join(w.capitalize() for w in words[1:])

        # Add publisher prefix if not already present
        if not schema_name.startswith(publisher_prefix):
            schema_name = publisher_prefix + schema_name

        # Build configuration JSON
        config = {
            "$kind": "BotConfiguration",
            "settings": {
                "GenerativeActionsEnabled": orchestration
            },
            "isAgentConnectable": True,
            "gPTSettings": {
                "$kind": "GPTSettings",
                "defaultSchemaName": f"{schema_name}.gpt.default"
            },
            "aISettings": {
                "$kind": "AISettings",
                "useModelKnowledge": True,
                "isFileAnalysisEnabled": True,
                "isSemanticSearchEnabled": True,
                "contentModeration": "High",
                "optInUseLatestModels": False
            },
            "recognizer": {
                "$kind": "GenerativeAIRecognizer"
            }
        }

        # Add instructions if provided
        if instructions:
            config["gPTSettings"]["systemPrompt"] = instructions

        # Add description if provided
        if description:
            config["description"] = description

        bot_data = {
            "name": name,
            "schemaname": schema_name,
            "language": language,
            "runtimeprovider": 0,  # Power Virtual Agents
            "accesscontrolpolicy": 1,  # Copilot readers
            "authenticationmode": 2,  # Integrated
            "authenticationtrigger": 1,  # Always
            "template": "default-2.1.0",
            "configuration": json.dumps(config, indent=2),
        }

        return self.post("bots", bot_data)

    def update_bot(
        self,
        bot_id: str,
        name: Optional[str] = None,
        instructions: Optional[str] = None,
        description: Optional[str] = None,
        orchestration: Optional[bool] = None,
    ) -> None:
        """
        Update an existing Copilot Studio agent (bot).

        Args:
            bot_id: The bot's unique identifier
            name: New display name for the agent
            instructions: New system instructions for the agent
            description: New description for the agent
            orchestration: Enable/disable generative AI orchestration

        Note:
            Model selection and web search must be configured via the Copilot Studio portal UI.
            The API does not support these settings.
        """
        # Get current bot to preserve existing configuration
        current_bot = self.get_bot(bot_id)
        current_config = json.loads(current_bot.get("configuration", "{}"))

        bot_data = {}

        # Update name if provided
        if name is not None:
            bot_data["name"] = name

        # Update configuration fields if any are provided
        config_changed = False

        if orchestration is not None:
            if "settings" not in current_config:
                current_config["settings"] = {}
            current_config["settings"]["GenerativeActionsEnabled"] = orchestration
            config_changed = True

        if instructions is not None:
            if "gPTSettings" not in current_config:
                current_config["gPTSettings"] = {"$kind": "GPTSettings"}
            current_config["gPTSettings"]["systemPrompt"] = instructions
            config_changed = True

        if description is not None:
            current_config["description"] = description
            config_changed = True

        if config_changed:
            bot_data["configuration"] = json.dumps(current_config, indent=2)

        if not bot_data:
            raise ClientError("No updates provided. Specify at least one field to update.")

        self.patch(f"bots({bot_id})", bot_data)

    # =========================================================================
    # Authentication Methods
    # =========================================================================

    # Authentication mode constants
    AUTH_MODE_NONE = 1
    AUTH_MODE_INTEGRATED = 2
    AUTH_MODE_CUSTOM_AZURE_AD = 3

    AUTH_MODE_NAMES = {
        1: "None",
        2: "Integrated",
        3: "Custom Azure AD",
    }

    def get_bot_auth(self, bot_id: str) -> dict:
        """
        Get authentication configuration for a bot.

        Args:
            bot_id: The bot's unique identifier

        Returns:
            Dict containing authentication settings:
                - mode: Authentication mode integer (1=None, 2=Integrated, 3=Custom Azure AD)
                - mode_name: Human-readable authentication mode name
                - trigger: Authentication trigger (0=As Needed, 1=Always)
                - configuration: Authentication configuration JSON (if any)
        """
        bot = self.get_bot(bot_id)

        auth_mode = bot.get("authenticationmode", 2)
        auth_trigger = bot.get("authenticationtrigger", 1)
        auth_config = bot.get("authenticationconfiguration")

        return {
            "mode": auth_mode,
            "mode_name": self.AUTH_MODE_NAMES.get(auth_mode, f"Unknown({auth_mode})"),
            "trigger": auth_trigger,
            "trigger_name": "Always" if auth_trigger == 1 else "As Needed",
            "configuration": json.loads(auth_config) if auth_config else None,
        }

    def update_bot_auth(
        self,
        bot_id: str,
        mode: Optional[int] = None,
        trigger: Optional[int] = None,
        configuration: Optional[dict] = None,
    ) -> None:
        """
        Update authentication configuration for a bot.

        Args:
            bot_id: The bot's unique identifier
            mode: Authentication mode:
                - 1 = None (no authentication required)
                - 2 = Integrated (Microsoft Entra ID integrated)
                - 3 = Custom Azure AD (manual Microsoft Entra ID configuration)
            trigger: Authentication trigger:
                - 0 = As Needed (authenticate only when required)
                - 1 = Always (require authentication for all conversations)
            configuration: Authentication configuration dict (for Custom Azure AD mode)

        Note:
            When changing to Custom Azure AD (mode 3), you may also need to configure
            the authentication settings via the Copilot Studio portal, including:
            - Service provider settings
            - Client ID and tenant ID
            - Token exchange URL
        """
        bot_data = {}

        if mode is not None:
            if mode not in self.AUTH_MODE_NAMES:
                raise ClientError(f"Invalid authentication mode: {mode}. Valid modes: 1=None, 2=Integrated, 3=Custom Azure AD")
            bot_data["authenticationmode"] = mode

        if trigger is not None:
            if trigger not in (0, 1):
                raise ClientError(f"Invalid authentication trigger: {trigger}. Valid triggers: 0=As Needed, 1=Always")
            bot_data["authenticationtrigger"] = trigger

        if configuration is not None:
            bot_data["authenticationconfiguration"] = json.dumps(configuration)

        if not bot_data:
            raise ClientError("No updates provided. Specify at least one field to update.")

        self.patch(f"bots({bot_id})", bot_data)

    # =========================================================================
    # Application Insights Methods
    # =========================================================================

    def get_bot_app_insights(self, bot_id: str) -> dict:
        """
        Get Application Insights configuration for a bot.

        Args:
            bot_id: The bot's unique identifier

        Returns:
            Dict containing Application Insights settings:
                - enabled: Whether App Insights is configured
                - connectionString: The App Insights connection string (if set)
                - logActivities: Whether activity logging is enabled
                - logSensitiveProperties: Whether sensitive property logging is enabled
        """
        bot = self.get_bot(bot_id)
        config = json.loads(bot.get("configuration", "{}"))

        # Extract App Insights settings from configuration
        app_insights = config.get("applicationInsights", {})

        return {
            "enabled": bool(app_insights.get("connectionString")),
            "connectionString": app_insights.get("connectionString", ""),
            "logActivities": app_insights.get("logActivities", False),
            "logSensitiveProperties": app_insights.get("logSensitiveProperties", False),
        }

    def update_bot_app_insights(
        self,
        bot_id: str,
        connection_string: Optional[str] = None,
        log_activities: Optional[bool] = None,
        log_sensitive_properties: Optional[bool] = None,
        disable: bool = False,
    ) -> None:
        """
        Update Application Insights configuration for a bot.

        Args:
            bot_id: The bot's unique identifier
            connection_string: App Insights connection string (from Azure portal)
            log_activities: Enable logging of incoming/outgoing messages and events
            log_sensitive_properties: Enable logging of sensitive properties (userid, name, text, speak)
            disable: Set to True to disable Application Insights (clears connection string)

        Note:
            - Multiple agents can share the same App Insights instance by using the same connection string.
            - The connection string can be found in your Azure Application Insights resource overview.
            - After enabling, telemetry will be available in the Application Insights Logs section.
        """
        # Get current bot configuration
        current_bot = self.get_bot(bot_id)
        current_config = json.loads(current_bot.get("configuration", "{}"))

        # Initialize applicationInsights section if not present
        if "applicationInsights" not in current_config:
            current_config["applicationInsights"] = {}

        app_insights = current_config["applicationInsights"]

        # Handle disable
        if disable:
            app_insights["connectionString"] = ""
            app_insights["logActivities"] = False
            app_insights["logSensitiveProperties"] = False
        else:
            # Update connection string if provided
            if connection_string is not None:
                app_insights["connectionString"] = connection_string

            # Update logging options if provided
            if log_activities is not None:
                app_insights["logActivities"] = log_activities

            if log_sensitive_properties is not None:
                app_insights["logSensitiveProperties"] = log_sensitive_properties

        # Save updated configuration
        bot_data = {
            "configuration": json.dumps(current_config, indent=2)
        }

        self.patch(f"bots({bot_id})", bot_data)

    def get_app_insights_workspace_id(self, app_id: str) -> str:
        """
        Get Log Analytics workspace ID from Application Insights ApplicationId.

        This method queries Azure Resource Manager to find the App Insights resource
        and extract its linked Log Analytics workspace ID.

        Args:
            app_id: The ApplicationId from the App Insights connection string

        Returns:
            The Log Analytics workspace ID (GUID)

        Raises:
            ClientError: If the App Insights resource cannot be found
        """
        # Get ARM token
        arm_token = get_access_token_from_azure_cli("https://management.azure.com")

        # List all subscriptions to search for the App Insights resource
        headers = {
            "Authorization": f"Bearer {arm_token}",
            "Content-Type": "application/json",
        }

        # First, get subscriptions
        subs_url = "https://management.azure.com/subscriptions?api-version=2022-12-01"
        try:
            subs_response = self._http_client.get(subs_url, headers=headers, timeout=30.0)
            subs_response.raise_for_status()
            subscriptions = subs_response.json().get("value", [])
        except Exception as e:
            raise ClientError(f"Failed to list subscriptions: {e}")

        # Search each subscription for the App Insights resource
        for sub in subscriptions:
            sub_id = sub.get("subscriptionId")
            if not sub_id:
                continue

            # List App Insights components in this subscription
            components_url = (
                f"https://management.azure.com/subscriptions/{sub_id}"
                f"/providers/Microsoft.Insights/components"
                f"?api-version=2020-02-02"
            )

            try:
                response = self._http_client.get(components_url, headers=headers, timeout=30.0)
                if response.status_code == 200:
                    components = response.json().get("value", [])
                    for component in components:
                        # Check if this is the matching App Insights resource
                        if component.get("properties", {}).get("AppId") == app_id:
                            # Found it! Extract workspace ID
                            workspace_resource_id = component.get("properties", {}).get("WorkspaceResourceId")
                            if workspace_resource_id:
                                # Extract workspace ID from resource path
                                # Format: /subscriptions/.../workspaces/{workspace-name}
                                # We need to get the workspace GUID from the workspace resource
                                workspace_url = (
                                    f"https://management.azure.com{workspace_resource_id}"
                                    f"?api-version=2023-09-01"
                                )
                                ws_response = self._http_client.get(workspace_url, headers=headers, timeout=30.0)
                                if ws_response.status_code == 200:
                                    workspace = ws_response.json()
                                    # The customerId property is the workspace ID (GUID)
                                    workspace_id = workspace.get("properties", {}).get("customerId")
                                    if workspace_id:
                                        return workspace_id
                            raise ClientError(
                                f"App Insights resource found but no Log Analytics workspace linked. "
                                f"Please link the App Insights resource to a Log Analytics workspace."
                            )
            except httpx.HTTPStatusError:
                continue  # Try next subscription

        raise ClientError(
            f"Could not find Application Insights resource with ApplicationId: {app_id}. "
            f"Ensure you have access to the subscription containing this resource."
        )

    def query_app_insights(
        self,
        app_id: str,
        query: str,
        timespan: str = "P1D"
    ) -> dict:
        """
        Execute a KQL query against Application Insights.

        Args:
            app_id: The Application Insights App ID (from connection string)
            query: KQL query string
            timespan: ISO 8601 duration (e.g., "P1D" for 1 day, "PT24H" for 24 hours)

        Returns:
            Query results containing tables with columns and rows

        Raises:
            ClientError: If the query fails
        """
        # Get Application Insights API token
        ai_token = get_access_token_from_azure_cli("https://api.applicationinsights.io")

        url = f"https://api.applicationinsights.io/v1/apps/{app_id}/query"

        headers = {
            "Authorization": f"Bearer {ai_token}",
            "Content-Type": "application/json",
        }

        body = {
            "query": query,
            "timespan": timespan,
        }

        try:
            response = self._http_client.post(url, headers=headers, json=body, timeout=60.0)
            response.raise_for_status()
            result = response.json()

            # Check for query errors (semantic errors like missing tables)
            if "error" in result:
                error = result["error"]
                # Check if it's a "table not found" error
                inner = error.get("innererror", {})
                if "SEM0100" in str(inner) or "Failed to resolve table" in str(inner):
                    # Return empty result instead of error - no telemetry data yet
                    return {"tables": [{"name": "PrimaryResult", "columns": [], "rows": []}]}
                raise ClientError(f"Query error: {error.get('message', str(error))}")

            return result
        except httpx.HTTPStatusError as e:
            error_detail = ""
            try:
                error_body = e.response.json()
                if "error" in error_body:
                    error_msg = error_body["error"].get("message", "")
                    inner = error_body["error"].get("innererror", {})
                    # Check if it's a "table not found" error (400 Bad Request)
                    if "SEM0100" in str(inner) or "Failed to resolve table" in str(inner):
                        # Return empty result - no telemetry data yet
                        return {"tables": [{"name": "PrimaryResult", "columns": [], "rows": []}]}
                    error_detail = error_msg or str(error_body)
            except Exception:
                error_detail = e.response.text[:500] if e.response.text else str(e)
            raise ClientError(f"Application Insights query failed: HTTP {e.response.status_code}: {error_detail}")
        except httpx.RequestError as e:
            raise ClientError(f"Query request failed: {e}")

    def get_bot_telemetry(
        self,
        bot_id: str,
        timespan: str = "P1D",
        events_only: bool = False,
    ) -> dict:
        """
        Get telemetry data for a bot from Application Insights.

        This method retrieves the bot's App Insights configuration and queries
        the telemetry data directly using the Application Insights API.

        Args:
            bot_id: The bot's unique identifier
            timespan: ISO 8601 duration (e.g., "P1D" for 1 day, "PT1H" for 1 hour)
            events_only: If True, only query customEvents table

        Returns:
            Query results containing telemetry data

        Raises:
            ClientError: If App Insights is not configured or query fails
        """
        # Get bot's App Insights config
        config = self.get_bot_app_insights(bot_id)
        if not config["enabled"]:
            raise ClientError(
                "Application Insights is not configured for this agent. "
                "Use 'copilot agent analytics enable' to configure it first."
            )

        # Parse ApplicationId from connection string
        connection_string = config["connectionString"]
        app_id = parse_connection_string(connection_string).get("ApplicationId")
        if not app_id:
            raise ClientError(
                "Could not extract ApplicationId from connection string. "
                "Please check the App Insights configuration."
            )

        # Build KQL query
        if events_only:
            query = """
customEvents
| project timestamp, name, customDimensions, customMeasurements
| order by timestamp desc
"""
        else:
            query = """
customEvents
| extend _table = "customEvents"
| project timestamp, _table, name, message = "", customDimensions
| union (
    requests
    | extend _table = "requests"
    | project timestamp, _table, name, message = "", customDimensions
)
| union (
    traces
    | extend _table = "traces"
    | project timestamp, _table, name = "", message, customDimensions
)
| union (
    exceptions
    | extend _table = "exceptions"
    | project timestamp, _table, name = type, message = outerMessage, customDimensions
)
| order by timestamp desc
"""

        # Execute query using app_id directly (not workspace ID)
        return self.query_app_insights(app_id, query, timespan)

    def list_knowledge_sources(self, bot_id: str, source_type: Optional[str] = None) -> list[dict]:
        """
        List knowledge sources for a bot.

        Args:
            bot_id: The bot's unique identifier
            source_type: Filter by type - 'file' (14), 'connector' (16), or None for all

        Returns:
            List of knowledge source records
        """
        # Component types: 14 = Bot File Attachment, 16 = Knowledge Source (connectors)
        type_filter = ""
        if source_type == "file":
            type_filter = " and componenttype eq 14"
        elif source_type == "connector":
            type_filter = " and componenttype eq 16"
        else:
            # Get both file attachments and knowledge sources
            type_filter = " and (componenttype eq 14 or componenttype eq 16)"

        result = self.get(f"botcomponents?$filter=_parentbotid_value eq {bot_id}{type_filter}")
        return result.get("value", [])

    def add_file_knowledge_source(
        self,
        bot_id: str,
        name: str,
        content: str,
        description: Optional[str] = None,
    ) -> str:
        """
        Add a file-based knowledge source to a bot.

        Args:
            bot_id: The bot's unique identifier
            name: Display name for the knowledge source
            content: Text content for the knowledge source
            description: Optional description (auto-generated if not provided)

        Returns:
            The created component ID
        """
        # Get bot schema name for generating component schema name
        bot = self.get_bot(bot_id)
        bot_schema = bot.get("schemaname", f"cr83c_bot{bot_id[:8]}")

        # Generate schema name from display name
        clean_name = re.sub(r'[^a-zA-Z0-9]', '', name)
        schema_name = f"{bot_schema}.file.{clean_name}"

        # Auto-generate description if not provided
        if not description:
            description = f"This knowledge source searches information contained in {name}"

        component_data = {
            "componenttype": 14,  # Bot File Attachment
            "name": name,
            "schemaname": schema_name,
            "description": description,
            "content": content,
            "parentbotid@odata.bind": f"/bots({bot_id})"
        }

        # Use longer timeout for file operations
        url = f"{self.api_url}/botcomponents"
        headers = self._get_headers()
        response = self._http_client.post(url, headers=headers, json=component_data, timeout=120.0)
        response.raise_for_status()

        # Extract component ID from OData-EntityId header
        entity_id = response.headers.get("OData-EntityId", "")
        if entity_id:
            # Extract GUID from URL like .../botcomponents(guid)
            match = re.search(r'botcomponents\(([^)]+)\)', entity_id)
            if match:
                return match.group(1)
        return ""

    def add_azure_ai_search_knowledge_source(
        self,
        bot_id: str,
        name: str,
        search_endpoint: str,
        search_index: str,
        api_key: str,
        description: Optional[str] = None,
    ) -> str:
        """
        Add an Azure AI Search knowledge source to a bot.

        Args:
            bot_id: The bot's unique identifier
            name: Display name for the knowledge source
            search_endpoint: Azure AI Search endpoint URL
            search_index: Name of the search index
            api_key: Azure AI Search API key
            description: Optional description

        Returns:
            The created component ID

        Note:
            Azure AI Search knowledge sources require:
            1. A connection reference to the Azure AI Search connector
            2. Component type 16 (Knowledge Source)

            This method creates the necessary configuration.
        """
        # Get bot schema name for generating component schema name
        bot = self.get_bot(bot_id)
        bot_schema = bot.get("schemaname", f"cr83c_bot{bot_id[:8]}")

        # Generate schema name from display name
        clean_name = re.sub(r'[^a-zA-Z0-9]', '', name)
        schema_name = f"{bot_schema}.knowledge.{clean_name}"

        # Auto-generate description if not provided
        if not description:
            description = f"Azure AI Search knowledge source: {name}"

        # Build the knowledge source configuration
        # This follows the pattern used by Copilot Studio for Azure AI Search
        knowledge_config = {
            "$kind": "AzureAISearchKnowledgeSource",
            "endpoint": search_endpoint,
            "indexName": search_index,
            "apiKey": api_key,
        }

        component_data = {
            "componenttype": 16,  # Knowledge Source
            "name": name,
            "schemaname": schema_name,
            "description": description,
            "data": json.dumps(knowledge_config),
            "parentbotid@odata.bind": f"/bots({bot_id})"
        }

        # Use longer timeout
        url = f"{self.api_url}/botcomponents"
        headers = self._get_headers()
        response = self._http_client.post(url, headers=headers, json=component_data, timeout=120.0)
        response.raise_for_status()

        # Extract component ID from OData-EntityId header
        entity_id = response.headers.get("OData-EntityId", "")
        if entity_id:
            match = re.search(r'botcomponents\(([^)]+)\)', entity_id)
            if match:
                return match.group(1)
        return ""

    # Backwards compatibility alias
    def add_knowledge_source(
        self,
        bot_id: str,
        name: str,
        content: str,
        description: Optional[str] = None,
    ) -> str:
        """Alias for add_file_knowledge_source for backwards compatibility."""
        return self.add_file_knowledge_source(bot_id, name, content, description)

    def remove_knowledge_source(self, component_id: str) -> None:
        """
        Remove a knowledge source from a bot.

        Args:
            component_id: The knowledge source component's unique identifier
        """
        self.delete(f"botcomponents({component_id})")

    # =========================================================================
    # Tool Methods (Connected Agents, Flows, etc.)
    # =========================================================================

    def add_connected_agent_tool(
        self,
        bot_id: str,
        target_bot_id: str,
        name: Optional[str] = None,
        description: Optional[str] = None,
        pass_conversation_history: bool = True,
    ) -> str:
        """
        Add a connected agent tool to a bot.

        This creates an InvokeConnectedAgentTaskAction component that allows
        the bot to invoke another Copilot Studio agent as a sub-agent.

        Args:
            bot_id: The parent bot's unique identifier
            target_bot_id: The target agent's unique identifier to connect to
            name: Display name for the tool (defaults to target agent's name)
            description: Description of when to use this tool (for orchestration)
            pass_conversation_history: Whether to pass conversation history to the connected agent

        Returns:
            The created component ID

        Note:
            The target agent must:
            - Be in the same environment
            - Be published
            - Have "Let other agents connect" enabled in settings
        """
        # Get parent bot schema name
        bot = self.get_bot(bot_id)
        bot_schema = bot.get("schemaname", f"cr83c_bot{bot_id[:8]}")

        # Get target bot details
        target_bot = self.get_bot(target_bot_id)
        target_bot_name = target_bot.get("name", "Connected Agent")
        target_bot_schema = target_bot.get("schemaname", f"cr83c_bot{target_bot_id[:8]}")

        # Use target bot name if no name provided
        if not name:
            name = target_bot_name

        # Generate clean name for schema (remove spaces and special chars)
        clean_name = re.sub(r'[^a-zA-Z0-9]', '', name)
        schema_name = f"{bot_schema}.InvokeConnectedAgentTaskAction.{clean_name}"

        # Auto-generate description if not provided
        if not description:
            target_description = target_bot.get("description", "")
            if target_description:
                description = target_description
            else:
                description = f"Invoke the {target_bot_name} agent to handle specialized tasks."

        # Build the connected agent tool configuration in YAML format
        # This follows the Copilot Studio TaskDialog pattern
        tool_yaml = f"""kind: TaskDialog
modelDescription: {description}
schemaName: {schema_name}
action:
  kind: InvokeConnectedAgentTaskAction
  botSchemaName: {target_bot_schema}
  passConversationHistory: {str(pass_conversation_history).lower()}
inputType: {{}}
outputType: {{}}"""

        component_data = {
            "componenttype": 9,  # Topic (V2)
            "name": name,
            "schemaname": schema_name,
            "description": description,
            "data": tool_yaml,
            "parentbotid@odata.bind": f"/bots({bot_id})"
        }

        # Create the component
        url = f"{self.api_url}/botcomponents"
        headers = self._get_headers()
        response = self._http_client.post(url, headers=headers, json=component_data, timeout=120.0)
        response.raise_for_status()

        # Extract component ID from OData-EntityId header
        entity_id = response.headers.get("OData-EntityId", "")
        if entity_id:
            match = re.search(r'botcomponents\(([^)]+)\)', entity_id)
            if match:
                return match.group(1)
        return ""

    def remove_tool(self, component_id: str, cleanup_connection_ref: bool = True) -> None:
        """
        Remove a tool from a bot.

        For connector tools, this also removes the associated bot-specific
        connection reference if it's not used by any other tools.

        Args:
            component_id: The tool component's unique identifier
            cleanup_connection_ref: Whether to clean up unused connection references (default True)
        """
        import yaml as yaml_lib

        headers_req = self._get_headers()
        conn_ref_id = None
        conn_ref_logical_name = None

        # For connector tools, check if we need to clean up the connection reference
        if cleanup_connection_ref:
            try:
                tool = self.get(f"botcomponents({component_id})")
                schema_name = tool.get("schemaname", "") or ""
                data = tool.get("data", "") or ""

                # Check if this is a connector tool
                if ".action." in schema_name and data:
                    parsed_data = yaml_lib.safe_load(data) or {}
                    actions = parsed_data.get("actions") or []
                    if actions:
                        action = actions[0]
                        conn_ref_logical_name = action.get("connectionReferenceLogicalName", "")

                        if conn_ref_logical_name:
                            # Get the connection reference ID
                            check_url = f"{self.api_url}/connectionreferences?$filter=connectionreferencelogicalname eq '{conn_ref_logical_name}'&$select=connectionreferenceid"
                            check_response = self._http_client.get(check_url, headers=headers_req, timeout=60.0)
                            check_response.raise_for_status()
                            refs = check_response.json().get("value", [])
                            if refs:
                                conn_ref_id = refs[0].get("connectionreferenceid", "")
            except Exception:
                pass  # Don't fail tool removal if connection ref lookup fails

        # Delete the tool
        self.delete(f"botcomponents({component_id})")

        # Clean up connection reference if it exists and is no longer used
        if conn_ref_id and conn_ref_logical_name:
            try:
                # Check if any other botcomponents still reference this connection reference
                assoc_url = f"{self.api_url}/connectionreferences({conn_ref_id})/botcomponent_connectionreference?$select=botcomponentid&$top=1"
                assoc_response = self._http_client.get(assoc_url, headers=headers_req, timeout=60.0)
                assoc_response.raise_for_status()
                remaining = assoc_response.json().get("value", [])

                if not remaining:
                    # No other tools using this connection reference, safe to delete
                    self.delete(f"connectionreferences({conn_ref_id})")
            except Exception:
                pass  # Don't fail if cleanup fails - the tool was already removed

    def update_tool(
        self,
        component_id: str,
        name: Optional[str] = None,
        description: Optional[str] = None,
        availability: Optional[bool] = None,
        confirmation: Optional[bool] = None,
        confirmation_message: Optional[str] = None,
        inputs: Optional[dict] = None,
    ) -> dict:
        """
        Update a tool's attributes.

        Args:
            component_id: The tool component's unique identifier
            name: New display name for the tool
            description: New description for the tool (used by AI for orchestration)
            availability: Whether agent can use this tool dynamically (True = anytime, False = only from topics)
            confirmation: Whether to ask user for confirmation before running
            confirmation_message: Custom message to show when asking for confirmation
            inputs: Input parameter defaults as dict, e.g., {"workspace": "123", "project": "456"}

        Returns:
            The updated component data
        """
        # Get current component data
        component = self.get(f"botcomponents({component_id})")

        updates = {}
        data = component.get("data", "")

        if name is not None:
            updates["name"] = name

        if description is not None:
            updates["description"] = description
            # Also update modelDescription in the YAML data
            if data:
                # Match modelDescription: followed by the value (until next line with key or end)
                pattern = r'(modelDescription:)\s*[^\n]*'
                # Escape any special chars in description for YAML
                escaped_desc = description.replace('\\', '\\\\').replace('"', '\\"')
                replacement = f'\\1 {escaped_desc}'
                data = re.sub(pattern, replacement, data)

        # Handle availability setting (allowDynamicInvocation)
        if availability is not None and data:
            # Check if allowDynamicInvocation already exists
            if 'allowDynamicInvocation:' in data:
                # Update existing value
                pattern = r'(allowDynamicInvocation:)\s*(true|false)'
                replacement = f'\\1 {str(availability).lower()}'
                data = re.sub(pattern, replacement, data)
            else:
                # Add after kind: TaskDialog line
                pattern = r'(kind:\s*TaskDialog\n)'
                replacement = f'\\1allowDynamicInvocation: {str(availability).lower()}\n'
                data = re.sub(pattern, replacement, data)

        # Handle confirmation setting
        if confirmation is not None and data:
            if confirmation:
                # Add or update confirmation block
                message = confirmation_message or "Do you want to proceed with this action?"
                # Escape special characters in the message
                escaped_message = message.replace('\\', '\\\\').replace('"', '\\"').replace('\n', '\\n')

                confirmation_block = f'''confirmation:
  activity: "{escaped_message}"
  mode: Strict
'''
                # Check if confirmation block already exists
                if 'confirmation:' in data:
                    # Update existing confirmation block
                    # Match the entire confirmation block (multi-line)
                    pattern = r'confirmation:\s*\n\s*activity:[^\n]*\n\s*mode:[^\n]*\n?'
                    data = re.sub(pattern, confirmation_block, data)
                else:
                    # Add confirmation block after kind: TaskDialog (or allowDynamicInvocation if present)
                    if 'allowDynamicInvocation:' in data:
                        pattern = r'(allowDynamicInvocation:[^\n]*\n)'
                        replacement = f'\\1{confirmation_block}'
                    else:
                        pattern = r'(kind:\s*TaskDialog\n)'
                        replacement = f'\\1{confirmation_block}'
                    data = re.sub(pattern, replacement, data)
            else:
                # Remove confirmation block if it exists
                pattern = r'confirmation:\s*\n\s*activity:[^\n]*\n\s*mode:[^\n]*\n?'
                data = re.sub(pattern, '', data)
        elif confirmation_message is not None and data:
            # Only updating the confirmation message (confirmation is not explicitly set)
            if 'confirmation:' in data:
                escaped_message = confirmation_message.replace('\\', '\\\\').replace('"', '\\"').replace('\n', '\\n')
                pattern = r'(confirmation:\s*\n\s*activity:)[^\n]*'
                replacement = f'\\1 "{escaped_message}"'
                data = re.sub(pattern, replacement, data)

        # Handle input default values
        if inputs is not None and data:
            data = self._update_tool_inputs(data, inputs)

        # Check if YAML data was modified
        if data != component.get("data", ""):
            updates["data"] = data

        if not updates:
            return component

        # PATCH the component
        url = f"{self.api_url}/botcomponents({component_id})"
        headers = self._get_headers()
        response = self._http_client.patch(url, headers=headers, json=updates, timeout=60.0)
        response.raise_for_status()

        # Return updated component
        return self.get(f"botcomponents({component_id})")

    def add_tool(
        self,
        bot_id: str,
        tool_type: str,
        tool_id: str,
        name: Optional[str] = None,
        description: Optional[str] = None,
        inputs: Optional[dict] = None,
        outputs: Optional[dict] = None,
        # Type-specific parameters
        connection_ref: Optional[str] = None,
        no_history: bool = False,
        method: str = "GET",
        headers: Optional[dict] = None,
        body: Optional[str] = None,
        force: bool = False,
    ) -> str:
        """
        Add a tool to a bot.

        Args:
            bot_id: The parent bot's unique identifier
            tool_type: Tool type: 'connector', 'prompt', 'flow', 'http', 'agent'
            tool_id: Tool identifier (format depends on tool_type)
            name: Display name for the tool (auto-generated if not provided)
            description: Description for AI orchestration
            inputs: Input parameter schema (JSON dict)
            outputs: Output parameter schema (JSON dict)
            connection_ref: Connection reference ID (for connector/flow)
            no_history: Don't pass conversation history (for agent)
            method: HTTP method (for http)
            headers: HTTP headers (for http)
            body: Request body template (for http)
            force: Force adding tool even if operation has internal visibility

        Returns:
            The created component ID
        """
        # Get parent bot schema name
        bot = self.get_bot(bot_id)
        bot_schema = bot.get("schemaname", f"cr83c_bot{bot_id[:8]}")

        # Dispatch to type-specific generator
        generators = {
            'connector': self._generate_connector_tool_yaml,
            'prompt': self._generate_prompt_tool_yaml,
            'flow': self._generate_flow_tool_yaml,
            'http': self._generate_http_tool_yaml,
            'agent': self._generate_agent_tool_yaml,
        }

        generator = generators.get(tool_type.lower())
        if not generator:
            raise ClientError(f"Invalid tool type: {tool_type}. Must be one of: {', '.join(generators.keys())}")

        # Generate YAML and metadata
        tool_yaml, schema_name, resolved_name, resolved_description = generator(
            bot_id=bot_id,
            bot_schema=bot_schema,
            tool_id=tool_id,
            name=name,
            description=description,
            inputs=inputs,
            outputs=outputs,
            connection_ref=connection_ref,
            no_history=no_history,
            method=method,
            headers=headers,
            body=body,
            force=force,
        )

        component_data = {
            "componenttype": 9,  # Topic (V2)
            "name": resolved_name,
            "schemaname": schema_name,
            "description": resolved_description,
            "data": tool_yaml,
            "parentbotid@odata.bind": f"/bots({bot_id})"
        }

        # Create the component
        url = f"{self.api_url}/botcomponents"
        headers_req = self._get_headers()
        response = self._http_client.post(url, headers=headers_req, json=component_data, timeout=120.0)
        response.raise_for_status()

        # Extract component ID from OData-EntityId header
        component_id = ""
        entity_id = response.headers.get("OData-EntityId", "")
        if entity_id:
            match = re.search(r'botcomponents\(([^)]+)\)', entity_id)
            if match:
                component_id = match.group(1)

        # For connector tools, create and associate the bot-specific connection reference
        if tool_type.lower() == 'connector' and component_id and connection_ref:
            try:
                # Parse connector_id from tool_id (format: "connector_id:operation_id")
                connector_id = tool_id.split(':')[0] if ':' in tool_id else tool_id

                # Build the connection reference logical name
                conn_ref_logical_name = f"{bot_schema}.{connector_id}.{connection_ref}"

                # Check if connection reference already exists
                check_url = f"{self.api_url}/connectionreferences?$filter=connectionreferencelogicalname eq '{conn_ref_logical_name}'&$select=connectionreferenceid"
                check_response = self._http_client.get(check_url, headers=headers_req, timeout=60.0)
                check_response.raise_for_status()
                existing_refs = check_response.json().get("value", [])

                conn_ref_id = ""
                if existing_refs:
                    # Use existing connection reference
                    conn_ref_id = existing_refs[0].get("connectionreferenceid", "")
                else:
                    # Create a new bot-specific connection reference
                    conn_ref_data = {
                        "connectionreferencelogicalname": conn_ref_logical_name,
                        "connectionreferencedisplayname": conn_ref_logical_name,
                        "connectorid": f"/providers/Microsoft.PowerApps/apis/{connector_id}",
                        "connectionid": connection_ref,
                        "statecode": 0,
                        "statuscode": 1
                    }

                    conn_ref_url = f"{self.api_url}/connectionreferences"
                    conn_ref_response = self._http_client.post(
                        conn_ref_url, headers=headers_req, json=conn_ref_data, timeout=120.0
                    )
                    conn_ref_response.raise_for_status()

                    # Extract connection reference ID
                    conn_ref_entity_id = conn_ref_response.headers.get("OData-EntityId", "")
                    if conn_ref_entity_id:
                        match = re.search(r'connectionreferences\(([^)]+)\)', conn_ref_entity_id)
                        if match:
                            conn_ref_id = match.group(1)

                # Associate the connection reference with the botcomponent
                if conn_ref_id:
                    assoc_url = f"{self.api_url}/botcomponents({component_id})/botcomponent_connectionreference/$ref"
                    assoc_data = {
                        "@odata.id": f"{self.api_url}/connectionreferences({conn_ref_id})"
                    }
                    assoc_response = self._http_client.post(
                        assoc_url, headers=headers_req, json=assoc_data, timeout=120.0
                    )
                    assoc_response.raise_for_status()

            except Exception as e:
                # Log warning but don't fail - the tool was created successfully
                import sys
                print(f"Warning: Failed to create connection reference association: {e}", file=sys.stderr)

        return component_id

    def _update_tool_inputs(self, data: str, inputs: dict) -> str:
        """
        Update input default values in tool YAML data.

        Uses regex to preserve the original YAML structure while only modifying
        the specific input defaultValue fields.

        Args:
            data: The YAML data string from the tool component
            inputs: Dict of property names to default values, e.g., {"workspace": "123", "projects": "456"}

        Returns:
            Updated YAML data string
        """
        # Check if inputs section exists
        has_inputs_section = 'inputs:' in data

        for prop_name, default_value in inputs.items():
            # Check if this input already exists with defaultValue
            # Pattern matches: "- kind: ManualTaskInput\n    propertyName: X\n    defaultValue: Y"
            pattern_with_default = rf'(- kind: ManualTaskInput\s*\n\s*propertyName: {prop_name}\s*\n\s*)defaultValue:[^\n]*'
            if re.search(pattern_with_default, data):
                # Update existing defaultValue
                data = re.sub(pattern_with_default, rf'\1defaultValue: {default_value}', data)
            else:
                # Check if input exists without defaultValue
                pattern = rf'(- kind: ManualTaskInput\s*\n\s*propertyName: {prop_name})\s*\n'
                if re.search(pattern, data):
                    # Add defaultValue after propertyName
                    data = re.sub(pattern, rf'\1\n    defaultValue: {default_value}\n', data)
                elif not has_inputs_section:
                    # Need to add inputs section - add after kind: TaskDialog
                    new_input = f"inputs:\n  - kind: ManualTaskInput\n    propertyName: {prop_name}\n    defaultValue: {default_value}\n\n"
                    data = re.sub(r'(kind: TaskDialog\n)', rf'\1{new_input}', data)
                    has_inputs_section = True
                else:
                    # inputs section exists but this input doesn't - add to it
                    new_input = f"  - kind: ManualTaskInput\n    propertyName: {prop_name}\n    defaultValue: {default_value}\n\n"
                    data = re.sub(r'(inputs:\n)', rf'\1{new_input}', data)

        return data

    def _build_input_output_yaml(self, inputs: Optional[dict], outputs: Optional[dict]) -> tuple[str, str]:
        """Build inputType and outputType YAML sections."""
        if inputs:
            input_props = []
            for prop_name, prop_config in inputs.items():
                prop_type = prop_config.get("type", "String") if isinstance(prop_config, dict) else "String"
                prop_desc = prop_config.get("description", "") if isinstance(prop_config, dict) else ""
                if prop_desc:
                    input_props.append(f"    {prop_name}:\n      type: {prop_type}\n      description: {prop_desc}")
                else:
                    input_props.append(f"    {prop_name}:\n      type: {prop_type}")
            input_yaml = "inputType:\n  properties:\n" + "\n".join(input_props) if input_props else "inputType: {}"
        else:
            input_yaml = "inputType: {}"

        if outputs:
            output_props = []
            for prop_name, prop_config in outputs.items():
                prop_type = prop_config.get("type", "String") if isinstance(prop_config, dict) else "String"
                output_props.append(f"    {prop_name}:\n      type: {prop_type}")
            output_yaml = "outputType:\n  properties:\n" + "\n".join(output_props) if output_props else "outputType: {}"
        else:
            output_yaml = "outputType: {}"

        return input_yaml, output_yaml

    def _build_connector_outputs_yaml(self, connector_id: str, operation_id: str) -> str:
        """Build outputs YAML from connector's swagger response schema."""
        try:
            # Get connector details
            connector = self.get_connector(connector_id)
            swagger = connector.get('properties', {}).get('swagger', {})
            paths = swagger.get('paths', {})
            definitions = swagger.get('definitions', {})

            # Find the operation
            for path, methods in paths.items():
                for method, details in methods.items():
                    if details.get('operationId') == operation_id:
                        # Get response schema
                        responses = details.get('responses', {})
                        for code, resp in responses.items():
                            schema = resp.get('schema', {})
                            if '$ref' in schema:
                                # Resolve reference
                                ref_name = schema['$ref'].split('/')[-1]
                                schema = definitions.get(ref_name, {})

                            # Extract property names recursively
                            props = self._extract_property_names(schema, definitions, prefix='')
                            if props:
                                outputs_lines = ['outputs:']
                                for prop in props:
                                    outputs_lines.append(f'  - propertyName: {prop}')
                                    outputs_lines.append('')  # Empty line after each
                                return '\n'.join(outputs_lines) + '\n'
            return ''
        except Exception:
            return ''

    def _extract_property_names(self, schema: dict, definitions: dict, prefix: str = '', max_depth: int = 2) -> list:
        """Extract property names from schema, handling nested objects."""
        if max_depth <= 0:
            return []

        props = []
        properties = schema.get('properties', {})

        for name, prop_schema in properties.items():
            full_name = f'{prefix}{name}' if prefix else name

            # Check if it's a reference to another definition
            if '$ref' in prop_schema:
                ref_name = prop_schema['$ref'].split('/')[-1]
                nested_schema = definitions.get(ref_name, {})
                nested_props = self._extract_property_names(nested_schema, definitions, f'{full_name}.', max_depth - 1)
                props.extend(nested_props)
            elif prop_schema.get('type') == 'object' and 'properties' in prop_schema:
                nested_props = self._extract_property_names(prop_schema, definitions, f'{full_name}.', max_depth - 1)
                props.extend(nested_props)
            elif prop_schema.get('type') == 'array':
                # For arrays, just add the property name
                props.append(full_name)
            else:
                props.append(full_name)

        return props

    def _generate_agent_tool_yaml(
        self, bot_id: str, bot_schema: str, tool_id: str,
        name: Optional[str], description: Optional[str],
        inputs: Optional[dict], outputs: Optional[dict],
        no_history: bool = False, **kwargs
    ) -> tuple[str, str, str, str]:
        """Generate YAML for InvokeConnectedAgentTaskAction."""
        # Get target bot details
        target_bot = self.get_bot(tool_id)
        target_bot_name = target_bot.get("name", "Connected Agent")
        target_bot_schema = target_bot.get("schemaname", f"cr83c_bot{tool_id[:8]}")

        # Use target bot name if no name provided
        resolved_name = name or target_bot_name

        # Generate clean name for schema
        clean_name = re.sub(r'[^a-zA-Z0-9]', '', resolved_name)
        schema_name = f"{bot_schema}.InvokeConnectedAgentTaskAction.{clean_name}"

        # Auto-generate description if not provided
        if not description:
            target_description = target_bot.get("description", "")
            resolved_description = target_description or f"Invoke the {target_bot_name} agent to handle specialized tasks."
        else:
            resolved_description = description

        pass_history = str(not no_history).lower()
        input_yaml, output_yaml = self._build_input_output_yaml(inputs, outputs)

        tool_yaml = f"""kind: TaskDialog
modelDescription: {resolved_description}
schemaName: {schema_name}
action:
  kind: InvokeConnectedAgentTaskAction
  botSchemaName: {target_bot_schema}
  passConversationHistory: {pass_history}
{input_yaml}
{output_yaml}"""

        return tool_yaml, schema_name, resolved_name, resolved_description

    def _generate_connector_tool_yaml(
        self, bot_id: str, bot_schema: str, tool_id: str,
        name: Optional[str], description: Optional[str],
        inputs: Optional[dict], outputs: Optional[dict],
        connection_ref: Optional[str] = None,
        force: bool = False, **kwargs
    ) -> tuple[str, str, str, str]:
        """Generate YAML for InvokeConnectorTaskAction."""
        # Parse connector_id:operation_id format
        if ':' not in tool_id:
            raise ClientError("Connector tool --id must be in format 'connector_id:operation_id' (e.g., 'shared_asana:GetTask')")

        # Connection reference is required for connector tools
        if not connection_ref:
            raise ClientError(
                "Connector tools require --connection-ref parameter.\n"
                "Use 'copilot connection-references list --table' to find existing connection references."
            )

        connector_id, operation_id = tool_id.split(':', 1)

        # Validate operation exists and get its details from swagger
        operation_details = None
        operation_description = None
        try:
            connector = self.get_connector(connector_id)
            swagger = connector.get('properties', {}).get('swagger', {})
            paths = swagger.get('paths', {})

            # Find the operation in swagger
            for path, methods in paths.items():
                for method, details in methods.items():
                    if details.get('operationId') == operation_id:
                        operation_details = details
                        operation_description = details.get('description') or details.get('summary', '')
                        visibility = details.get('x-ms-visibility', '')
                        if visibility == 'internal' and not force:
                            raise ClientError(
                                f"Operation '{operation_id}' has internal visibility and cannot be used as a tool.\n"
                                f"Internal operations are not exposed in the Copilot Studio UI and may not work correctly.\n"
                                f"Use --force to add anyway, or use 'copilot connectors get {connector_id}' to see available operations."
                            )
                        break
                if operation_details:
                    break

            # If operation not found, reject with helpful error
            if not operation_details:
                # Get list of available operations for error message
                available_ops = []
                for path, methods in paths.items():
                    for method, details in methods.items():
                        op_id = details.get('operationId', '')
                        visibility = details.get('x-ms-visibility', '')
                        if op_id and visibility != 'internal':
                            available_ops.append(op_id)

                # Suggest similar operations
                similar = [op for op in available_ops if operation_id.lower().replace('_', '') in op.lower().replace('_', '')
                           or op.lower().replace('_', '') in operation_id.lower().replace('_', '')]

                error_msg = f"Operation '{operation_id}' not found in connector '{connector_id}'."
                if similar:
                    error_msg += f"\n\nDid you mean one of these?\n  " + "\n  ".join(similar[:5])
                error_msg += f"\n\nUse 'copilot connectors get {connector_id}' to see all available operations."
                raise ClientError(error_msg)

        except ClientError:
            raise
        except Exception as e:
            raise ClientError(f"Failed to validate operation: {e}")

        # Build full connection reference: {bot_schema}.{connector_id}.{connection_id}
        # connection_ref should be the connection GUID
        full_connection_ref = f"{bot_schema}.{connector_id}.{connection_ref}"

        # Get connector display name for naming
        connector_display_name = connector_id.replace('shared_', '').title()

        # Get operation display name from swagger
        operation_display_name = operation_details.get('summary', operation_id)

        # UI default name format: "{Connector} - {OperationDisplayName}"
        if name:
            resolved_name = name
        else:
            resolved_name = f"{connector_display_name} - {operation_display_name}"

        # Generate schema name matching UI pattern: {bot}.action.{Connector}-{OperationId}_{random}
        # Use random 3-char suffix like UI does for uniqueness
        random_suffix = ''.join(random.choices(string.ascii_uppercase + string.digits, k=3))
        schema_name = f"{bot_schema}.action.{connector_display_name}-{operation_id}_{random_suffix}"

        # Use swagger description if not provided by user
        resolved_description = description
        if not resolved_description:
            resolved_description = operation_description or f"Invoke {operation_id} operation from {connector_id} connector."

        # Build outputs from connector response schema
        outputs_yaml = self._build_connector_outputs_yaml(connector_id, operation_id)

        tool_yaml = f"""kind: TaskDialog
modelDisplayName: {resolved_name}
modelDescription: {resolved_description}
{outputs_yaml}action:
  kind: InvokeConnectorTaskAction
  connectionReference: {full_connection_ref}
  connectionProperties:
    mode: Invoker

  operationId: {operation_id}

outputMode: All"""

        return tool_yaml, schema_name, resolved_name, resolved_description

    def _generate_prompt_tool_yaml(
        self, bot_id: str, bot_schema: str, tool_id: str,
        name: Optional[str], description: Optional[str],
        inputs: Optional[dict], outputs: Optional[dict], **kwargs
    ) -> tuple[str, str, str, str]:
        """Generate YAML for InvokePromptTaskAction."""
        # Try to get prompt details
        try:
            prompt = self.get_prompt(tool_id)
            prompt_name = prompt.get("msdyn_name", "AI Prompt")
            prompt_schema = prompt.get("schemaname", "")
        except Exception:
            prompt_name = "AI Prompt"
            prompt_schema = ""

        # Use prompt name if no name provided
        resolved_name = name or prompt_name

        # Generate clean name for schema
        clean_name = re.sub(r'[^a-zA-Z0-9]', '', resolved_name)
        schema_name = f"{bot_schema}.InvokePromptTaskAction.{clean_name}"

        # Auto-generate description if not provided
        resolved_description = description or f"Invoke the {prompt_name} AI Builder prompt."

        input_yaml, output_yaml = self._build_input_output_yaml(inputs, outputs)

        # Build action section
        action_lines = [
            "action:",
            "  kind: InvokePromptTaskAction",
            f"  promptId: {tool_id}",
        ]
        if prompt_schema:
            action_lines.append(f"  promptSchemaName: {prompt_schema}")

        tool_yaml = f"""kind: TaskDialog
modelDescription: {resolved_description}
schemaName: {schema_name}
{chr(10).join(action_lines)}
{input_yaml}
{output_yaml}"""

        return tool_yaml, schema_name, resolved_name, resolved_description

    def _generate_flow_tool_yaml(
        self, bot_id: str, bot_schema: str, tool_id: str,
        name: Optional[str], description: Optional[str],
        inputs: Optional[dict], outputs: Optional[dict],
        connection_ref: Optional[str] = None, **kwargs
    ) -> tuple[str, str, str, str]:
        """Generate YAML for InvokeFlowTaskAction."""
        # Use flow ID as name if not provided
        resolved_name = name or f"Flow {tool_id[:8]}"

        # Generate clean name for schema
        clean_name = re.sub(r'[^a-zA-Z0-9]', '', resolved_name)
        schema_name = f"{bot_schema}.InvokeFlowTaskAction.{clean_name}"

        # Auto-generate description if not provided
        resolved_description = description or f"Invoke Power Automate flow."

        input_yaml, output_yaml = self._build_input_output_yaml(inputs, outputs)

        # Build action section with proper flow ID format
        flow_ref = tool_id if tool_id.startswith('/providers/') else f"/providers/Microsoft.Flow/flows/{tool_id}"
        action_lines = [
            "action:",
            "  kind: InvokeFlowTaskAction",
            f"  flowId: {flow_ref}",
        ]
        if connection_ref:
            action_lines.append(f"  connectionReference: {connection_ref}")

        tool_yaml = f"""kind: TaskDialog
modelDescription: {resolved_description}
schemaName: {schema_name}
{chr(10).join(action_lines)}
{input_yaml}
{output_yaml}"""

        return tool_yaml, schema_name, resolved_name, resolved_description

    def _generate_http_tool_yaml(
        self, bot_id: str, bot_schema: str, tool_id: str,
        name: Optional[str], description: Optional[str],
        inputs: Optional[dict], outputs: Optional[dict],
        method: str = "GET", headers: Optional[dict] = None,
        body: Optional[str] = None, **kwargs
    ) -> tuple[str, str, str, str]:
        """Generate YAML for InvokeHttpTaskAction."""
        # tool_id is the URL for HTTP tools
        url = tool_id

        # Use method + shortened URL as name if not provided
        resolved_name = name or f"{method} API"

        # Generate clean name for schema
        clean_name = re.sub(r'[^a-zA-Z0-9]', '', resolved_name)
        schema_name = f"{bot_schema}.InvokeHttpTaskAction.{clean_name}"

        # Auto-generate description if not provided
        resolved_description = description or f"Make {method} request to {url}"

        input_yaml, output_yaml = self._build_input_output_yaml(inputs, outputs)

        # Build action section
        action_lines = [
            "action:",
            "  kind: InvokeHttpTaskAction",
            f"  method: {method.upper()}",
            f"  url: {url}",
        ]

        if headers:
            action_lines.append("  headers:")
            for header_name, header_value in headers.items():
                action_lines.append(f"    {header_name}: {header_value}")

        if body:
            # Use YAML literal block for body
            body_indented = body.replace('\n', '\n    ')
            action_lines.append(f"  body: |\n    {body_indented}")

        tool_yaml = f"""kind: TaskDialog
modelDescription: {resolved_description}
schemaName: {schema_name}
{chr(10).join(action_lines)}
{input_yaml}
{output_yaml}"""

        return tool_yaml, schema_name, resolved_name, resolved_description

    # =========================================================================
    # Environment Methods
    # =========================================================================

    def list_environments(self) -> list[dict]:
        """
        List all Power Platform environments accessible to the user.

        Returns:
            List of environment records from Business App Platform API

        Note:
            This uses the BAP (Business App Platform) API to list environments.
        """
        bap_token = get_access_token_from_azure_cli("https://api.bap.microsoft.com/")

        url = (
            "https://api.bap.microsoft.com/providers/Microsoft.BusinessAppPlatform/environments"
            "?api-version=2021-04-01"
        )

        headers = {
            "Authorization": f"Bearer {bap_token}",
            "Accept": "application/json",
        }

        try:
            response = self._http_client.get(url, headers=headers, timeout=60.0)
            response.raise_for_status()
            data = response.json()
            return data.get("value", [])
        except httpx.HTTPStatusError as e:
            error_detail = ""
            try:
                error_body = e.response.json()
                error_detail = f": {error_body}"
            except Exception:
                error_detail = f": {e.response.text[:200]}" if e.response.text else ""
            raise ClientError(f"Failed to list environments (HTTP {e.response.status_code}){error_detail}")

    def get_environment(self, environment_id: str) -> dict:
        """
        Get details for a specific Power Platform environment.

        Args:
            environment_id: The environment ID (e.g., Default-<tenant-id> or GUID)

        Returns:
            Environment record from Business App Platform API
        """
        bap_token = get_access_token_from_azure_cli("https://api.bap.microsoft.com/")

        url = (
            f"https://api.bap.microsoft.com/providers/Microsoft.BusinessAppPlatform/environments/{environment_id}"
            "?api-version=2021-04-01"
        )

        headers = {
            "Authorization": f"Bearer {bap_token}",
            "Accept": "application/json",
        }

        try:
            response = self._http_client.get(url, headers=headers, timeout=60.0)
            response.raise_for_status()
            return response.json()
        except httpx.HTTPStatusError as e:
            error_detail = ""
            try:
                error_body = e.response.json()
                error_detail = f": {error_body}"
            except Exception:
                error_detail = f": {e.response.text[:200]}" if e.response.text else ""
            raise ClientError(f"Failed to get environment (HTTP {e.response.status_code}){error_detail}")

    # =========================================================================
    # Connector Methods
    # =========================================================================

    def list_connectors(
        self,
        environment_id: Optional[str] = None,
        include_actions: bool = False,
    ) -> list[dict]:
        """
        List all available connectors (both custom and managed) in the environment.

        Args:
            environment_id: Power Platform environment ID. If not provided,
                            will use DATAVERSE_ENVIRONMENT_ID from config.
            include_actions: If True, fetch full connector details including
                            swagger/actions for each connector. This is slower
                            but provides operation details.

        Returns:
            List of connector records from Power Apps API

        Note:
            This uses the Power Apps API to list all connectors available
            in the environment, including both managed (Microsoft) and
            custom connectors.
        """
        # Get environment ID from config if not provided
        if not environment_id:
            config = get_config()
            environment_id = config.environment_id
            if not environment_id:
                raise ClientError(
                    "Environment ID not found. Please set DATAVERSE_ENVIRONMENT_ID "
                    "in your .env file (e.g., Default-<tenant-id> or the environment GUID).\n\n"
                    "You can find your environment ID in the Power Platform admin center."
                )

        powerapps_token = get_access_token_from_azure_cli("https://service.powerapps.com/")

        url = (
            f"https://api.powerapps.com/providers/Microsoft.PowerApps/apis"
            f"?api-version=2016-11-01"
            f"&$filter=environment eq '{environment_id}'"
        )

        headers = {
            "Authorization": f"Bearer {powerapps_token}",
            "Accept": "application/json",
        }

        try:
            response = self._http_client.get(url, headers=headers, timeout=60.0)
            response.raise_for_status()
            data = response.json()
            connectors = data.get("value", [])

            # If include_actions, fetch full details for each connector
            if include_actions:
                detailed_connectors = []
                for conn in connectors:
                    connector_id = conn.get("name", "")
                    if connector_id:
                        try:
                            detailed = self.get_connector(connector_id, environment_id)
                            detailed_connectors.append(detailed)
                        except Exception:
                            # If we can't get details, use the basic info
                            detailed_connectors.append(conn)
                return detailed_connectors

            return connectors
        except httpx.HTTPStatusError as e:
            error_detail = ""
            try:
                error_body = e.response.json()
                if "error" in error_body:
                    error_detail = error_body["error"].get("message", str(error_body))
            except Exception:
                error_detail = e.response.text[:500] if e.response.text else str(e)
            raise ClientError(f"Failed to list connectors: HTTP {e.response.status_code}: {error_detail}")
        except httpx.RequestError as e:
            raise ClientError(f"Request failed: {e}")

    def get_connector(self, connector_id: str, environment_id: Optional[str] = None) -> dict:
        """
        Get a specific connector by ID.

        Args:
            connector_id: The connector's unique identifier (e.g., shared_office365)
            environment_id: Power Platform environment ID. If not provided,
                            will use DATAVERSE_ENVIRONMENT_ID from config.

        Returns:
            Connector record from Power Apps API
        """
        # Get environment ID from config if not provided
        if not environment_id:
            config = get_config()
            environment_id = config.environment_id
            if not environment_id:
                raise ClientError(
                    "Environment ID not found. Please set DATAVERSE_ENVIRONMENT_ID "
                    "in your .env file (e.g., Default-<tenant-id> or the environment GUID).\n\n"
                    "You can find your environment ID in the Power Platform admin center."
                )

        powerapps_token = get_access_token_from_azure_cli("https://service.powerapps.com/")

        url = (
            f"https://api.powerapps.com/providers/Microsoft.PowerApps/apis/{connector_id}"
            f"?api-version=2016-11-01"
            f"&$filter=environment eq '{environment_id}'"
        )

        headers = {
            "Authorization": f"Bearer {powerapps_token}",
            "Accept": "application/json",
        }

        try:
            response = self._http_client.get(url, headers=headers, timeout=30.0)
            response.raise_for_status()
            return response.json()
        except httpx.HTTPStatusError as e:
            error_detail = ""
            try:
                error_body = e.response.json()
                if "error" in error_body:
                    error_detail = error_body["error"].get("message", str(error_body))
            except Exception:
                error_detail = e.response.text[:500] if e.response.text else str(e)
            raise ClientError(f"Failed to get connector: HTTP {e.response.status_code}: {error_detail}")
        except httpx.RequestError as e:
            raise ClientError(f"Request failed: {e}")

    # =========================================================================
    # Flow Methods
    # =========================================================================

    def list_flows(self, category: int = None) -> list[dict]:
        """
        List Power Automate cloud flows in the environment.

        Args:
            category: Optional filter by flow category:
                - 0: Standard (automated/scheduled flows)
                - 5: Instant (button/HTTP triggered flows)
                - 6: Business process flows

        Returns:
            List of workflow (flow) records

        Note:
            Flows that can be used as agent tools are typically category 5 (instant)
            or flows with HTTP triggers.
        """
        endpoint = "workflows?$filter=type eq 1"  # type 1 = Flow (vs 2 = Action)
        if category is not None:
            endpoint += f" and category eq {category}"
        endpoint += "&$orderby=name"
        result = self.get(endpoint)
        return result.get("value", [])

    def get_flow(self, workflow_id: str) -> dict:
        """
        Get a specific flow by ID.

        Args:
            workflow_id: The workflow's unique identifier (GUID)

        Returns:
            Workflow (flow) record
        """
        return self.get(f"workflows({workflow_id})")

    # =========================================================================
    # Prompt Methods (AI Builder Prompts)
    # =========================================================================

    # GptPowerPrompt template ID - identifies AI Builder prompts
    GPT_POWER_PROMPT_TEMPLATE_ID = "edfdb190-3791-45d8-9a6c-8f90a37c278a"

    def list_prompts(self) -> list[dict]:
        """
        List AI Builder prompts available as agent tools.

        AI Builder prompts are custom prompts that can be attached to
        Copilot Studio agents as tools. They use GPT models to perform
        specific tasks like classification, extraction, or content generation.

        Returns:
            List of prompt (msdyn_aimodel) records with GptPowerPrompt template

        Note:
            Prompts are stored as msdyn_aimodels with the GptPowerPrompt template.
            This filters out other AI model types like Invoice Processing, etc.
        """
        result = self.get(
            f"msdyn_aimodels?$filter=_msdyn_templateid_value eq {self.GPT_POWER_PROMPT_TEMPLATE_ID}"
            f"&$orderby=msdyn_name"
        )
        return result.get("value", [])

    def get_prompt(self, prompt_id: str) -> dict:
        """
        Get a specific AI Builder prompt by ID.

        Args:
            prompt_id: The prompt's unique identifier (GUID)

        Returns:
            Prompt (msdyn_aimodel) record
        """
        return self.get(f"msdyn_aimodels({prompt_id})")

    def delete_prompt(self, prompt_id: str) -> None:
        """
        Delete an AI Builder prompt by ID.

        Args:
            prompt_id: The prompt's unique identifier (GUID)

        Note:
            Managed (system) prompts cannot be deleted.
            Only custom prompts can be deleted.
        """
        self.delete(f"msdyn_aimodels({prompt_id})")

    def get_prompt_configuration(self, prompt_id: str, active_only: bool = True) -> dict:
        """
        Get the AI Configuration for a prompt, including the prompt text.

        Args:
            prompt_id: The prompt's unique identifier (GUID)
            active_only: If True (default), return the published/active configuration.
                        If False, return the most recently modified configuration.

        Returns:
            Dict containing:
                - configuration_id: The AI configuration ID
                - prompt_text: The full prompt text (all literal parts concatenated)
                - prompt_parts: The raw prompt parts array
                - custom_configuration: The full custom configuration JSON
                - model_type: The GPT model type (e.g., gpt-41-mini)
                - status: Status code (0=Draft, 7=Published, etc.)
                - version: Version string (major.minor)

        Raises:
            ClientError: If no configuration found for the prompt
        """
        import json

        # Get AI configurations for this model
        # Type 190690001 = RunConfiguration (the ones with prompt text)
        result = self.get(
            f"msdyn_aiconfigurations?"
            f"$filter=_msdyn_aimodelid_value eq {prompt_id} and msdyn_type eq 190690001"
            f"&$orderby=modifiedon desc"
        )
        configs = result.get("value", [])

        if not configs:
            raise ClientError(f"No AI configuration found for prompt {prompt_id}")

        # Find the appropriate config
        config = None
        if active_only:
            # Look for published config (statuscode=7)
            for c in configs:
                if c.get("statuscode") == 7:
                    config = c
                    break
            # Fall back to most recent if no published config
            if not config:
                config = configs[0]
        else:
            config = configs[0]

        config_id = config.get("msdyn_aiconfigurationid")
        custom_config_str = config.get("msdyn_customconfiguration", "")
        status = config.get("statuscode", 0)
        major = config.get("msdyn_majoriterationnumber", 1)
        minor = config.get("msdyn_minoriterationnumber", 0)

        # Parse the custom configuration JSON
        prompt_text = ""
        prompt_parts = []
        model_type = ""
        custom_config = {}

        if custom_config_str:
            try:
                custom_config = json.loads(custom_config_str)
                prompt_parts = custom_config.get("prompt", [])

                # Extract just the literal text parts to form the prompt text
                text_parts = []
                for part in prompt_parts:
                    if part.get("type") == "literal":
                        text_parts.append(part.get("text", ""))

                prompt_text = "".join(text_parts)

                # Get model type
                model_params = custom_config.get("modelParameters", {})
                model_type = model_params.get("modelType", "")

            except json.JSONDecodeError:
                pass

        return {
            "configuration_id": config_id,
            "prompt_text": prompt_text,
            "prompt_parts": prompt_parts,
            "custom_configuration": custom_config,
            "model_type": model_type,
            "status": status,
            "version": f"{major}.{minor}",
        }

    def update_prompt(
        self,
        prompt_id: str,
        prompt_text: Optional[str] = None,
        model_type: Optional[str] = None,
        publish: bool = True,
    ) -> None:
        """
        Update an AI Builder prompt's text or model type.

        This method handles the full publish workflow:
        1. Finds the active (published) configuration
        2. Unpublishes it if currently published
        3. Updates the configuration
        4. Republishes it (if publish=True)

        Args:
            prompt_id: The prompt's unique identifier (GUID)
            prompt_text: New prompt text (replaces all literal parts with single text)
            model_type: New model type (e.g., gpt-41-mini, gpt-4o, gpt-4o-mini)
            publish: If True (default), republish after updating

        Raises:
            ClientError: If update fails or no configuration found
        """
        import json
        import time

        if not prompt_text and not model_type:
            raise ClientError("Must provide prompt_text or model_type to update")

        # Get current active configuration
        config_info = self.get_prompt_configuration(prompt_id, active_only=True)
        config_id = config_info["configuration_id"]
        custom_config = config_info["custom_configuration"]
        status = config_info["status"]
        version = config_info["version"]

        if not custom_config:
            raise ClientError("No custom configuration found for prompt")

        # If published, unpublish first
        if status == 7:  # Published
            self.post(
                f"msdyn_aiconfigurations({config_id})/Microsoft.Dynamics.CRM.UnpublishAIConfiguration",
                {"version": version}
            )
            # Wait for unpublish to complete
            for _ in range(15):
                config = self.get(f"msdyn_aiconfigurations({config_id})")
                if config.get("statuscode") != 4:  # 4 = Unpublishing
                    break
                time.sleep(1)

        # Update prompt text if provided
        if prompt_text is not None:
            # Get existing input variables from prompt parts
            input_vars = [
                part for part in custom_config.get("prompt", [])
                if part.get("type") == "inputVariable"
            ]

            # Create new prompt parts with the new text
            new_prompt_parts = [{"type": "literal", "text": prompt_text}]

            # Append any input variables that were in the original
            for var in input_vars:
                new_prompt_parts.append(var)

            custom_config["prompt"] = new_prompt_parts

        # Update model type if provided
        if model_type is not None:
            if "modelParameters" not in custom_config:
                custom_config["modelParameters"] = {}
            custom_config["modelParameters"]["modelType"] = model_type

        # Update the configuration
        update_data = {
            "msdyn_customconfiguration": json.dumps(custom_config)
        }
        self.patch(f"msdyn_aiconfigurations({config_id})", update_data)

        # Republish if requested
        if publish:
            self.post(
                f"msdyn_aiconfigurations({config_id})/Microsoft.Dynamics.CRM.PublishAIConfiguration",
                {"version": version}
            )
            # Wait for publish to complete
            for _ in range(30):
                config = self.get(f"msdyn_aiconfigurations({config_id})")
                current_status = config.get("statuscode")
                if current_status == 7:  # Published
                    break
                if current_status in [10, 11, 12, 13]:  # Failed states
                    raise ClientError(f"Publish failed with status {current_status}")
                time.sleep(1)

    # =========================================================================
    # REST API Methods (Custom Connectors)
    # =========================================================================

    def list_rest_apis(self) -> list[dict]:
        """
        List REST API tools (custom connectors) available for agents.

        REST API tools are custom connectors defined with OpenAPI specifications
        that can be attached to Copilot Studio agents as tools.

        Returns:
            List of connector records with connectortype=1 (CustomConnector)

        Note:
            These are custom connectors stored in Dataverse, not the Power Apps
            connector catalog (which is queried by list_connectors()).
        """
        result = self.get(
            "connectors?$filter=connectortype eq 1"
            "&$orderby=displayname"
        )
        return result.get("value", [])

    def get_rest_api(self, connector_id: str) -> dict:
        """
        Get a specific REST API tool (custom connector) by ID.

        Args:
            connector_id: The connector's unique identifier (GUID)

        Returns:
            Connector record
        """
        return self.get(f"connectors({connector_id})")

    def delete_rest_api(self, connector_id: str) -> None:
        """
        Delete a REST API tool (custom connector) by ID.

        Args:
            connector_id: The connector's unique identifier (GUID)
        """
        self.delete(f"connectors({connector_id})")

    # =========================================================================
    # MCP Server Methods (Model Context Protocol)
    # =========================================================================

    def list_mcp_servers(self, environment_id: Optional[str] = None) -> list[dict]:
        """
        List MCP (Model Context Protocol) servers available as agent tools.

        MCP servers are connectors that implement the Model Context Protocol,
        allowing agents to connect to external data sources and tools.

        Args:
            environment_id: Power Platform environment ID. If not provided,
                            will use DATAVERSE_ENVIRONMENT_ID from config.

        Returns:
            List of MCP server connector records from Power Apps API

        Note:
            MCP servers are identified by having 'mcp' in their connector ID,
            name, or description.
        """
        # Get all connectors from Power Apps API
        connectors = self.list_connectors(environment_id)

        # Filter for MCP servers
        mcp_servers = []
        for connector in connectors:
            name = connector.get("name", "")
            props = connector.get("properties", {})
            display_name = props.get("displayName", "")
            description = (props.get("description", "") or "").lower()

            # Check for MCP indicators
            is_mcp = False
            if "mcpserver" in name.lower():
                is_mcp = True
            elif "mcp" in name.lower() and name.startswith("shared_"):
                is_mcp = True
            elif "mcp" in display_name.lower():
                is_mcp = True
            elif "model context protocol" in description:
                is_mcp = True

            if is_mcp:
                mcp_servers.append(connector)

        return mcp_servers

    def get_mcp_server(self, connector_id: str) -> dict:
        """
        Get a specific MCP server connector by ID.

        Args:
            connector_id: The connector's unique identifier (e.g., shared_microsoftlearndocsmcpserver)

        Returns:
            MCP server connector record from Power Apps API
        """
        return self.get_connector(connector_id)

    # =========================================================================
    # Transcript Methods
    # =========================================================================

    def list_transcripts(
        self,
        bot_id: Optional[str] = None,
        bot_name: Optional[str] = None,
        limit: int = 20,
        select: Optional[list[str]] = None,
    ) -> list[dict]:
        """
        List conversation transcripts, optionally filtered by bot ID or name.

        Args:
            bot_id: Optional bot ID to filter transcripts
            bot_name: Optional bot name to filter transcripts (resolved to ID)
            limit: Maximum number of transcripts to return (default: 20)
            select: Optional list of fields to select

        Returns:
            List of transcript records

        Raises:
            ClientError: If bot_name is provided but no matching bot is found
        """
        # Resolve bot name to ID if provided
        filter_bot_id = bot_id
        if bot_name and not bot_id:
            bot = self.get_bot_by_name(bot_name)
            if not bot:
                raise ClientError(f"No bot found with name: {bot_name}")
            filter_bot_id = bot.get("botid")

        endpoint = "conversationtranscripts"
        params = []

        # Default fields if not specified
        if not select:
            select = [
                "conversationtranscriptid",
                "name",
                "conversationstarttime",
                "_bot_conversationtranscriptid_value",
                "schematype",
            ]
        params.append(f"$select={','.join(select)}")

        # Filter by bot if specified
        if filter_bot_id:
            params.append(f"$filter=_bot_conversationtranscriptid_value eq {filter_bot_id}")

        # Order by most recent first
        params.append("$orderby=conversationstarttime desc")
        params.append(f"$top={limit}")

        if params:
            endpoint += "?" + "&".join(params)

        result = self.get(endpoint)
        return result.get("value", [])

    def get_transcript(self, transcript_id: str) -> dict:
        """
        Get a single transcript with full content.

        Args:
            transcript_id: The transcript's unique identifier

        Returns:
            Transcript record including full content
        """
        return self.get(f"conversationtranscripts({transcript_id})")

    # =========================================================================
    # Solution Management Methods
    # =========================================================================

    def list_solutions(self, select: Optional[list[str]] = None) -> list[dict]:
        """
        List all solutions in the environment.

        Args:
            select: Optional list of fields to select

        Returns:
            List of solution records
        """
        endpoint = "solutions"
        params = {}
        if select:
            params["$select"] = ",".join(select)
        # Filter out system solutions, only show unmanaged solutions by default
        params["$filter"] = "ismanaged eq false"
        params["$orderby"] = "friendlyname"
        result = self.get(endpoint, params=params if params else None)
        return result.get("value", [])

    def get_solution(self, solution_id: str) -> dict:
        """
        Get a specific solution by ID.

        Args:
            solution_id: The solution's unique identifier (GUID) or unique name

        Returns:
            Solution record
        """
        # Try to get by GUID first
        try:
            return self.get(f"solutions({solution_id})")
        except ClientError:
            # If that fails, try to find by unique name
            result = self.get(f"solutions?$filter=uniquename eq '{solution_id}'")
            solutions = result.get("value", [])
            if not solutions:
                raise ClientError(f"Solution not found: {solution_id}")
            return solutions[0]

    def get_solution_component_type(self, entity_logical_name: str) -> Optional[int]:
        """
        Get the solution component type for a given entity logical name.

        Args:
            entity_logical_name: The logical name of the entity (e.g., 'bot', 'connectionreference')

        Returns:
            The component type integer value, or None if not found
        """
        # Query solutioncomponentdefinitions to find the component type
        result = self.get(
            f"solutioncomponentdefinitions?$filter=primaryentityname eq '{entity_logical_name}'"
            "&$select=solutioncomponenttype,name,primaryentityname"
        )
        definitions = result.get("value", [])
        if definitions:
            return definitions[0].get("solutioncomponenttype")
        return None

    def get_dependencies(self, object_id: str, component_type: int) -> list[dict]:
        """
        Get dependencies that would prevent a component from being deleted.

        Args:
            object_id: The GUID of the component
            component_type: The solution component type integer

        Returns:
            List of dependency records with dependent component info
        """
        result = self.get(
            f"RetrieveDependenciesForDelete(ObjectId={object_id},ComponentType={component_type})"
        )
        return result.get("value", [])

    def get_dependencies_for_entity(self, object_id: str, entity_name: str) -> list[dict]:
        """
        Get dependencies for a component, resolving entity name to component type.

        Args:
            object_id: The GUID of the component
            entity_name: The logical name of the entity (e.g., 'connector', 'msdyn_aimodel')

        Returns:
            List of dependency records, or empty list if component type not found
        """
        component_type = self.get_solution_component_type(entity_name)
        if component_type is None:
            return []
        return self.get_dependencies(object_id, component_type)

    def add_solution_component(
        self,
        solution_unique_name: str,
        component_id: str,
        component_type: int,
        add_required_components: bool = False,
    ) -> dict:
        """
        Add a component to an unmanaged solution.

        Args:
            solution_unique_name: The unique name of the target solution
            component_id: The GUID of the component to add
            component_type: The component type integer value
            add_required_components: Whether to also add required dependencies

        Returns:
            Response from the AddSolutionComponent action
        """
        action_data = {
            "ComponentId": component_id,
            "ComponentType": component_type,
            "SolutionUniqueName": solution_unique_name,
            "AddRequiredComponents": add_required_components,
        }
        return self.post("AddSolutionComponent", action_data)

    def remove_solution_component(
        self,
        solution_unique_name: str,
        component_id: str,
        component_type: int,
    ) -> dict:
        """
        Remove a component from an unmanaged solution.

        Args:
            solution_unique_name: The unique name of the target solution
            component_id: The GUID of the component to remove (the actual entity ID, e.g., bot ID)
            component_type: The component type integer value

        Returns:
            Response from the RemoveSolutionComponent action

        Note:
            The RemoveSolutionComponent Web API action expects the component's actual
            entity ID (e.g., botid) passed as 'solutioncomponentid', not the
            solutioncomponent record ID. This is counterintuitive but required.
        """
        action_data = {
            "SolutionComponent": {
                "@odata.type": "Microsoft.Dynamics.CRM.solutioncomponent",
                "solutioncomponentid": component_id,
            },
            "ComponentType": component_type,
            "SolutionUniqueName": solution_unique_name,
        }
        return self.post("RemoveSolutionComponent", action_data)

    def get_solution_components(
        self,
        solution_id: str,
        component_type: Optional[int] = None,
    ) -> list[dict]:
        """
        Get components in a solution.

        Args:
            solution_id: The solution's unique identifier (GUID)
            component_type: Optional filter by component type

        Returns:
            List of solution component records
        """
        endpoint = f"solutioncomponents?$filter=_solutionid_value eq {solution_id}"
        if component_type is not None:
            endpoint += f" and componenttype eq {component_type}"
        result = self.get(endpoint)
        return result.get("value", [])

    def add_bot_to_solution(
        self,
        solution_unique_name: str,
        bot_id: str,
        add_required_components: bool = True,
    ) -> dict:
        """
        Add a Copilot agent (bot) to an unmanaged solution.

        Args:
            solution_unique_name: The unique name of the target solution
            bot_id: The bot's unique identifier (GUID)
            add_required_components: Whether to also add required dependencies (default: True)

        Returns:
            Response from the AddSolutionComponent action
        """
        # Get the component type for bot
        component_type = self.get_solution_component_type("bot")
        if component_type is None:
            raise ClientError("Could not determine component type for 'bot' entity")

        return self.add_solution_component(
            solution_unique_name=solution_unique_name,
            component_id=bot_id,
            component_type=component_type,
            add_required_components=add_required_components,
        )

    def remove_bot_from_solution(
        self,
        solution_unique_name: str,
        bot_id: str,
    ) -> dict:
        """
        Remove a Copilot agent (bot) from an unmanaged solution.

        Args:
            solution_unique_name: The unique name of the target solution
            bot_id: The bot's unique identifier (GUID)

        Returns:
            Response from the RemoveSolutionComponent action
        """
        # Get the component type for bot
        component_type = self.get_solution_component_type("bot")
        if component_type is None:
            raise ClientError("Could not determine component type for 'bot' entity")

        return self.remove_solution_component(
            solution_unique_name=solution_unique_name,
            component_id=bot_id,
            component_type=component_type,
        )

    def add_connection_reference_to_solution(
        self,
        solution_unique_name: str,
        connection_reference_id: str,
        add_required_components: bool = False,
    ) -> dict:
        """
        Add a connection reference to an unmanaged solution.

        Args:
            solution_unique_name: The unique name of the target solution
            connection_reference_id: The connection reference's unique identifier (GUID)
            add_required_components: Whether to also add required dependencies

        Returns:
            Response from the AddSolutionComponent action
        """
        # Get the component type for connectionreference
        component_type = self.get_solution_component_type("connectionreference")
        if component_type is None:
            raise ClientError("Could not determine component type for 'connectionreference' entity")

        return self.add_solution_component(
            solution_unique_name=solution_unique_name,
            component_id=connection_reference_id,
            component_type=component_type,
            add_required_components=add_required_components,
        )

    def remove_connection_reference_from_solution(
        self,
        solution_unique_name: str,
        connection_reference_id: str,
    ) -> dict:
        """
        Remove a connection reference from an unmanaged solution.

        Args:
            solution_unique_name: The unique name of the target solution
            connection_reference_id: The connection reference's unique identifier (GUID)

        Returns:
            Response from the RemoveSolutionComponent action
        """
        # Get the component type for connectionreference
        component_type = self.get_solution_component_type("connectionreference")
        if component_type is None:
            raise ClientError("Could not determine component type for 'connectionreference' entity")

        return self.remove_solution_component(
            solution_unique_name=solution_unique_name,
            component_id=connection_reference_id,
            component_type=component_type,
        )

    def get_bot_connection_reference(self, bot_id: str) -> Optional[dict]:
        """
        Get the provider connection reference for a specific bot.

        Args:
            bot_id: The bot's unique identifier

        Returns:
            Connection reference record, or None if not found
        """
        bot = self.get_bot(bot_id)
        provider_ref_id = bot.get("_providerconnectionreferenceid_value")
        if provider_ref_id:
            result = self.get(f"connectionreferences({provider_ref_id})")
            return result if result else None
        return None

    # =========================================================================
    # Publisher Methods
    # =========================================================================

    def list_publishers(self) -> list[dict]:
        """
        List all publishers in the environment.

        Returns:
            List of publisher records
        """
        result = self.get("publishers?$orderby=friendlyname")
        return result.get("value", [])

    def get_publisher(self, publisher_id: str) -> dict:
        """
        Get a publisher by ID or unique name.

        Args:
            publisher_id: The publisher's unique identifier (GUID) or unique name

        Returns:
            Publisher record
        """
        # Check if it's a GUID or unique name
        if self._is_guid(publisher_id):
            return self.get(f"publishers({publisher_id})")
        else:
            # Query by unique name
            result = self.get(f"publishers?$filter=uniquename eq '{publisher_id}'")
            publishers = result.get("value", [])
            if not publishers:
                raise ClientError(f"Publisher '{publisher_id}' not found")
            return publishers[0]

    def create_publisher(
        self,
        unique_name: str,
        friendly_name: str,
        customization_prefix: str,
        customization_option_value_prefix: int,
        description: Optional[str] = None,
    ) -> dict:
        """
        Create a new publisher.

        Args:
            unique_name: Unique name for the publisher (alphanumeric, no spaces)
            friendly_name: Display name for the publisher
            customization_prefix: Prefix for customizations (2-8 lowercase letters)
            customization_option_value_prefix: Prefix for option values (10000-99999)
            description: Optional description

        Returns:
            Created publisher record (from OData-EntityId header)
        """
        publisher_data = {
            "uniquename": unique_name,
            "friendlyname": friendly_name,
            "customizationprefix": customization_prefix,
            "customizationoptionvalueprefix": customization_option_value_prefix,
        }

        if description:
            publisher_data["description"] = description

        return self.post("publishers", publisher_data)

    def delete_publisher(self, publisher_id: str) -> None:
        """
        Delete a publisher by ID or unique name.

        Args:
            publisher_id: The publisher's unique identifier (GUID) or unique name

        Note:
            Publishers cannot be deleted if they have solutions associated with them.
        """
        # Resolve publisher ID if it's a unique name
        if not self._is_guid(publisher_id):
            publisher = self.get_publisher(publisher_id)
            publisher_id = publisher.get("publisherid")
            if not publisher_id:
                raise ClientError(f"Could not resolve publisher ID for '{publisher_id}'")

        self.delete(f"publishers({publisher_id})")

    # =========================================================================
    # Solution Creation Methods
    # =========================================================================

    def create_solution(
        self,
        unique_name: str,
        friendly_name: str,
        publisher_id: str,
        version: str = "1.0.0.0",
        description: Optional[str] = None,
    ) -> dict:
        """
        Create a new unmanaged solution.

        Args:
            unique_name: Unique name for the solution (alphanumeric, no spaces)
            friendly_name: Display name for the solution
            publisher_id: The publisher's unique identifier (GUID) or unique name
            version: Version string (default: "1.0.0.0")
            description: Optional description

        Returns:
            Created solution record (from OData-EntityId header)
        """
        # Resolve publisher ID if it's a unique name
        if not self._is_guid(publisher_id):
            publisher = self.get_publisher(publisher_id)
            publisher_id = publisher.get("publisherid")
            if not publisher_id:
                raise ClientError(f"Could not resolve publisher ID for '{publisher_id}'")

        solution_data = {
            "uniquename": unique_name,
            "friendlyname": friendly_name,
            "version": version,
            "publisherid@odata.bind": f"/publishers({publisher_id})",
        }

        if description:
            solution_data["description"] = description

        return self.post("solutions", solution_data)

    def delete_solution(self, solution_id: str) -> None:
        """
        Delete a solution by ID or unique name.

        Args:
            solution_id: The solution's unique identifier (GUID) or unique name
        """
        # Resolve to GUID if it's a unique name
        if not self._is_guid(solution_id):
            solution = self.get_solution(solution_id)
            solution_id = solution.get("solutionid")
            if not solution_id:
                raise ClientError(f"Could not resolve solution ID for '{solution_id}'")

        self.delete(f"solutions({solution_id})")

    def _is_guid(self, value: str) -> bool:
        """Check if a string is a valid GUID format."""
        import re
        guid_pattern = r'^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$'
        return bool(re.match(guid_pattern, value))

    # =========================================================================
    # Power Platform Connection Methods
    # =========================================================================

    def create_azure_ai_search_connection(
        self,
        connection_name: str,
        search_endpoint: str,
        api_key: str,
        environment_id: str,
    ) -> dict:
        """
        Create a Power Platform connection for Azure AI Search.

        This creates a connection that can be used by Copilot Studio to access
        Azure AI Search indexes. After creation, the connection must be linked
        to an agent as a knowledge source through the Copilot Studio UI.

        Args:
            connection_name: Display name for the connection
            search_endpoint: Azure AI Search endpoint URL (e.g., https://mysearch.search.windows.net)
            api_key: Azure AI Search admin or query key
            environment_id: Power Platform environment ID (e.g., Default-<tenant-id>)

        Returns:
            Dict containing connection details including:
            - name: Connection ID (GUID)
            - properties.displayName: Display name
            - properties.statuses: Connection status

        Raises:
            ClientError: If connection creation fails

        Note:
            This method creates the Power Platform connection but does NOT
            add it as a knowledge source to a Copilot agent. Copilot Studio
            requires the connection to be linked through the UI, as the
            knowledge source configuration involves complex internal linking
            that is not exposed via public APIs.

            After creating a connection, use the Copilot Studio UI to:
            1. Open your agent
            2. Go to Knowledge > Add knowledge
            3. Select Azure AI Search
            4. The connection you created will be available to select
            5. Specify the index name and complete the setup
        """
        import uuid

        # Generate a new connection ID
        connection_id = str(uuid.uuid4())

        # Get access token for Power Apps API
        powerapps_token = get_access_token_from_azure_cli("https://service.powerapps.com/")

        # Build the connection request
        connection_data = {
            "properties": {
                "environment": {
                    "id": f"/providers/Microsoft.PowerApps/environments/{environment_id}",
                    "name": environment_id
                },
                "displayName": connection_name,
                "connectionParametersSet": {
                    "name": "adminkey",
                    "values": {
                        "ConnectionEndpoint": {
                            "value": search_endpoint
                        },
                        "AdminKey": {
                            "value": api_key
                        }
                    }
                }
            }
        }

        # Create the connection via Power Apps API
        url = (
            f"https://api.powerapps.com/providers/Microsoft.PowerApps/apis/"
            f"shared_azureaisearch/connections/{connection_id}"
            f"?api-version=2016-11-01&$filter=environment%20eq%20%27{environment_id}%27"
        )

        headers = {
            "Authorization": f"Bearer {powerapps_token}",
            "Content-Type": "application/json",
            "Accept": "application/json",
        }

        try:
            response = self._http_client.put(url, headers=headers, json=connection_data, timeout=60.0)
            response.raise_for_status()
            return response.json()
        except httpx.HTTPStatusError as e:
            error_detail = ""
            try:
                error_body = e.response.json()
                if "error" in error_body:
                    error_detail = error_body["error"].get("message", str(error_body))
            except Exception:
                error_detail = e.response.text[:500] if e.response.text else str(e)
            raise ClientError(f"Failed to create connection: HTTP {e.response.status_code}: {error_detail}")
        except httpx.RequestError as e:
            raise ClientError(f"Connection request failed: {e}")

    def list_connection_references(self) -> list[dict]:
        """
        List all connection references in the Dataverse environment.

        Connection references are solution-aware references to connections
        used in flows and agents.

        Returns:
            List of connection reference objects
        """
        select = (
            "connectionreferenceid,connectionreferencelogicalname,"
            "connectionreferencedisplayname,connectorid,connectionid,statecode"
        )
        url = f"{self.api_url}/connectionreferences?$select={select}&$orderby=connectionreferencedisplayname"
        headers = self._get_headers()
        response = self._http_client.get(url, headers=headers, timeout=60.0)
        response.raise_for_status()
        data = response.json()
        return data.get("value", [])

    def delete_connection_reference(self, connection_reference_id: str) -> bool:
        """
        Delete a connection reference from the Dataverse environment.

        Args:
            connection_reference_id: The connection reference's unique identifier (GUID)

        Returns:
            True if deletion was successful

        Raises:
            ClientError: If the connection reference cannot be deleted
        """
        url = f"{self.api_url}/connectionreferences({connection_reference_id})"
        headers = self._get_headers()
        response = self._http_client.delete(url, headers=headers, timeout=60.0)
        response.raise_for_status()
        return True

    def update_connection_reference(
        self,
        connection_reference_id: str,
        connection_id: Optional[str] = None,
        display_name: Optional[str] = None,
    ) -> dict:
        """
        Update a connection reference in the Dataverse environment.

        Args:
            connection_reference_id: The connection reference's unique identifier (GUID)
            connection_id: New connection ID to associate with this reference
            display_name: New display name for the connection reference

        Returns:
            Updated connection reference object

        Raises:
            ClientError: If the connection reference cannot be updated
        """
        url = f"{self.api_url}/connectionreferences({connection_reference_id})"
        headers = self._get_headers()

        # Build update payload with only provided fields
        payload = {}
        if connection_id is not None:
            payload["connectionid"] = connection_id
        if display_name is not None:
            payload["connectionreferencedisplayname"] = display_name

        if not payload:
            raise ClientError("No update fields provided")

        response = self._http_client.patch(url, headers=headers, json=payload, timeout=60.0)
        response.raise_for_status()

        # Fetch and return the updated record
        get_response = self._http_client.get(url, headers=headers, timeout=60.0)
        get_response.raise_for_status()
        return get_response.json()

    def get_connection_reference(self, connection_reference_id: str) -> dict:
        """
        Get a single connection reference by ID.

        Args:
            connection_reference_id: The connection reference's unique identifier (GUID)

        Returns:
            Connection reference object

        Raises:
            ClientError: If the connection reference is not found
        """
        url = f"{self.api_url}/connectionreferences({connection_reference_id})"
        headers = self._get_headers()
        response = self._http_client.get(url, headers=headers, timeout=60.0)
        response.raise_for_status()
        return response.json()

    def create_connection_reference(
        self,
        display_name: str,
        connector_id: str,
        connection_id: Optional[str] = None,
        description: Optional[str] = None,
    ) -> dict:
        """
        Create a new connection reference in the Dataverse environment.

        Connection references are solution-aware pointers to connections,
        allowing flows and agents to reference connections without being
        directly tied to them.

        Args:
            display_name: Display name for the connection reference
            connector_id: Connector identifier (e.g., 'shared_asana' or full path
                          '/providers/Microsoft.PowerApps/apis/shared_asana')
            connection_id: Optional connection ID to link to an existing connection
            description: Optional description for the connection reference

        Returns:
            Created connection reference object

        Raises:
            ClientError: If the connection reference cannot be created
        """
        # Normalize connector_id to full path format if not already
        if not connector_id.startswith("/"):
            connector_id = f"/providers/Microsoft.PowerApps/apis/{connector_id}"

        # Generate logical name from display name (lowercase, alphanumeric + underscore)
        import re
        logical_name = re.sub(r"[^a-z0-9_]", "_", display_name.lower())
        # Add prefix to ensure uniqueness
        logical_name = f"cr_{logical_name}"

        payload = {
            "connectionreferencedisplayname": display_name,
            "connectionreferencelogicalname": logical_name,
            "connectorid": connector_id,
        }

        if connection_id is not None:
            payload["connectionid"] = connection_id

        if description is not None:
            payload["description"] = description

        url = f"{self.api_url}/connectionreferences"
        headers = self._get_headers()

        try:
            response = self._http_client.post(url, headers=headers, json=payload, timeout=60.0)
            response.raise_for_status()

            # Extract the created reference ID from response headers
            # OData-EntityId header contains the URL with the new ID
            entity_id_header = response.headers.get("OData-EntityId", "")
            if entity_id_header:
                # Extract GUID from URL like https://.../connectionreferences(guid)
                import re as re_module
                match = re_module.search(r"connectionreferences\(([^)]+)\)", entity_id_header)
                if match:
                    created_id = match.group(1)
                    # Fetch the created record
                    return self.get_connection_reference(created_id)

            # If we can't get the ID from headers, return what we can
            return payload

        except httpx.HTTPStatusError as e:
            error_detail = str(e)
            if e.response is not None:
                try:
                    error_body = e.response.json()
                    if "error" in error_body:
                        error_detail = error_body["error"].get("message", str(error_body))
                except Exception:
                    error_detail = e.response.text[:500] if e.response.text else str(e)
            raise ClientError(f"Failed to create connection reference: HTTP {e.response.status_code}: {error_detail}")
        except httpx.RequestError as e:
            raise ClientError(f"Connection reference request failed: {e}")

    def list_connections(
        self, connector_id: Optional[str] = None, environment_id: Optional[str] = None
    ) -> list[dict]:
        """
        List connections in the environment.

        If connector_id is provided, lists connections for that specific connector.
        If connector_id is not provided, lists all connections across all connectors
        using the admin API.

        Args:
            connector_id: Optional connector identifier (e.g., shared_office365).
                          If not provided, returns all connections.
            environment_id: Power Platform environment ID. If not provided,
                            will use DATAVERSE_ENVIRONMENT_ID from config.

        Returns:
            List of connection objects
        """
        # Get environment ID from config if not provided
        if not environment_id:
            config = get_config()
            environment_id = config.environment_id
            if not environment_id:
                raise ClientError(
                    "Environment ID not found. Please set DATAVERSE_ENVIRONMENT_ID "
                    "in your .env file (e.g., Default-<tenant-id> or the environment GUID)."
                )

        powerapps_token = get_access_token_from_azure_cli("https://service.powerapps.com/")

        # Use different endpoints based on whether connector_id is provided
        if connector_id:
            # Connector-specific endpoint
            url = (
                f"https://api.powerapps.com/providers/Microsoft.PowerApps/apis/"
                f"{connector_id}/connections"
                f"?api-version=2016-11-01&$filter=environment%20eq%20%27{environment_id}%27"
            )
        else:
            # Admin endpoint to get all connections
            url = (
                f"https://api.powerapps.com/providers/Microsoft.PowerApps/scopes/admin/"
                f"environments/{environment_id}/connections"
                f"?api-version=2016-11-01"
            )

        headers = {
            "Authorization": f"Bearer {powerapps_token}",
            "Accept": "application/json",
        }

        try:
            response = self._http_client.get(url, headers=headers, timeout=30.0)
            response.raise_for_status()
            data = response.json()
            return data.get("value", [])
        except httpx.HTTPStatusError as e:
            error_detail = ""
            try:
                error_body = e.response.json()
                if "error" in error_body:
                    error_detail = error_body["error"].get("message", str(error_body))
            except Exception:
                error_detail = e.response.text[:500] if e.response.text else str(e)
            raise ClientError(f"Failed to list connections: HTTP {e.response.status_code}: {error_detail}")
        except httpx.RequestError as e:
            raise ClientError(f"Request failed: {e}")

    def get_connection(
        self, connection_id: str, environment_id: Optional[str] = None
    ) -> dict:
        """
        Get a specific connection by ID.

        Searches all connections in the environment to find the one with
        the matching connection ID.

        Args:
            connection_id: The connection's unique identifier (GUID)
            environment_id: Power Platform environment ID. If not provided,
                            will use DATAVERSE_ENVIRONMENT_ID from config.

        Returns:
            Connection object with properties including connector ID

        Raises:
            ClientError: If the connection is not found
        """
        # List all connections and find the matching one
        connections = self.list_connections(environment_id=environment_id)

        for conn in connections:
            if conn.get("name") == connection_id:
                return conn

        raise ClientError(f"Connection '{connection_id}' not found")

    def test_connection(
        self, connector_id: str, connection_id: str, environment_id: Optional[str] = None
    ) -> dict:
        """
        Test authentication for a specific connection.

        Args:
            connector_id: The connector's unique identifier (e.g., shared_office365)
            connection_id: The connection's unique identifier (GUID)
            environment_id: Power Platform environment ID. If not provided,
                            will use DATAVERSE_ENVIRONMENT_ID from config.

        Returns:
            Test result with status information
        """
        # Get environment ID from config if not provided
        if not environment_id:
            config = get_config()
            environment_id = config.environment_id
            if not environment_id:
                raise ClientError(
                    "Environment ID not found. Please set DATAVERSE_ENVIRONMENT_ID "
                    "in your .env file (e.g., Default-<tenant-id> or the environment GUID)."
                )

        powerapps_token = get_access_token_from_azure_cli("https://service.powerapps.com/")

        url = (
            f"https://api.powerapps.com/providers/Microsoft.PowerApps/apis/"
            f"{connector_id}/connections/{connection_id}/testConnection"
            f"?api-version=2016-11-01"
        )

        headers = {
            "Authorization": f"Bearer {powerapps_token}",
            "Accept": "application/json",
            "Content-Type": "application/json",
        }

        try:
            response = self._http_client.post(url, headers=headers, json={}, timeout=60.0)
            # Test connection can return various status codes
            # 200 = success, 401/403 = auth failed, etc.
            result = {
                "status_code": response.status_code,
                "success": response.status_code == 200,
            }
            try:
                result["response"] = response.json()
            except Exception:
                result["response"] = response.text[:500] if response.text else ""
            return result
        except httpx.HTTPStatusError as e:
            error_detail = ""
            try:
                error_body = e.response.json()
                if "error" in error_body:
                    error_detail = error_body["error"].get("message", str(error_body))
            except Exception:
                error_detail = e.response.text[:500] if e.response.text else str(e)
            return {
                "status_code": e.response.status_code,
                "success": False,
                "error": error_detail or f"HTTP {e.response.status_code}",
            }
        except httpx.RequestError as e:
            return {
                "status_code": 0,
                "success": False,
                "error": f"Request failed: {e}",
            }

    def list_azure_ai_search_connections(self, environment_id: str) -> list[dict]:
        """
        List Azure AI Search connections in a Power Platform environment.

        Args:
            environment_id: Power Platform environment ID

        Returns:
            List of connection objects

        Note:
            This is a convenience method. Use list_connections() for any connector.
        """
        return self.list_connections("shared_azureaisearch", environment_id)

    def _list_azure_ai_search_connections_legacy(self, environment_id: str) -> list[dict]:
        """
        Legacy implementation - kept for reference.
        """
        powerapps_token = get_access_token_from_azure_cli("https://service.powerapps.com/")

        url = (
            f"https://api.powerapps.com/providers/Microsoft.PowerApps/apis/"
            f"shared_azureaisearch/connections"
            f"?api-version=2016-11-01&$filter=environment%20eq%20%27{environment_id}%27"
        )

        headers = {
            "Authorization": f"Bearer {powerapps_token}",
            "Accept": "application/json",
        }

        try:
            response = self._http_client.get(url, headers=headers, timeout=30.0)
            response.raise_for_status()
            data = response.json()
            return data.get("value", [])
        except httpx.HTTPStatusError as e:
            error_detail = ""
            try:
                error_body = e.response.json()
                if "error" in error_body:
                    error_detail = error_body["error"].get("message", str(error_body))
            except Exception:
                error_detail = e.response.text[:500] if e.response.text else str(e)
            raise ClientError(f"Failed to list connections: HTTP {e.response.status_code}: {error_detail}")
        except httpx.RequestError as e:
            raise ClientError(f"Request failed: {e}")

    def delete_connection(
        self, connection_id: str, connector_id: str, environment_id: str
    ) -> None:
        """
        Delete a Power Platform connection.

        Args:
            connection_id: The connection's unique identifier (GUID)
            connector_id: The connector's unique identifier (e.g., shared_asana, shared_office365)
            environment_id: Power Platform environment ID
        """
        powerapps_token = get_access_token_from_azure_cli("https://service.powerapps.com/")

        url = (
            f"https://api.powerapps.com/providers/Microsoft.PowerApps/apis/"
            f"{connector_id}/connections/{connection_id}"
            f"?api-version=2016-11-01&$filter=environment%20eq%20%27{environment_id}%27"
        )

        headers = {
            "Authorization": f"Bearer {powerapps_token}",
            "Accept": "application/json",
        }

        try:
            response = self._http_client.delete(url, headers=headers, timeout=30.0)
            response.raise_for_status()
        except httpx.HTTPStatusError as e:
            error_detail = ""
            try:
                error_body = e.response.json()
                if "error" in error_body:
                    error_detail = error_body["error"].get("message", str(error_body))
            except Exception:
                error_detail = e.response.text[:500] if e.response.text else str(e)
            raise ClientError(f"Failed to delete connection: HTTP {e.response.status_code}: {error_detail}")
        except httpx.RequestError as e:
            raise ClientError(f"Request failed: {e}")

    def create_connection(
        self,
        connector_id: str,
        connection_name: str,
        environment_id: str,
        parameters: Optional[dict] = None,
    ) -> dict:
        """
        Create a Power Platform connection for any connector.

        Args:
            connector_id: The connector's unique identifier (e.g., shared_asana)
            connection_name: Display name for the connection
            environment_id: Power Platform environment ID
            parameters: Connection parameters (connector-specific)

        Returns:
            Dict containing connection details

        Raises:
            ClientError: If connection creation fails
        """
        import uuid

        connection_id = str(uuid.uuid4())
        powerapps_token = get_access_token_from_azure_cli("https://service.powerapps.com/")

        # Build the connection request
        connection_data = {
            "properties": {
                "environment": {
                    "id": f"/providers/Microsoft.PowerApps/environments/{environment_id}",
                    "name": environment_id
                },
                "displayName": connection_name,
            }
        }

        # Add parameters if provided
        if parameters:
            connection_data["properties"]["connectionParameters"] = parameters

        url = (
            f"https://api.powerapps.com/providers/Microsoft.PowerApps/apis/"
            f"{connector_id}/connections/{connection_id}"
            f"?api-version=2016-11-01&$filter=environment%20eq%20%27{environment_id}%27"
        )

        headers = {
            "Authorization": f"Bearer {powerapps_token}",
            "Content-Type": "application/json",
            "Accept": "application/json",
        }

        try:
            response = self._http_client.put(url, headers=headers, json=connection_data, timeout=60.0)
            response.raise_for_status()
            return response.json()
        except httpx.HTTPStatusError as e:
            error_detail = ""
            try:
                error_body = e.response.json()
                if "error" in error_body:
                    error_detail = error_body["error"].get("message", str(error_body))
            except Exception:
                error_detail = e.response.text[:500] if e.response.text else str(e)
            raise ClientError(f"Failed to create connection: HTTP {e.response.status_code}: {error_detail}")
        except httpx.RequestError as e:
            raise ClientError(f"Connection request failed: {e}")

    def create_oauth_connection(
        self,
        connector_id: str,
        connection_name: str,
        environment_id: str,
    ) -> dict:
        """
        Create a Power Platform connection for an OAuth-based connector.

        This creates the connection record but the user must complete the
        OAuth consent flow via the returned URL or Power Platform portal.

        Args:
            connector_id: The connector's unique identifier (e.g., shared_asana)
            connection_name: Display name for the connection
            environment_id: Power Platform environment ID

        Returns:
            Dict containing connection details and consent URL if available

        Raises:
            ClientError: If connection creation fails
        """
        import uuid

        connection_id = str(uuid.uuid4())
        powerapps_token = get_access_token_from_azure_cli("https://service.powerapps.com/")

        # OAuth connections require minimal initial parameters
        connection_data = {
            "properties": {
                "environment": {
                    "id": f"/providers/Microsoft.PowerApps/environments/{environment_id}",
                    "name": environment_id
                },
                "displayName": connection_name,
            }
        }

        url = (
            f"https://api.powerapps.com/providers/Microsoft.PowerApps/apis/"
            f"{connector_id}/connections/{connection_id}"
            f"?api-version=2016-11-01&$filter=environment%20eq%20%27{environment_id}%27"
        )

        headers = {
            "Authorization": f"Bearer {powerapps_token}",
            "Content-Type": "application/json",
            "Accept": "application/json",
        }

        try:
            response = self._http_client.put(url, headers=headers, json=connection_data, timeout=60.0)
            response.raise_for_status()
            return response.json()
        except httpx.HTTPStatusError as e:
            error_detail = ""
            try:
                error_body = e.response.json()
                if "error" in error_body:
                    error_detail = error_body["error"].get("message", str(error_body))
            except Exception:
                error_detail = e.response.text[:500] if e.response.text else str(e)
            raise ClientError(f"Failed to create connection: HTTP {e.response.status_code}: {error_detail}")
        except httpx.RequestError as e:
            raise ClientError(f"Connection request failed: {e}")

    def get_consent_link(
        self,
        connector_id: str,
        connection_id: str,
        environment_id: str,
    ) -> str:
        """
        Get the OAuth consent link for an unauthenticated connection.

        Args:
            connector_id: The connector's unique identifier (e.g., shared_asana)
            connection_id: The connection's unique identifier (GUID)
            environment_id: Power Platform environment ID

        Returns:
            The consent URL to complete OAuth authentication

        Raises:
            ClientError: If getting consent link fails
        """
        powerapps_token = get_access_token_from_azure_cli("https://service.powerapps.com/")

        url = (
            f"https://api.powerapps.com/providers/Microsoft.PowerApps/apis/"
            f"{connector_id}/connections/{connection_id}/getConsentLink"
            f"?api-version=2016-11-01&$filter=environment eq '{environment_id}'"
        )

        headers = {
            "Authorization": f"Bearer {powerapps_token}",
            "Content-Type": "application/json",
            "Accept": "application/json",
        }

        body = {
            "redirectUrl": "https://make.powerapps.com"
        }

        try:
            response = self._http_client.post(url, headers=headers, json=body, timeout=30.0)
            response.raise_for_status()
            data = response.json()
            return data.get("consentLink", "")
        except httpx.HTTPStatusError as e:
            error_detail = ""
            try:
                error_body = e.response.json()
                if "error" in error_body:
                    error_detail = error_body["error"].get("message", str(error_body))
            except Exception:
                error_detail = e.response.text[:500] if e.response.text else str(e)
            raise ClientError(f"Failed to get consent link: HTTP {e.response.status_code}: {error_detail}")
        except httpx.RequestError as e:
            raise ClientError(f"Consent link request failed: {e}")

    def get_connection_user(
        self,
        connector_id: str,
        connection_id: str,
    ) -> dict:
        """
        Get the authenticated user for a connection by calling the connector's API.

        This works by calling a user/me endpoint through the connector to retrieve
        the authenticated user's information from the external service.

        Args:
            connector_id: The connector's unique identifier (e.g., shared_asana)
            connection_id: The connection's unique identifier (GUID)

        Returns:
            Dict with user info (varies by connector), or empty dict if not supported

        Raises:
            ClientError: If the request fails
        """
        apihub_token = get_access_token_from_azure_cli("https://apihub.azure.com")

        # Map connectors to their user info endpoints
        user_endpoints = {
            "shared_asana": "/v2/users/me",
            "shared_office365": "/v2/Me",
            "shared_sharepointonline": "/_api/web/currentuser",
            "shared_dynamicscrmonline": "/api/data/v9.2/WhoAmI",
        }

        # Get the appropriate endpoint for this connector
        endpoint = user_endpoints.get(connector_id)
        if not endpoint:
            return {}

        url = f"https://msmanaged-na.azure-apim.net/apim/{connector_id.replace('shared_', '')}/{connection_id}{endpoint}"

        headers = {
            "Authorization": f"Bearer {apihub_token}",
            "Accept": "application/json",
        }

        try:
            response = self._http_client.get(url, headers=headers, timeout=30.0)
            response.raise_for_status()
            return response.json()
        except Exception:
            # Silently return empty if we can't get user info
            return {}

    def close(self):
        """Close the HTTP client."""
        self._http_client.close()


# Global client instance
_client: Optional[DataverseClient] = None


def get_access_token_from_azure_cli(resource: str) -> str:
    """
    Get an access token using Azure CLI.

    Args:
        resource: The resource URL to get a token for

    Returns:
        Access token string

    Raises:
        ClientError: If token acquisition fails
    """
    try:
        result = subprocess.run(
            ["az", "account", "get-access-token", "--resource", resource, "--query", "accessToken", "-o", "tsv"],
            capture_output=True,
            text=True,
            check=True,
        )
        return result.stdout.strip()
    except subprocess.CalledProcessError as e:
        raise ClientError(
            f"Failed to get access token from Azure CLI. "
            f"Make sure you're logged in with 'az login'. Error: {e.stderr}"
        )
    except FileNotFoundError:
        raise ClientError(
            "Azure CLI not found. Please install Azure CLI and login with 'az login'."
        )


def get_client() -> DataverseClient:
    """
    Get or create the global Dataverse client.

    Uses Azure CLI authentication by default (requires 'az login').

    Returns:
        DataverseClient: Authenticated Dataverse API client

    Raises:
        ClientError: If credentials are missing or authentication fails
    """
    global _client

    if _client is not None:
        return _client

    config = get_config()

    # Check for missing credentials
    missing = config.get_missing_credentials()
    if missing:
        error_msg = (
            "Missing required credentials. Please set the following "
            "environment variables in your .env file:\n\n"
        )
        for cred in missing:
            error_msg += f"  - {cred}\n"

        error_msg += "\nRequired:\n"
        error_msg += "  DATAVERSE_URL - Your Dataverse environment URL (e.g., https://org1cb52429.crm.dynamics.com)\n"
        error_msg += "\nAuthentication is handled via Azure CLI (requires 'az login').\n"

        raise ClientError(error_msg)

    dataverse_url = config.dataverse_url

    # Always use Azure CLI authentication
    try:
        access_token = get_access_token_from_azure_cli(dataverse_url)
        _client = DataverseClient(dataverse_url, access_token)
        return _client
    except Exception as e:
        raise ClientError(f"Failed to authenticate with Azure CLI: {e}")


def reset_client():
    """Reset the global client instance (useful for testing)."""
    global _client
    if _client is not None:
        _client.close()
    _client = None
