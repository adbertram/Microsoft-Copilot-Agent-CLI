"""Dataverse API client for Copilot Studio agents."""
import subprocess
import json
import re
from typing import Optional, Any
import httpx
from .config import get_config


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

    def delete(self, endpoint: str) -> None:
        """Make a DELETE request."""
        self._request("DELETE", endpoint)

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
        include_system: bool = True,
        include_custom: bool = True,
    ) -> list[dict]:
        """
        List topics for a specific bot.

        Args:
            bot_id: The bot's unique identifier
            include_system: Include system topics (default: True)
            include_custom: Include custom topics (default: True)

        Returns:
            List of topic component records with metadata

        Note:
            Topic component types:
            - 0 = Topic (legacy)
            - 9 = Topic (V2)

            System topics typically have category 'SYSTEM' or start with
            system-related schema names.
        """
        # Query topics (componenttype 0 = Topic, 9 = Topic V2)
        result = self.get(
            f"botcomponents?$filter=_parentbotid_value eq {bot_id} "
            f"and (componenttype eq 0 or componenttype eq 9)"
            f"&$orderby=name"
        )
        topics = result.get("value", [])

        # Categorize topics as system or custom
        categorized = []
        for topic in topics:
            # Determine if system topic based on category or schema name patterns
            category = (topic.get("category") or "").upper()
            schema_name = (topic.get("schemaname") or "").lower()
            name = topic.get("name") or ""

            # System topic indicators
            is_system = (
                category == "SYSTEM"
                or schema_name.startswith("system.")
                or name.startswith("System:")
                or name in [
                    "Greeting",
                    "Goodbye",
                    "Escalate",
                    "Start over",
                    "Thank you",
                    "Fallback",
                    "Multiple topics matched",
                    "Conversational boosting",
                    "On Error",
                ]
            )

            topic["_is_system"] = is_system

            # Filter based on include flags
            if is_system and include_system:
                categorized.append(topic)
            elif not is_system and include_custom:
                categorized.append(topic)

        return categorized

    def delete_bot(self, bot_id: str) -> None:
        """
        Delete a bot by ID.

        Args:
            bot_id: The bot's unique identifier
        """
        self.delete(f"bots({bot_id})")

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

    def list_connection_references(self, bot_id: Optional[str] = None) -> list[dict]:
        """
        List connection references, optionally filtered by bot.

        Args:
            bot_id: Optional bot ID to filter connection references linked to a specific bot

        Returns:
            List of connection reference records
        """
        if bot_id:
            # Get the bot's provider connection reference
            bot = self.get_bot(bot_id)
            provider_ref_id = bot.get("_providerconnectionreferenceid_value")
            if provider_ref_id:
                result = self.get(f"connectionreferences({provider_ref_id})")
                return [result] if result else []
            return []
        else:
            result = self.get("connectionreferences")
            return result.get("value", [])

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

    def list_azure_ai_search_connections(self, environment_id: str) -> list[dict]:
        """
        List Azure AI Search connections in a Power Platform environment.

        Args:
            environment_id: Power Platform environment ID

        Returns:
            List of connection objects
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

    def delete_connection(self, connection_id: str, environment_id: str) -> None:
        """
        Delete a Power Platform connection.

        Args:
            connection_id: The connection's unique identifier (GUID)
            environment_id: Power Platform environment ID
        """
        powerapps_token = get_access_token_from_azure_cli("https://service.powerapps.com/")

        url = (
            f"https://api.powerapps.com/providers/Microsoft.PowerApps/apis/"
            f"shared_azureaisearch/connections/{connection_id}"
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
