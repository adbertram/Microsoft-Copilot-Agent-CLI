# Copilot CLI Guide

Complete reference for the `copilot` command-line interface for managing Microsoft Copilot Studio agents via the Dataverse API.

## Overview

The Copilot CLI provides access to:
- **Agents** - Create, update, delete, publish, and test agents
- **Topics** - Manage conversation flows (list, create, update, delete, enable/disable)
- **Agent Tools** - Connect sub-agents as tools for orchestration
- **Knowledge** - Add file-based and Azure AI Search knowledge sources
- **Analytics** - Query Application Insights telemetry for troubleshooting
- **Transcripts** - View conversation history for debugging
- **Connections** - Manage Power Platform connections
- **Tools** - Unified tool management (prompts, connectors, REST APIs, MCP servers)
- **Solutions** - Manage solutions, publishers, and solution components
- **Flows** - List and view Power Automate cloud flows
- **Environments** - List and view Power Platform environments

## Authentication

The CLI uses Azure CLI authentication to obtain tokens for the Dataverse API.

### Prerequisites

1. Install Azure CLI and authenticate:
```bash
az login
```

2. Set the Dataverse URL environment variable:
```bash
export DATAVERSE_URL=https://org1cb52429.api.crm.dynamics.com
```

Or add to your shell profile (`~/.zshrc`, `~/.bashrc`):
```bash
export DATAVERSE_URL=https://org1cb52429.api.crm.dynamics.com
```

---

## Agent Commands

### List Agents

```bash
copilot agent list                      # List all agents (JSON)
copilot agent list --table              # List as formatted table
copilot agent list -t                   # Short form
```

### Get Agent Details

```bash
copilot agent get <bot-id>              # Get agent details
copilot agent get <bot-id> --components # Include all components (topics, tools, knowledge)
```

### Create Agent

```bash
copilot agent create --name "My Agent"
copilot agent create --name "My Agent" --description "A helpful assistant"
copilot agent create --name "My Agent" --instructions "You are a helpful assistant"
copilot agent create --name "My Agent" --instructions-file ./prompt.txt
copilot agent create --name "My Agent" --no-orchestration
```

**Options:**
| Option | Description |
|--------|-------------|
| `-n, --name` | Display name for the agent (required) |
| `-d, --description` | Description for the agent |
| `-i, --instructions` | System instructions/prompt |
| `--instructions-file` | Path to file containing instructions |
| `--orchestration/--no-orchestration` | Enable/disable generative AI orchestration |

### Update Agent

```bash
copilot agent update <bot-id> --name "New Name"
copilot agent update <bot-id> --description "New description"
copilot agent update <bot-id> --instructions "New system prompt"
copilot agent update <bot-id> --instructions-file ./prompt.txt
copilot agent update <bot-id> --no-orchestration
```

**Options:**
| Option | Description |
|--------|-------------|
| `-n, --name` | New display name |
| `-d, --description` | New description |
| `-i, --instructions` | New system instructions |
| `--instructions-file` | Path to file containing new instructions |
| `--orchestration/--no-orchestration` | Enable/disable orchestration |

### Publish Agent

```bash
copilot agent publish <bot-id>          # Make latest changes live
```

**Note:** Changes to agents are not live until published.

### Delete Agent

```bash
copilot agent remove <bot-id>           # Delete (with confirmation)
copilot agent remove <bot-id> --force   # Delete without confirmation
```

### Test Agent (Send Prompt)

Send a message to an agent and get a response. Requires Direct Line secret or Entra ID authentication.

```bash
# Using Direct Line secret
copilot agent prompt <bot-id> --message "Hello" --secret "your-secret"

# Using environment variable
export DIRECTLINE_SECRET=your-secret
copilot agent prompt <bot-id> -m "Hello"

# Using Entra ID authentication
copilot agent prompt <bot-id> -m "Hello" --entra-id \
    --client-id <app-client-id> --tenant-id <tenant-id> \
    --token-endpoint "https://{ENV}.environment.api.powerplatform.com/..."

# With file attachment
copilot agent prompt <bot-id> -m "Review this document" --file ./draft.docx --secret "xxx"

# Verbose output with JSON response
copilot agent prompt <bot-id> -m "Hello" -s "xxx" --verbose --json
```

**Options:**
| Option | Description |
|--------|-------------|
| `-m, --message` | The message/prompt to send (required) |
| `-s, --secret` | Direct Line secret |
| `--entra-id` | Use Entra ID authentication |
| `--client-id` | Entra ID application client ID |
| `--tenant-id` | Entra ID tenant ID |
| `--token-endpoint` | Bot token endpoint URL |
| `-f, --file` | Path to file attachment |
| `-v, --verbose` | Show detailed progress |
| `-j, --json` | Output as JSON |
| `--timeout` | Total timeout in seconds (default: 120) |
| `--max-polls` | Maximum polling attempts (default: 30) |
| `--poll-interval` | Seconds between polls (default: 3) |

**Environment Variables:**
- `DIRECTLINE_SECRET` - Direct Line secret
- `ENTRA_CLIENT_ID` - Entra ID client ID
- `ENTRA_TENANT_ID` - Entra ID tenant ID
- `ENTRA_SCOPE` - OAuth scope (default: https://api.powerplatform.com/.default)
- `BOT_TOKEN_ENDPOINT` - Bot token endpoint

---

## Topic Commands

Topics define conversation flows using AdaptiveDialog YAML format.

### List Topics

```bash
copilot agent topic list --agentId <agent-id>           # List all topics
copilot agent topic list --agentId <agent-id> --table   # Formatted table
copilot agent topic list --agentId <agent-id> --system  # System topics only
copilot agent topic list --agentId <agent-id> --custom  # Custom topics only
```

**Options:**
| Option | Description |
|--------|-------------|
| `-a, --agentId` | Agent's unique identifier (required) |
| `-t, --table` | Display as formatted table |
| `-s, --system` | List only system (managed) topics |
| `-c, --custom` | List only custom (user-created) topics |

### Get Topic

```bash
copilot agent topic get <topic-id>              # Get topic details (JSON)
copilot agent topic get <topic-id> --yaml       # Output YAML content only
copilot agent topic get <topic-id> -o file.yaml # Save YAML to file
```

**Options:**
| Option | Description |
|--------|-------------|
| `-y, --yaml` | Output topic content as YAML |
| `-o, --output` | Write YAML content to file |

### Create Topic

```bash
# Create from YAML file
copilot agent topic create --agentId <agent-id> --name "My Topic" --file topic.yaml

# Create simple topic with triggers and message
copilot agent topic create --agentId <agent-id> --name "Greeting" \
    --triggers "hello,hi,hey there" --message "Hello! How can I help?"
```

**Options:**
| Option | Description |
|--------|-------------|
| `-a, --agentId` | Agent's unique identifier (required) |
| `-n, --name` | Display name for the topic (required) |
| `-f, --file` | Path to YAML file containing topic content |
| `-t, --triggers` | Comma-separated trigger phrases |
| `-m, --message` | Response message (for simple topics) |
| `-d, --description` | Optional description |

### Update Topic

```bash
copilot agent topic update <topic-id> --file updated.yaml
copilot agent topic update <topic-id> --name "New Name"
copilot agent topic update <topic-id> --triggers "new,phrases" --message "New response"
copilot agent topic update <topic-id> --description "Updated description"
```

### Enable/Disable Topic

```bash
copilot agent topic enable <topic-id>
copilot agent topic disable <topic-id>
```

### Delete Topic

```bash
copilot agent topic delete <topic-id>           # Delete (with confirmation)
copilot agent topic delete <topic-id> --force   # Delete without confirmation
```

---

## Agent Tool Commands

Tools extend an agent's capabilities by allowing it to invoke external operations during orchestration. Supported tool types:
- **Connector** - Power Platform connector operations (e.g., SharePoint, Outlook, Dynamics)
- **Prompt** - AI Builder prompts for text generation and analysis
- **Flow** - Power Automate flows for complex automation
- **HTTP** - Direct HTTP requests to external APIs
- **Agent** - Other Copilot agents as sub-agents

### List Agent Tools

```bash
copilot agent tool list --agentId <agent-id>              # List all tools
copilot agent tool list --agentId <agent-id> --table      # Formatted table
copilot agent tool list --agentId <agent-id> --category agent  # Only connected agents
```

**Options:**
| Option | Description |
|--------|-------------|
| `-a, --agentId` | Agent's unique identifier (required) |
| `-t, --table` | Display as formatted table |
| `--category` | Filter by category (e.g., `agent`) |

### Add Tool

Add tools of any type to an agent using the unified interface:

```bash
copilot agent tool add --agentId <agent-id> --toolType <type> --id <tool-id> [options]
```

**Core Options:**
| Option | Description |
|--------|-------------|
| `-a, --agentId` | Agent's unique identifier (required) |
| `-T, --toolType` | Tool type: `connector`, `prompt`, `flow`, `http`, `agent` (required) |
| `--id` | Tool identifier - format depends on tool type (required) |
| `-n, --name` | Display name for the tool |
| `-d, --description` | Description for AI orchestration |
| `--inputs` | JSON string defining input parameters |
| `--outputs` | JSON string defining output parameters |

**Type-Specific Options:**
| Option | Applies To | Description |
|--------|------------|-------------|
| `--connection-ref` | connector, flow | Connection reference name |
| `--no-history` | agent | Don't pass conversation history |
| `--method` | http | HTTP method (GET, POST, etc.) |
| `--headers` | http | JSON string of HTTP headers |
| `--body` | http | HTTP request body template |

#### Tool Type: Connector

Invoke Power Platform connector operations:

```bash
# Basic connector tool
copilot agent tool add -a <agent-id> --toolType connector \
    --id "shared_asana:GetTask" --name "Get Asana Task"

# With input parameters
copilot agent tool add -a <agent-id> --toolType connector \
    --id "shared_office365:SendEmail" --name "Send Email" \
    --inputs '{"to": "string", "subject": "string", "body": "string"}'
```

**ID Format:** `connector_id:operation_id` (e.g., `shared_asana:GetTask`)

#### Tool Type: Prompt

Invoke AI Builder prompts:

```bash
copilot agent tool add -a <agent-id> --toolType prompt \
    --id <prompt-guid> --name "Summarize Text"

copilot agent tool add -a <agent-id> --toolType prompt \
    --id "12345678-1234-1234-1234-123456789abc" \
    --name "Analyze Sentiment" \
    --description "Analyzes the sentiment of customer feedback"
```

**ID Format:** Prompt GUID

#### Tool Type: Flow

Invoke Power Automate flows:

```bash
copilot agent tool add -a <agent-id> --toolType flow \
    --id <flow-guid> --name "Process Order"

copilot agent tool add -a <agent-id> --toolType flow \
    --id "12345678-1234-1234-1234-123456789abc" \
    --name "Create Support Ticket" \
    --inputs '{"title": "string", "priority": "string"}'
```

**ID Format:** Flow GUID (auto-prefixed with `/providers/Microsoft.Flow/flows/`)

#### Tool Type: HTTP

Make direct HTTP requests:

```bash
# GET request
copilot agent tool add -a <agent-id> --toolType http \
    --id "https://api.example.com/data" --name "Fetch Data"

# POST request with headers and body
copilot agent tool add -a <agent-id> --toolType http \
    --id "https://api.example.com/submit" \
    --name "Submit Data" \
    --method POST \
    --headers '{"Content-Type": "application/json"}' \
    --body '{"key": "value"}'
```

**ID Format:** Full URL

#### Tool Type: Agent (Connected Agent)

Connect another agent as a sub-agent:

```bash
copilot agent tool add -a <parent-id> --toolType agent \
    --id <target-agent-id> --name "Expert Reviewer"

# Without passing conversation history
copilot agent tool add -a <parent-id> --toolType agent \
    --id <target-agent-id> --name "Specialized Helper" --no-history
```

**ID Format:** Target agent GUID

**Requirements for connected agents:**
- Must be in the same environment
- Must be published
- Must have "Let other agents connect" enabled in settings

### Remove Agent Tool

```bash
copilot agent tool remove <component-id>           # Remove (with confirmation)
copilot agent tool remove <component-id> --force   # Remove without confirmation
```

---

## Knowledge Commands

### List Knowledge Sources

```bash
copilot agent knowledge list --bot <bot-id>
copilot agent knowledge list --bot <bot-id> --table
```

### Add File-Based Knowledge

```bash
# From inline content
copilot agent knowledge file add --bot <bot-id> --name "FAQ" --content "Q: What? A: Test."

# From file
copilot agent knowledge file add --bot <bot-id> --name "Guide" --file ./document.md
```

**Options:**
| Option | Description |
|--------|-------------|
| `-b, --bot` | Bot's unique identifier (required) |
| `-n, --name` | Display name for knowledge source (required) |
| `-c, --content` | Text content |
| `-f, --file` | Path to file containing content |
| `-d, --description` | Description (auto-generated if not provided) |

### Add Azure AI Search Knowledge (Experimental)

```bash
copilot agent knowledge azure-ai-search add --bot <bot-id> \
    --name "Product Docs" \
    --endpoint https://mysearch.search.windows.net \
    --index products-index \
    --api-key <api-key>
```

### Remove Knowledge Source

```bash
copilot agent knowledge remove --bot <bot-id> <component-id>
copilot agent knowledge remove --bot <bot-id> <component-id> --force
```

---

## Tool Commands (Unified Tool Management)

The `copilot tool` command provides unified management of all tool types available to Copilot Studio agents.

### List All Tools

```bash
copilot tool list                              # List all tools (JSON)
copilot tool list --table                      # Formatted table
copilot tool list --installed --table          # Only installed tools
copilot tool list --type prompt --table        # Filter by type
copilot tool list --type connector --table     # Filter by type
copilot tool list --filter "excel" --table     # Search by name
```

**Options:**
| Option | Description |
|--------|-------------|
| `-T, --type` | Filter by type: `prompt`, `connector`, `mcp` |
| `-i, --installed` | Show only tools installed in your environment |
| `-f, --filter` | Filter by name (case-insensitive) |
| `-t, --table` | Display as formatted table |

**Output Columns:**
- **Name** - Tool display name
- **Type** - Tool type (Prompt, Connector, Custom Connector, MCP)
- **Subtype** - For custom connectors: Rest Api, Custom Connector, Flow, Dataverse
- **Publisher** - Tool publisher
- **Installed** - Whether the tool is installed in your environment
- **Deps** - Number of dependent components (for installed tools)
- **ID** - Unique identifier

### Remove Tool

Remove a custom tool (prompt or REST API) from the environment.

```bash
copilot tool remove <tool-id>                      # Auto-detect type
copilot tool remove <tool-id> --type prompt        # Specify type
copilot tool remove <tool-id> --type restapi       # Specify type
copilot tool remove <tool-id> --force              # Skip confirmation
```

**Options:**
| Option | Description |
|--------|-------------|
| `-T, --type` | Tool type: `prompt`, `restapi` (auto-detected if not specified) |
| `-f, --force` | Skip confirmation prompt |

**Note:** When removing REST APIs, associated AIPlugin wrappers are automatically deleted.

---

### Prompt Commands (AI Builder)

Manage AI Builder prompts that can be used as agent tools.

```bash
# List prompts
copilot tool prompt list                        # All prompts (JSON)
copilot tool prompt list --table                # Formatted table
copilot tool prompt list --custom --table       # Custom prompts only
copilot tool prompt list --system --table       # System prompts only
copilot tool prompt list --filter "classify"    # Filter by name

# Get prompt details
copilot tool prompt get <prompt-id>             # Metadata
copilot tool prompt get <prompt-id> --text      # Prompt text and configuration

# Update prompt
copilot tool prompt update <prompt-id> --text "New prompt text..."
copilot tool prompt update <prompt-id> --file prompt.txt
copilot tool prompt update <prompt-id> --model gpt-4o
copilot tool prompt update <prompt-id> --file prompt.txt --no-publish
```

**Update Options:**
| Option | Description |
|--------|-------------|
| `-t, --text` | New prompt text (inline) |
| `-f, --file` | Path to file containing new prompt text |
| `-m, --model` | Model type (e.g., gpt-41-mini, gpt-4o, gpt-4o-mini) |
| `--no-publish` | Skip republishing (changes won't be live) |

---

### Connector Commands

List and view Power Platform connectors.

```bash
# List connectors
copilot tool connector list                     # All connectors (JSON)
copilot tool connector list --table             # Formatted table
copilot tool connector list --custom --table    # Custom connectors only
copilot tool connector list --managed --table   # Managed (Microsoft) connectors only
copilot tool connector list --filter "office365"

# Get connector details
copilot tool connector get <connector-id>
```

**List Options:**
| Option | Description |
|--------|-------------|
| `-c, --custom` | Show only custom connectors |
| `-m, --managed` | Show only managed (Microsoft) connectors |
| `-f, --filter` | Filter by name or publisher |
| `-t, --table` | Display as formatted table |

---

### Connector Connections Commands

Manage and test connections (authenticated instances) for connectors.

```bash
# List connections for a connector
copilot tool connector connections list --connector-id shared_office365
copilot tool connector connections list -c shared_commondataserviceforapps --table
copilot tool connector connections list -c shared_podio --connection-id abc123

# Test connection authentication
copilot tool connector connections auth-test --connector-id shared_office365
copilot tool connector connections auth-test -c shared_commondataserviceforapps --table
copilot tool connector connections auth-test -c shared_podio --connection-id abc123
copilot tool connector connections auth-test -c shared_office365 --test-api
```

**List Options:**
| Option | Description |
|--------|-------------|
| `-c, --connector-id` | **(Required)** The connector ID (e.g., shared_office365) |
| `--connection-id` | Filter to a specific connection ID |
| `-t, --table` | Display as formatted table |

**Auth-Test Options:**
| Option | Description |
|--------|-------------|
| `-c, --connector-id` | **(Required)** The connector ID |
| `--connection-id` | Test a specific connection ID (tests all if not provided) |
| `-t, --table` | Display as formatted table |
| `--test-api` | Also call the testConnection API endpoint (not all connectors support this) |

**Connection Statuses:**
| Status | Description |
|--------|-------------|
| `Connected` | Connection is authenticated and ready to use |
| `Error` | Connection has an authentication or configuration issue |
| `Unauthenticated` | Connection needs to be authenticated |

---

### REST API Commands

List and view REST API tools (custom connectors defined with OpenAPI specs).

```bash
# List REST APIs
copilot tool restapi list                       # All REST APIs (JSON)
copilot tool restapi list --table               # Formatted table
copilot tool restapi list --filter "podio"

# Get REST API details
copilot tool restapi get <connector-id>
```

---

### MCP Server Commands

List and view Model Context Protocol servers.

```bash
# List MCP servers
copilot tool mcp list                           # All MCP servers (JSON)
copilot tool mcp list --table                   # Formatted table

# Get MCP server details
copilot tool mcp get <server-id>
```

---

## Solution Commands

Manage Dataverse solutions and solution components.

### List Solutions

```bash
copilot solution list                           # Unmanaged solutions only (JSON)
copilot solution list --table                   # Formatted table
copilot solution list --all                     # Include managed solutions
```

**Options:**
| Option | Description |
|--------|-------------|
| `-t, --table` | Display as formatted table |
| `-a, --all` | Include managed solutions (default: unmanaged only) |

### Get Solution Details

```bash
copilot solution get <solution-id>
```

### Create Solution

```bash
copilot solution create --name "My Solution" --unique-name MySolution --publisher MyPublisher
copilot solution create -n "My Solution" -u MySolution -p MyPublisher -v 1.0.0.0
copilot solution create -n "My Solution" -u MySolution -p MyPublisher -d "Description"
```

**Options:**
| Option | Description |
|--------|-------------|
| `-n, --name` | Display name for the solution (required) |
| `-u, --unique-name` | Unique name (no spaces, required) |
| `-p, --publisher` | Publisher's unique name or GUID (required) |
| `-v, --version` | Version (default: 1.0.0.0) |
| `-d, --description` | Optional description |

### Delete Solution

```bash
copilot solution delete <solution-id>
```

### Add Agent to Solution

```bash
copilot solution add-agent --solution MySolution --bot <bot-id>
copilot solution add-agent -s MySolution -b <bot-id> --no-connection
copilot solution add-agent -s MySolution -b <bot-id> --no-required
```

**Options:**
| Option | Description |
|--------|-------------|
| `-s, --solution` | Solution's unique name (required) |
| `-b, --bot` | Bot's unique identifier (required) |
| `--no-connection` | Don't add bot's connection reference |
| `--no-required` | Don't add required dependent components |

### Remove Agent from Solution

```bash
copilot solution remove-agent --solution MySolution --bot <bot-id>
```

### Add/Remove Connection Reference

```bash
copilot solution add-connection --solution MySolution --connection <connection-ref-id>
copilot solution remove-connection --solution MySolution --connection <connection-ref-id>
```

---

### Publisher Commands

```bash
# List publishers
copilot solution publisher list
copilot solution publisher list --table

# Get publisher details
copilot solution publisher get <publisher-id>

# Create publisher
copilot solution publisher create \
    --name "My Publisher" \
    --unique-name MyPublisher \
    --prefix mypub \
    --option-prefix 10000
copilot solution publisher create -n "My Publisher" -u MyPublisher -x mypub -o 10000 -d "Description"

# Delete publisher
copilot solution publisher delete <publisher-id>
```

**Create Options:**
| Option | Description |
|--------|-------------|
| `-n, --name` | Display name (required) |
| `-u, --unique-name` | Unique name, no spaces (required) |
| `-x, --prefix` | Customization prefix, 2-8 lowercase letters (required) |
| `-o, --option-prefix` | Option value prefix, 10000-99999 (required) |
| `-d, --description` | Optional description |

---

### Connection Reference Commands

```bash
# List connection references
copilot solution connection list
copilot solution connection list --table
```

---

## Flow Commands

List and view Power Automate cloud flows.

### List Flows

```bash
copilot flow list                               # All flows (JSON)
copilot flow list --table                       # Formatted table
copilot flow list --category 5                  # Instant flows only
copilot flow list --category 0                  # Automated flows only
```

**Options:**
| Option | Description |
|--------|-------------|
| `-c, --category` | Filter by category: 0=Automated, 5=Instant, 6=Business Process |
| `-t, --table` | Display as formatted table |

**Flow Categories:**
- **0** - Automated (automated/scheduled flows)
- **5** - Instant (button/HTTP triggered flows) - best for agent tools
- **6** - Business Process flows

### Get Flow Details

```bash
copilot flow get <flow-id>
```

---

## Environment Commands

List and view Power Platform environments.

### List Environments

```bash
copilot environment list                        # All environments (JSON)
copilot environment list --table                # Formatted table
copilot environment list --filter "dev"         # Filter by name
```

**Options:**
| Option | Description |
|--------|-------------|
| `-f, --filter` | Filter by name (case-insensitive) |
| `-t, --table` | Display as formatted table |

### Get Environment Details

```bash
copilot environment get <environment-id>
copilot environment get Default-<tenant-id>
```

---

## Analytics Commands (Application Insights)

Query Application Insights telemetry for troubleshooting agent behavior.

### Get Analytics Configuration

```bash
copilot agent analytics get <bot-id>
```

### Enable/Disable Analytics

```bash
copilot agent analytics enable <bot-id>
copilot agent analytics disable <bot-id>
```

### Update Logging Options

```bash
copilot agent analytics update <bot-id>
```

### Query Telemetry

```bash
copilot agent analytics query <bot-id>                    # Last 24 hours
copilot agent analytics query <bot-id> --timespan 7d      # Last 7 days
copilot agent analytics query <bot-id> -t 1h              # Last hour
copilot agent analytics query <bot-id> --events           # Custom events only (faster)
copilot agent analytics query <bot-id> -t 1h -l 50        # Limit to 50 rows
copilot agent analytics query <bot-id> --json             # Raw JSON output
```

**Options:**
| Option | Description |
|--------|-------------|
| `-t, --timespan` | Time range (e.g., `1h`, `24h`, `7d`, `30d`) - default: `24h` |
| `-e, --events` | Query only customEvents table (faster) |
| `-l, --limit` | Maximum rows to display (default: 100) |
| `--json` | Output raw JSON response |

---

## Transcript Commands

View conversation transcripts for debugging and troubleshooting.

### List Transcripts

```bash
copilot agent transcript list                             # List recent transcripts
copilot agent transcript list --table                     # Formatted table
copilot agent transcript list --bot "Agent Name"          # Filter by bot name
copilot agent transcript list --bot <bot-id>              # Filter by bot ID
copilot agent transcript list --limit 10                  # Limit results
```

**Options:**
| Option | Description |
|--------|-------------|
| `-b, --bot` | Filter by bot name or ID |
| `-l, --limit` | Maximum transcripts to return (default: 20) |
| `-t, --table` | Display as formatted table |

### Get Transcript Content

```bash
copilot agent transcript get <transcript-id>
```

---

## Connection Commands

Manage Power Platform connections for Copilot Studio integrations.

### List Connections

```bash
copilot agent connection list --environment Default-<tenant-id>
copilot agent connection list --environment Default-<tenant-id> --table
```

### Create Connection

```bash
copilot agent connection create \
    --name "My Search Connection" \
    --endpoint https://mysearch.search.windows.net \
    --api-key <api-key> \
    --environment Default-<tenant-id>
```

### Delete Connection

```bash
copilot agent connection delete <connection-id> --environment Default-<tenant-id>
```

---

## Topic YAML Schema Reference

Topics use AdaptiveDialog YAML format. See [topic-yaml-schema.md](./topic-yaml-schema.md) for the complete schema reference including:
- Trigger types (OnRecognizedIntent, OnActivity, OnConversationStart, etc.)
- Action nodes (SendMessage, Question, ConditionGroup, BeginDialog, etc.)
- Variables and entities
- Conditions and Power Fx expressions

### Simple Topic Example

```yaml
kind: AdaptiveDialog
beginDialog:
  kind: OnRecognizedIntent
  id: main
  intent:
    displayName: Greeting
    triggerQueries:
      - hello
      - hi
      - hey there

  actions:
    - kind: SendMessage
      id: sendMessage_greeting
      message: Hello! How can I help you today?
```

### Topic with Condition Branches

```yaml
kind: AdaptiveDialog
beginDialog:
  kind: OnRecognizedIntent
  id: main
  intent:
    displayName: Product Help
    triggerQueries:
      - help with product
      - product question

  actions:
    - kind: ConditionGroup
      id: conditionGroup_product
      conditions:
        - id: condition_moveit
          condition: =Topic.product = "MOVEit"
          displayName: MOVEit
          actions:
            - kind: BeginDialog
              id: beginDialog_moveit
              dialog: cr83c_myAgent.InvokeConnectedAgentTaskAction.MOVEitExpert

        - id: condition_sitefinity
          condition: =Topic.product = "Sitefinity"
          displayName: Sitefinity
          actions:
            - kind: BeginDialog
              id: beginDialog_sitefinity
              dialog: cr83c_myAgent.InvokeConnectedAgentTaskAction.SitefinityExpert

      elseActions:
        - kind: SendMessage
          id: sendMessage_unknown
          message: I couldn't determine which product you need help with.
```

---

## Troubleshooting

### Common Issues

**Authentication Errors:**
```bash
# Ensure you're logged in to Azure CLI
az login

# Verify Dataverse URL is set
echo $DATAVERSE_URL
```

**Agent Not Responding:**
1. Check if agent is published: `copilot agent get <bot-id>`
2. Check analytics for errors: `copilot agent analytics query <bot-id> -t 1h`
3. Review recent transcripts: `copilot agent transcript list --bot <bot-id>`

**Topic Not Triggering:**
1. Ensure topic is enabled: `copilot agent topic list --agentId <bot-id> --table`
2. Check trigger phrases match user input
3. Verify topic YAML syntax: `copilot agent topic get <topic-id> --yaml`

**Connected Agent Not Working:**
1. Verify target agent is published
2. Check "Let other agents connect" is enabled in target agent settings
3. Review tool configuration: `copilot agent tool list --agentId <agent-id> --table`

### Debug Workflow

1. **Check agent status:**
   ```bash
   copilot agent get <bot-id>
   ```

2. **Query recent telemetry:**
   ```bash
   copilot agent analytics query <bot-id> -t 1h --events
   ```

3. **Review conversation transcripts:**
   ```bash
   copilot agent transcript list --bot <bot-id> --limit 5 --table
   copilot agent transcript get <transcript-id>
   ```

4. **Verify topic configuration:**
   ```bash
   copilot agent topic list --agentId <bot-id> --table
   copilot agent topic get <topic-id> --yaml
   ```

5. **Check tool dependencies:**
   ```bash
   copilot tool list --installed --table
   ```
