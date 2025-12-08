# Copilot CLI Guide

Complete reference for the `copilot` command-line interface for managing Microsoft Copilot Studio agents via the Dataverse API.

## Overview

The Copilot CLI provides access to:
- **Agents** - Create, update, delete, publish, and test agents
- **Topics** - Manage conversation flows (list, create, update, delete, enable/disable)
- **Tools** - Connect sub-agents as tools for orchestration
- **Knowledge** - Add file-based and Azure AI Search knowledge sources
- **Analytics** - Query Application Insights telemetry for troubleshooting
- **Transcripts** - View conversation history for debugging
- **Connections** - Manage Power Platform connections

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

**Environment Variables:**
- `DIRECTLINE_SECRET` - Direct Line secret
- `ENTRA_CLIENT_ID` - Entra ID client ID
- `ENTRA_TENANT_ID` - Entra ID tenant ID
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

## Tool Commands (Connected Agents)

Tools allow an agent to invoke other agents as sub-agents during orchestration.

### List Tools

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

### Add Connected Agent Tool

```bash
copilot agent tool add --agentId <parent-id> --target <child-agent-id>
copilot agent tool add -a <parent-id> -t <child-id> --name "Expert Reviewer"
copilot agent tool add -a <parent-id> -t <child-id> --description "Handles X tasks" --no-history
```

**Options:**
| Option | Description |
|--------|-------------|
| `-a, --agentId` | Parent agent's ID (required) |
| `-t, --target` | Target agent's ID to connect (required) |
| `-n, --name` | Display name for the tool |
| `-d, --description` | Description for AI orchestration |
| `--no-history` | Don't pass conversation history to connected agent |

**Requirements for target agent:**
- Must be in the same environment
- Must be published
- Must have "Let other agents connect" enabled in settings

### Remove Tool

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

## Prompt Commands (AI Builder)

Manage AI Builder prompts that can be used as agent tools for classification, extraction, and content generation.

### List Prompts

```bash
copilot prompt list                     # List all prompts (JSON)
copilot prompt list --table             # List as formatted table
copilot prompt list --custom            # Show only custom prompts
copilot prompt list --system            # Show only system prompts
copilot prompt list --filter "classify" # Filter by name
```

### Get Prompt Details

```bash
copilot prompt get <prompt-id>          # Get prompt metadata
copilot prompt get <prompt-id> --text   # Get prompt text and configuration
```

### Update Prompt

Update the prompt text or model type for an AI Builder prompt. The prompt is automatically republished after updating.

```bash
copilot prompt update <prompt-id> --text "New prompt text..."
copilot prompt update <prompt-id> --file prompt.txt
copilot prompt update <prompt-id> --model gpt-4o
copilot prompt update <prompt-id> --file prompt.txt --model gpt-4o
copilot prompt update <prompt-id> --file prompt.txt --no-publish
```

**Options:**
| Option | Description |
|--------|-------------|
| `-t, --text` | New prompt text (inline) |
| `-f, --file` | Path to file containing new prompt text |
| `-m, --model` | Model type (e.g., gpt-41-mini, gpt-4o, gpt-4o-mini) |
| `--no-publish` | Skip republishing (changes won't be live) |

**Note:** Updates are automatically published. Use `--no-publish` to update without publishing.

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
