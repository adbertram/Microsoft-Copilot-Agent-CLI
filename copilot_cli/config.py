"""Configuration management for CopilotAgent CLI."""
import os
from pathlib import Path
from typing import Optional
from dotenv import load_dotenv


class Config:
    """Configuration manager for CopilotAgent CLI authentication and settings."""

    def __init__(self):
        """Initialize configuration by loading from .env file."""
        # Get the directory where this config.py file is located
        config_dir = Path(__file__).parent.parent  # Go up from copilotagent_cli/ to CopilotAgent-CLI/
        cli_env_path = config_dir / ".env"

        self.env_file_path = cli_env_path

        if cli_env_path.exists():
            load_dotenv(cli_env_path, override=True)
        else:
            # Create default .env if it doesn't exist
            cli_env_path.parent.mkdir(parents=True, exist_ok=True)
            cli_env_path.touch()

    @property
    def dataverse_url(self) -> Optional[str]:
        """Get Dataverse environment URL (e.g., https://org1cb52429.crm.dynamics.com)."""
        return os.getenv("DATAVERSE_URL")

    @property
    def environment_id(self) -> Optional[str]:
        """Get Power Platform environment ID (e.g., Default-<tenant-id>)."""
        return os.getenv("DATAVERSE_ENVIRONMENT_ID") or os.getenv("POWERPLATFORM_ENVIRONMENT_ID")

    @property
    def tenant_id(self) -> Optional[str]:
        """Get Azure AD tenant ID."""
        return os.getenv("AZURE_TENANT_ID")

    @property
    def client_id(self) -> Optional[str]:
        """Get Azure AD application (client) ID."""
        return os.getenv("AZURE_CLIENT_ID")

    @property
    def client_secret(self) -> Optional[str]:
        """Get Azure AD client secret."""
        return os.getenv("AZURE_CLIENT_SECRET")

    def has_service_principal_auth(self) -> bool:
        """Check if service principal authentication credentials are available."""
        return bool(
            self.dataverse_url
            and self.tenant_id
            and self.client_id
            and self.client_secret
        )

    def has_cli_auth(self) -> bool:
        """Check if we can use Azure CLI authentication."""
        return bool(self.dataverse_url)

    def get_missing_credentials(self) -> list[str]:
        """Get list of missing credentials for any authentication method."""
        missing = []

        # Dataverse URL is always required
        if not self.dataverse_url:
            missing.append("DATAVERSE_URL")
            return missing

        # Check if we have complete service principal auth
        if self.has_service_principal_auth():
            return []  # No missing credentials

        # For CLI auth, only DATAVERSE_URL is required
        # But note that user needs to be logged in via `az login`
        return []

    def get_auth_method(self) -> str:
        """Determine which authentication method to use."""
        if self.has_service_principal_auth():
            return "service_principal"
        elif self.has_cli_auth():
            return "azure_cli"
        else:
            return "none"


# Global config instance
_config: Optional[Config] = None


def get_config() -> Config:
    """Get or create the global config instance."""
    global _config
    if _config is None:
        _config = Config()
    return _config
