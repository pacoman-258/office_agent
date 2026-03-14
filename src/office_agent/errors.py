class OfficeAgentError(Exception):
    """Base exception for the application."""


class ConfigError(OfficeAgentError):
    """Raised when runtime configuration is invalid."""


class ProviderError(OfficeAgentError):
    """Raised when an LLM provider request fails."""


class SpecGenerationError(OfficeAgentError):
    """Raised when a provider response cannot be converted into a spec."""


class RenderError(OfficeAgentError):
    """Raised when rendering the PowerPoint fails."""


class OfficeAutomationError(OfficeAgentError):
    """Raised when PowerPoint automation or visual finalization fails."""
