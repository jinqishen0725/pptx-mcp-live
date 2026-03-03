"""Error types for PowerPoint MCP Live."""


class ToolError(Exception):
    """Raised when a tool operation fails."""

    def __init__(self, message: str):
        self.message = message
        super().__init__(message)
