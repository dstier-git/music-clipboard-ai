"""Sequence processing tools for MuseScore MCP."""

from ..client import MuseScoreClient
from ..types import ActionSequence


def setup_sequence_tools(mcp, client: MuseScoreClient):
    """Setup sequence processing tools."""
    
    @mcp.tool()
    async def processSequence(sequence: ActionSequence):
        """Process a sequence of commands."""
        return await client.send_command("processSequence", {"sequence": sequence})