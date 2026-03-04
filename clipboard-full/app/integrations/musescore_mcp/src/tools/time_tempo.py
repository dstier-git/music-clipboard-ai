"""Time signature and tempo tools for MuseScore MCP."""

from ..client import MuseScoreClient


def setup_time_tempo_tools(mcp, client: MuseScoreClient):
    """Setup time signature and tempo tools."""
    
    @mcp.tool()
    async def set_time_signature(numerator: int = 4, denominator: int = 4):
        """Set the time signature.
        
        Args:
            numerator: Top number of time signature (beats per measure)
            denominator: Bottom number of time signature (note value that gets the beat)
        """
        return await client.send_command("setTimeSignature", {
            "numerator": numerator,
            "denominator": denominator
        })