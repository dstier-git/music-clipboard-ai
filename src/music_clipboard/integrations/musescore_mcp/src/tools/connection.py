"""Connection and utility tools for MuseScore MCP."""

from ..client import MuseScoreClient


def setup_connection_tools(mcp, client: MuseScoreClient):
    """Setup connection and utility tools."""
    
    @mcp.tool()
    async def connect_to_musescore():
        """Connect to the MuseScore WebSocket API."""
        result = await client.connect()
        return {"success": result}

    @mcp.tool()
    async def ping_musescore():
        """Ping the MuseScore WebSocket API to check connection."""
        return await client.send_command("ping")

    @mcp.tool()
    async def get_score():
        """Get information about the current score."""
        return await client.send_command("getScore")

    @mcp.tool()
    async def export_score(path: str, format: str = ""):
        """Export the current score to disk.

        Args:
            path: Absolute output path (for example /tmp/output.mid).
            format: Optional export format override (for example mid, pdf, musicxml).
        """
        params = {"path": path}
        if format and format.strip():
            params["format"] = format.strip().lower()
        return await client.send_command("exportScore", params)
