"""Cursor and navigation tools for MuseScore MCP."""

from ..client import MuseScoreClient


def setup_navigation_tools(mcp, client: MuseScoreClient):
    """Setup cursor and navigation tools."""
    
    @mcp.tool()
    async def get_cursor_info():
        """Get information about the current cursor position."""
        return await client.send_command("getCursorInfo")

    @mcp.tool()
    async def go_to_measure(measure: int):
        """Navigate to a specific measure."""
        return await client.send_command("goToMeasure", {"measure": measure})

    @mcp.tool()
    async def go_to_final_measure():
        """Navigate to the final measure of the score."""
        return await client.send_command("goToFinalMeasure")

    @mcp.tool()
    async def go_to_beginning_of_score():
        """Navigate to the beginning of the score."""
        return await client.send_command("goToBeginningOfScore")

    @mcp.tool()
    async def next_element():
        """Move cursor to the next element."""
        return await client.send_command("nextElement")

    @mcp.tool()
    async def prev_element():
        """Move cursor to the previous element."""
        return await client.send_command("prevElement")

    @mcp.tool()
    async def next_staff():
        """Move cursor to the next staff."""
        return await client.send_command("nextStaff")

    @mcp.tool()
    async def prev_staff():
        """Move cursor to the previous staff."""
        return await client.send_command("prevStaff")

    @mcp.tool()
    async def select_current_measure():
        """Select the current measure."""
        return await client.send_command("selectCurrentMeasure")