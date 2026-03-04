"""Staff and instrument tools for MuseScore MCP."""

from ..client import MuseScoreClient


def setup_staff_instruments_tools(mcp, client: MuseScoreClient):
    """Setup staff and instrument tools."""
    
    @mcp.tool()
    async def add_instrument(instrument_id: str):
        """Add a new staff/instrument to the score.
        
        Args:
            instrument_id: ID of the instrument to add
        """
        return await client.send_command("addInstrument", {
            "instrumentId": instrument_id
        })

    @mcp.tool()
    async def set_staff_mute(staff: int, mute: bool):
        """Mute or unmute a staff.
        
        Args:
            staff: Staff number (0-based)
            mute: True to mute, False to unmute
        """
        return await client.send_command("setStaffMute", {
            "staff": staff,
            "mute": mute
        })

    @mcp.tool()
    async def set_instrument_sound(staff: int, instrument_id: str):
        """Change the sound of an instrument on a staff.
        
        Args:
            staff: Staff number (0-based)
            instrument_id: ID of the new instrument sound
        """
        return await client.send_command("setInstrumentSound", {
            "staff": staff,
            "instrumentId": instrument_id
        })