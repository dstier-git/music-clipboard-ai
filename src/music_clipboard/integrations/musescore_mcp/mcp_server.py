import logging
import sys
from pathlib import Path
from mcp.server.fastmcp import FastMCP

# Ensure local imports work when launched via absolute script path.
SCRIPT_DIR = Path(__file__).resolve().parent
if str(SCRIPT_DIR) not in sys.path:
    sys.path.insert(0, str(SCRIPT_DIR))

from src.client import MuseScoreClient
from src.tools import (
    setup_connection_tools,
    setup_navigation_tools,
    setup_notes_measures_tools,
    setup_staff_instruments_tools,
    setup_time_tempo_tools,
    setup_sequence_tools,
    setup_shell_tools,
)

# Set up logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[logging.StreamHandler(sys.stderr)]
)
logger = logging.getLogger("MuseScoreMCP")

# Create the MCP app and client
mcp = FastMCP("MuseScore Assistant")
client = MuseScoreClient()

# Setup all tool categories
setup_connection_tools(mcp, client)
setup_navigation_tools(mcp, client)
setup_notes_measures_tools(mcp, client)
setup_staff_instruments_tools(mcp, client)
setup_time_tempo_tools(mcp, client)
setup_sequence_tools(mcp, client)
setup_shell_tools(mcp)

# Main entry point
if __name__ == "__main__":
    sys.stderr.write("MuseScore MCP Server starting up...\n")
    sys.stderr.flush()
    logger.info("MuseScore MCP Server is running")
    mcp.run()
