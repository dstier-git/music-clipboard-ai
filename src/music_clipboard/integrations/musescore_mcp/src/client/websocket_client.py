"""WebSocket client for communicating with MuseScore."""

import websockets
import json
import logging
from typing import Dict, Any, Optional

logger = logging.getLogger("MuseScoreMCP.Client")


class MuseScoreClient:
    """Client to communicate with MuseScore WebSocket API."""
    
    def __init__(self, host: str = "localhost", port: int = 8765):
        self.uri = f"ws://{host}:{port}"
        self.websocket = None
    
    async def connect(self):
        """Connect to the MuseScore WebSocket API."""
        try:
            self.websocket = await websockets.connect(self.uri)
            logger.info(f"Connected to MuseScore API at {self.uri}")
            return True
        except Exception as e:
            logger.error(f"Failed to connect to MuseScore API: {str(e)}")
            return False
    
    async def send_command(self, action: str, params: Optional[Dict[str, Any]] = None) -> Dict[str, Any]:
        """Send a command to MuseScore and wait for response."""
        if not self.websocket:
            connected = await self.connect()
            if not connected:
                return {"error": "Not connected to MuseScore"}
        
        if params is None:
            params = {}
        
        command = {"action": action, "params": params}
        
        try:
            logger.info(f"Sending command: {json.dumps(command)}")
            await self.websocket.send(json.dumps(command))
            response = await self.websocket.recv()
            logger.info(f"Received response: {response}")
            return json.loads(response)
        except Exception as e:
            logger.error(f"Error sending command: {str(e)}")
            return {"error": str(e)}
    
    async def close(self):
        """Close the WebSocket connection."""
        if self.websocket:
            await self.websocket.close()
            self.websocket = None
            logger.info("Disconnected from MuseScore API")