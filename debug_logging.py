#!/usr/bin/env python3
"""Enhanced debug script with logging."""

import sys
import os
import logging
import asyncio
from pathlib import Path

# Enable detailed logging
logging.basicConfig(level=logging.DEBUG)
logger = logging.getLogger(__name__)

# Add src to path
sys.path.insert(0, str(Path(__file__).parent / "src"))

# Set required environment variable
os.environ["MICROSOFT_MCP_CLIENT_ID"] = os.environ.get("MICROSOFT_MCP_CLIENT_ID", "test-id")

from microsoft_mcp.tools import mcp
import mcp.types as types

# Enable FastMCP logging
import fastmcp
fastmcp.settings.log_level = "DEBUG"

async def debug_with_logging():
    """Debug with detailed logging."""
    print("=== Debug with Detailed Logging ===")
    
    # Check server configuration
    print(f"Server name: {mcp._mcp_server.name}")
    print(f"Server instructions: {mcp._mcp_server.instructions}")
    
    # Check handlers
    print("\nRegistered handlers:")
    for handler_type, handler_func in mcp._mcp_server.request_handlers.items():
        print(f"  {handler_type.__name__}: {handler_func}")
    
    # Test creating requests that might fail
    print("\n=== Testing Request Creation ===")
    
    # Test various ping request formats
    test_cases = [
        ("Ping with no params", {"method": "ping"}),
        ("Ping with null params", {"method": "ping", "params": None}),
        ("Ping with empty params", {"method": "ping", "params": {}}),
    ]
    
    for name, data in test_cases:
        try:
            # Try to create JSONRPCRequest first
            jsonrpc_request = types.JSONRPCRequest(
                jsonrpc="2.0",
                id="test",
                method=data["method"],
                params=data.get("params")
            )
            print(f"  {name}: JSONRPCRequest created successfully")
            
            # Try to convert to ClientRequest
            client_request = types.ClientRequest.model_validate(
                jsonrpc_request.model_dump(by_alias=True, mode="json", exclude_none=True)
            )
            print(f"  {name}: ClientRequest conversion successful")
            
            # Try to access the root request
            root_request = client_request.root
            print(f"  {name}: Root request type: {type(root_request)}")
            
        except Exception as e:
            print(f"  {name}: Error - {e}")
            import traceback
            traceback.print_exc()

if __name__ == "__main__":
    asyncio.run(debug_with_logging())