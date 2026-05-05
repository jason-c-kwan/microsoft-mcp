#!/usr/bin/env python3
"""Debug script to test MCP server behavior."""

import sys
import os
import json
import asyncio
from pathlib import Path

# Add src to path
sys.path.insert(0, str(Path(__file__).parent / "src"))

# Set required environment variable
os.environ["MICROSOFT_MCP_CLIENT_ID"] = os.environ.get("MICROSOFT_MCP_CLIENT_ID", "test-id")

from microsoft_mcp.tools import mcp
import mcp.types as types
from mcp.server.stdio import stdio_server
from mcp.server.lowlevel.server import Server as MCPServer
from mcp.server.models import InitializationOptions
from mcp.server.lowlevel.server import NotificationOptions

async def debug_server():
    """Debug the server setup and handlers."""
    print("=== Debug Server Setup ===")
    
    # Check what handlers are registered
    print("Registered handlers:")
    for handler_type, handler_func in mcp._mcp_server.request_handlers.items():
        print(f"  {handler_type}: {handler_func}")
    
    # Test handler functionality
    print("\n=== Testing Handler Direct Calls ===")
    
    # Test ping handler
    try:
        ping_request = types.PingRequest(method="ping")
        result = await mcp._mcp_server.request_handlers[types.PingRequest](ping_request)
        print(f"Ping handler works: {result}")
    except Exception as e:
        print(f"Ping handler error: {e}")
        import traceback
        traceback.print_exc()
    
    # Test list tools handler
    try:
        list_tools_request = types.ListToolsRequest(method="tools/list")
        result = await mcp._mcp_server.request_handlers[types.ListToolsRequest](list_tools_request)
        print(f"List tools handler works: {type(result)}")
        if hasattr(result, 'result') and hasattr(result.result, 'tools'):
            print(f"Number of tools: {len(result.result.tools)}")
    except Exception as e:
        print(f"List tools handler error: {e}")
        import traceback
        traceback.print_exc()
    
    print("\n=== Server Configuration ===")
    print(f"Server name: {mcp._mcp_server.name}")
    print(f"Has lifespan: {getattr(mcp, '_has_lifespan', 'Unknown')}")

if __name__ == "__main__":
    asyncio.run(debug_server())