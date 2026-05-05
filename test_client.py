#!/usr/bin/env python3
"""Test script to send requests to the MCP server and see what happens."""

import sys
import os
import json
import asyncio
from pathlib import Path

# Add src to path
sys.path.insert(0, str(Path(__file__).parent / "src"))

# Set required environment variable
os.environ["MICROSOFT_MCP_CLIENT_ID"] = os.environ.get("MICROSOFT_MCP_CLIENT_ID", "test-id")

import mcp.types as types
from mcp import ClientSession, StdioServerParameters
from mcp.client.stdio import stdio_client

async def test_server():
    """Test the server with actual requests."""
    print("=== Testing Server with Actual Requests ===")
    
    # Set up server parameters
    server_params = StdioServerParameters(
        command="uv",
        args=["run", "microsoft-mcp"],
        env={"MICROSOFT_MCP_CLIENT_ID": "test-id"}
    )
    
    try:
        async with stdio_client(server_params) as (read, write):
            async with ClientSession(read, write) as session:
                # Test initialize
                print("1. Testing initialize...")
                try:
                    init_result = await session.initialize()
                    print(f"   Initialize success: {init_result is not None}")
                except Exception as e:
                    print(f"   Initialize error: {e}")
                    return
                
                # Test ping with no params
                print("2. Testing ping (no params)...")
                try:
                    ping_result = await session._send_request(
                        types.ClientRequest(types.PingRequest(method="ping")),
                        types.ServerResult
                    )
                    print(f"   Ping success: {ping_result}")
                except Exception as e:
                    print(f"   Ping error: {e}")
                
                # Test ping with empty params
                print("3. Testing ping (empty params)...")
                try:
                    ping_result = await session._send_request(
                        types.ClientRequest(types.PingRequest(method="ping", params={})),
                        types.ServerResult
                    )
                    print(f"   Ping with empty params success: {ping_result}")
                except Exception as e:
                    print(f"   Ping with empty params error: {e}")
                
                # Test tools/list
                print("4. Testing tools/list...")
                try:
                    tools_result = await session._send_request(
                        types.ClientRequest(types.ListToolsRequest(method="tools/list")),
                        types.ServerResult
                    )
                    print(f"   Tools/list success: {tools_result is not None}")
                except Exception as e:
                    print(f"   Tools/list error: {e}")
                    
    except Exception as e:
        print(f"Connection error: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    asyncio.run(test_server())