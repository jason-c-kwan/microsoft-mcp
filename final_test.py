#!/usr/bin/env python3
"""Final test to replicate the exact issue."""

import sys
import os
import asyncio
from pathlib import Path

# Add src to path
sys.path.insert(0, str(Path(__file__).parent / "src"))

# Set required environment variable
os.environ["MICROSOFT_MCP_CLIENT_ID"] = os.environ.get("MICROSOFT_MCP_CLIENT_ID", "test-id")

async def test_exact_scenario():
    """Test the exact scenario that's failing."""
    print("=== Testing Exact Failure Scenario ===")
    
    # Import after setting environment
    from microsoft_mcp.tools import mcp
    import mcp.types as types
    
    print("1. Checking initial handler registration...")
    has_ping = types.PingRequest in mcp._mcp_server.request_handlers
    print(f"   Ping handler registered initially: {has_ping}")
    if has_ping:
        print(f"   Handler function: {mcp._mcp_server.request_handlers[types.PingRequest]}")
    
    print("\n2. Testing our enhanced registration...")
    # This should run our _create_enhanced_handlers function
    # Let's check what handlers we have now
    has_ping = types.PingRequest in mcp._mcp_server.request_handlers
    print(f"   Ping handler registered after enhancement: {has_ping}")
    if has_ping:
        print(f"   Handler function: {mcp._mcp_server.request_handlers[types.PingRequest]}")
    
    print("\n3. Testing direct handler call (should work)...")
    try:
        if has_ping:
            ping_request = types.PingRequest(method="ping")
            result = await mcp._mcp_server.request_handlers[types.PingRequest](ping_request)
            print(f"   Direct call success: {result}")
        else:
            print("   No ping handler to test")
    except Exception as e:
        print(f"   Direct call error: {e}")
        import traceback
        traceback.print_exc()
    
    print("\n4. Checking all registered handlers...")
    for req_type, handler in mcp._mcp_server.request_handlers.items():
        print(f"   {req_type.__name__}: {handler}")

if __name__ == "__main__":
    asyncio.run(test_exact_scenario())