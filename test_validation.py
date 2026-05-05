#!/usr/bin/env python3
"""Test the exact validation that's failing."""

import sys
import os
import asyncio
from pathlib import Path

# Add src to path
sys.path.insert(0, str(Path(__file__).parent / "src"))

# Set required environment variable
os.environ["MICROSOFT_MCP_CLIENT_ID"] = os.environ.get("MICROSOFT_MCP_CLIENT_ID", "test-id")

from microsoft_mcp.tools import mcp
import mcp.types as types

async def test_validation():
    """Test the exact validation that might be failing."""
    print("=== Testing Exact Validation ===")
    
    # Test the exact flow that happens in session
    print("\n1. Creating JSONRPCRequest (like client would send)...")
    
    # Test case 1: No params
    try:
        jsonrpc_request = types.JSONRPCRequest(
            jsonrpc="2.0",
            id="test1",
            method="ping"
            # No params field
        )
        print(f"   JSONRPCRequest (no params): {jsonrpc_request}")
        
        # Convert to dict like in session
        dumped = jsonrpc_request.model_dump(by_alias=True, mode="json", exclude_none=True)
        print(f"   Dumped dict: {dumped}")
        
        # Try to validate as ClientRequest
        client_request = types.ClientRequest.model_validate(dumped)
        print(f"   ClientRequest: {client_request}")
        print(f"   Root type: {type(client_request.root)}")
        
        # This is what happens in session - validate as receive_request_type
        validated = mcp._mcp_server.request_handlers  # This should work
        print(f"   Session receive_request_type validation would work")
        
    except Exception as e:
        print(f"   Error with no params: {e}")
        import traceback
        traceback.print_exc()
    
    # Test case 2: Empty params
    print("\n2. Testing with empty params...")
    try:
        jsonrpc_request = types.JSONRPCRequest(
            jsonrpc="2.0",
            id="test2",
            method="ping",
            params={}
        )
        print(f"   JSONRPCRequest (empty params): {jsonrpc_request}")
        
        # Convert to dict like in session
        dumped = jsonrpc_request.model_dump(by_alias=True, mode="json", exclude_none=True)
        print(f"   Dumped dict: {dumped}")
        
        # Try to validate as ClientRequest
        client_request = types.ClientRequest.model_validate(dumped)
        print(f"   ClientRequest: {client_request}")
        print(f"   Root type: {type(client_request.root)}")
        
    except Exception as e:
        print(f"   Error with empty params: {e}")
        import traceback
        traceback.print_exc()

    # Test case 3: Direct handler call simulation
    print("\n3. Testing direct handler call...")
    try:
        # Create the exact type of request that should be handled
        ping_request = types.PingRequest(method="ping")
        print(f"   PingRequest: {ping_request}")
        
        # Call our registered handler
        handler = mcp._mcp_server.request_handlers[types.PingRequest]
        print(f"   Handler: {handler}")
        
        result = await handler(ping_request)
        print(f"   Handler result: {result}")
        
    except Exception as e:
        print(f"   Error in direct handler call: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    asyncio.run(test_validation())