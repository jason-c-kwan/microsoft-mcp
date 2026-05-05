#!/usr/bin/env python3
"""Minimal test to replicate client behavior."""

import sys
import os
import json
import asyncio
from pathlib import Path

# Add src to path
sys.path.insert(0, str(Path(__file__).parent / "src"))

# Set required environment variable
os.environ["MICROSOFT_MCP_CLIENT_ID"] = os.environ.get("MICROSOFT_MCP_CLIENT_ID", "test-id")

def test_json_parsing():
    """Test parsing the exact JSON that the client sends."""
    print("=== Testing JSON Parsing ===")
    
    # Test case 1: No params
    json1 = {"jsonrpc": "2.0", "method": "ping", "id": "2"}
    print(f"1. JSON (no params): {json1}")
    
    try:
        import mcp.types as types
        jsonrpc_request = types.JSONRPCRequest.model_validate(json1)
        print(f"   Parsed as JSONRPCRequest: {jsonrpc_request}")
        
        # Convert to dict and back (like session does)
        dumped = jsonrpc_request.model_dump(by_alias=True, mode="json", exclude_none=True)
        print(f"   Dumped: {dumped}")
        
        # Validate as ClientRequest
        client_request = types.ClientRequest.model_validate(dumped)
        print(f"   ClientRequest: {client_request}")
        print(f"   Root type: {type(client_request.root)}")
        
    except Exception as e:
        print(f"   Error: {e}")
        import traceback
        traceback.print_exc()
    
    # Test case 2: Empty params
    json2 = {"jsonrpc": "2.0", "method": "ping", "id": "3", "params": {}}
    print(f"\n2. JSON (empty params): {json2}")
    
    try:
        import mcp.types as types
        jsonrpc_request = types.JSONRPCRequest.model_validate(json2)
        print(f"   Parsed as JSONRPCRequest: {jsonrpc_request}")
        
        # Convert to dict and back (like session does)
        dumped = jsonrpc_request.model_dump(by_alias=True, mode="json", exclude_none=True)
        print(f"   Dumped: {dumped}")
        
        # Validate as ClientRequest
        client_request = types.ClientRequest.model_validate(dumped)
        print(f"   ClientRequest: {client_request}")
        print(f"   Root type: {type(client_request.root)}")
        
    except Exception as e:
        print(f"   Error: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    test_json_parsing()