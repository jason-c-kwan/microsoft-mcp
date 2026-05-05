#!/usr/bin/env python3

import asyncio
import json
import subprocess
import os
from pathlib import Path

async def test_proper_initialization_sequence():
    # Start the Microsoft Graph MCP server
    env = dict(os.environ)
    env["MICROSOFT_MCP_CLIENT_ID"] = "test-client-id"  # Add required env var
    
    process = subprocess.Popen(
        ["uv", "run", "python", "-m", "microsoft_mcp.server"],
        env=env,
        cwd="/Users/jkwan2/git_repos/microsoft-mcp",
        stdout=subprocess.PIPE,
        stdin=subprocess.PIPE,
        stderr=subprocess.PIPE,
        text=False
    )
    
    try:
        print("=== Testing Proper MCP Initialization Sequence ===")
        
        # Give the server a moment to start
        await asyncio.sleep(1)
        
        # Check if server started successfully
        if process.poll() is not None:
            print("   Server failed to start with return code:", process.returncode)
            stderr_output = process.stderr.read()
            if stderr_output:
                print("   Server stderr:", stderr_output.decode())
            return
        
        # Step 1: Send initialize request
        print("1. Sending initialize request...")
        init_request = {
            "jsonrpc": "2.0",
            "method": "initialize",
            "id": "1",
            "params": {
                "protocolVersion": "2024-11-05",
                "capabilities": {
                    "tools": {}
                },
                "clientInfo": {
                    "name": "test_client",
                    "version": "1.0.0"
                }
            }
        }
        
        init_json = json.dumps(init_request) + "\n"
        print("   Sending: " + init_json.strip())
        
        process.stdin.write(init_json.encode())
        await asyncio.get_event_loop().run_in_executor(None, process.stdin.flush)
        
        # Read initialize response
        try:
            line = await asyncio.wait_for(
                asyncio.get_event_loop().run_in_executor(None, process.stdout.readline),
                timeout=10.0
            )
            
            if line:
                response = json.loads(line.decode().strip())
                print("   Initialize Response: " + str(response))
            else:
                print("   No initialize response received")
                return
        except Exception as e:
            print("   Error reading initialize response: " + str(e))
            return
        
        # Step 2: Send initialized notification (this is required!)
        print("\n2. Sending initialized notification...")
        initialized_notification = {
            "jsonrpc": "2.0",
            "method": "notifications/initialized"
        }
        
        initialized_json = json.dumps(initialized_notification) + "\n"
        print("   Sending: " + initialized_json.strip())
        
        process.stdin.write(initialized_json.encode())
        await asyncio.get_event_loop().run_in_executor(None, process.stdin.flush)
        
        # Small delay to let the server process the notification
        await asyncio.sleep(0.1)
        
        # Step 3: Send ping request (no params field)
        print("\n3. Sending ping request (no params field)...")
        ping_request1 = {
            "jsonrpc": "2.0",
            "method": "ping",
            "id": "2"
        }
        
        ping_json1 = json.dumps(ping_request1) + "\n"
        print("   Sending: " + ping_json1.strip())
        
        process.stdin.write(ping_json1.encode())
        await asyncio.get_event_loop().run_in_executor(None, process.stdin.flush)
        
        # Read response
        try:
            line = await asyncio.wait_for(
                asyncio.get_event_loop().run_in_executor(None, process.stdout.readline),
                timeout=10.0
            )
            
            if line:
                response = json.loads(line.decode().strip())
                print("   Ping Response: " + str(response))
            else:
                print("   No ping response received")
        except Exception as e:
            print("   Ping Error: " + str(e))
        
        # Step 4: Send ping request with empty params
        print("\n4. Sending ping request with empty params...")
        ping_request2 = {
            "jsonrpc": "2.0",
            "method": "ping",
            "id": "3",
            "params": {}
        }
        
        ping_json2 = json.dumps(ping_request2) + "\n"
        print("   Sending: " + ping_json2.strip())
        
        process.stdin.write(ping_json2.encode())
        await asyncio.get_event_loop().run_in_executor(None, process.stdin.flush)
        
        # Read response
        try:
            line = await asyncio.wait_for(
                asyncio.get_event_loop().run_in_executor(None, process.stdout.readline),
                timeout=10.0
            )
            
            if line:
                response = json.loads(line.decode().strip())
                print("   Ping with params Response: " + str(response))
            else:
                print("   No ping with params response received")
        except Exception as e:
            print("   Ping with params Error: " + str(e))
            
        # Step 5: Send tools/list request (no params field)
        print("\n5. Sending tools/list request (no params field)...")
        tools_request1 = {
            "jsonrpc": "2.0",
            "method": "tools/list",
            "id": "4"
        }
        
        tools_json1 = json.dumps(tools_request1) + "\n"
        print("   Sending: " + tools_json1.strip())
        
        process.stdin.write(tools_json1.encode())
        await asyncio.get_event_loop().run_in_executor(None, process.stdin.flush)
        
        # Read response
        try:
            line = await asyncio.wait_for(
                asyncio.get_event_loop().run_in_executor(None, process.stdout.readline),
                timeout=10.0
            )
            
            if line:
                response = json.loads(line.decode().strip())
                print("   Tools/List Response: " + str(response))
            else:
                print("   No tools/list response received")
        except Exception as e:
            print("   Tools/List Error: " + str(e))
            
        # Step 6: Send tools/list request with empty params
        print("\n6. Sending tools/list request with empty params...")
        tools_request2 = {
            "jsonrpc": "2.0",
            "method": "tools/list",
            "id": "5",
            "params": {}
        }
        
        tools_json2 = json.dumps(tools_request2) + "\n"
        print("   Sending: " + tools_json2.strip())
        
        process.stdin.write(tools_json2.encode())
        await asyncio.get_event_loop().run_in_executor(None, process.stdin.flush)
        
        # Read response
        try:
            line = await asyncio.wait_for(
                asyncio.get_event_loop().run_in_executor(None, process.stdout.readline),
                timeout=10.0
            )
            
            if line:
                response = json.loads(line.decode().strip())
                print("   Tools/List with params Response: " + str(response))
            else:
                print("   No tools/list with params response received")
        except Exception as e:
            print("   Tools/List with params Error: " + str(e))
            
    finally:
        # Clean up
        if process.poll() is None:
            process.terminate()
            try:
                process.wait(timeout=5)
            except subprocess.TimeoutExpired:
                process.kill()
                process.wait()
        
        # Print any remaining stderr output
        stderr_output = process.stderr.read()
        if stderr_output:
            print("   Final stderr output:", stderr_output.decode())

if __name__ == "__main__":
    asyncio.run(test_proper_initialization_sequence())