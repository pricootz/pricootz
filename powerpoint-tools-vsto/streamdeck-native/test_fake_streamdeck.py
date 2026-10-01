import asyncio, json, os, subprocess, sys, time
import websockets

PORT = 37654
PLUGIN_UUID = "com.pricop.powerpoint-tools"
REGISTER_EVENT = "registerPlugin"
EXE = os.path.abspath(r"streamdeck-native\bin\Release\net48\Pricop.PowerPointTools.StreamDeck.exe")

async def handler(ws):
    reg_raw = await asyncio.wait_for(ws.recv(), timeout=5)
    reg = json.loads(reg_raw)
    assert reg.get("event") == REGISTER_EVENT, reg
    assert reg.get("uuid") == PLUGIN_UUID, reg

    await ws.send(json.dumps({
        "action": "com.pricop.powerpoint-tools.test-connection",
        "context": "test-context",
        "device": "test-device",
        "event": "keyDown",
        "payload": {
            "controller": "Keypad",
            "coordinates": {"column": 0, "row": 0},
            "isInMultiAction": False,
            "settings": {},
            "state": 0,
            "userDesiredState": 0
        }
    }))

    response_raw = await asyncio.wait_for(ws.recv(), timeout=8)
    response = json.loads(response_raw)
    assert response.get("event") == "showAlert", response
    assert response.get("context") == "test-context", response

async def main():
    server = await websockets.serve(handler, "127.0.0.1", PORT)

    proc = subprocess.Popen([
        EXE,
        "-port", str(PORT),
        "-pluginUUID", PLUGIN_UUID,
        "-registerEvent", REGISTER_EVENT,
        "-info", "{}"
    ])

    try:
        await asyncio.wait_for(server.wait_closed(), timeout=15)
    except asyncio.TimeoutError:
        server.close()
        await server.wait_closed()
    finally:
        if proc.poll() is None:
            proc.terminate()
            try:
                proc.wait(timeout=3)
            except subprocess.TimeoutExpired:
                proc.kill()

if __name__ == "__main__":
    async def run():
        done = asyncio.Event()

        async def wrapped_handler(ws):
            try:
                await handler(ws)
                done.set()
            except Exception:
                done.set()
                raise

        server = await websockets.serve(wrapped_handler, "127.0.0.1", PORT)
        proc = subprocess.Popen([
            EXE,
            "-port", str(PORT),
            "-pluginUUID", PLUGIN_UUID,
            "-registerEvent", REGISTER_EVENT,
            "-info", "{}"
        ])

        try:
            await asyncio.wait_for(done.wait(), timeout=15)
        finally:
            server.close()
            await server.wait_closed()
            if proc.poll() is None:
                proc.terminate()
                try:
                    proc.wait(timeout=3)
                except subprocess.TimeoutExpired:
                    proc.kill()

    asyncio.run(run())
    print("FAKE STREAM DECK TEST OK")
