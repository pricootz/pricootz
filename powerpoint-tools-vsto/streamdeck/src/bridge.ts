import net from "node:net";
import fs from "node:fs";
import path from "node:path";

const HOST = "127.0.0.1";
const FIRST_PORT = 32145;
const LAST_PORT = 32154;

export type BridgeResult = {
  ok: boolean;
  response: string;
};

function candidatePorts(): number[] {
  const ports: number[] = [];
  const localAppData = process.env.LOCALAPPDATA;

  if (localAppData) {
    try {
      const portFile = path.join(localAppData, "Pricop", "PowerPointTools", "bridge.port");
      const value = Number.parseInt(fs.readFileSync(portFile, "utf8").trim(), 10);
      if (Number.isInteger(value) && value >= FIRST_PORT && value <= LAST_PORT) {
        ports.push(value);
      }
    } catch {
      // PowerPoint may not have created the port file yet.
    }
  }

  for (let port = FIRST_PORT; port <= LAST_PORT; port++) {
    if (!ports.includes(port)) ports.push(port);
  }

  return ports;
}

function sendToPort(port: number, command: string): Promise<BridgeResult> {
  return new Promise<BridgeResult>((resolve) => {
    const socket = net.createConnection({ host: HOST, port });
    let settled = false;
    let buffer = "";

    const finish = (ok: boolean, response: string) => {
      if (settled) return;
      settled = true;
      clearTimeout(timer);
      socket.destroy();
      resolve({ ok, response });
    };

    const timer = setTimeout(() => finish(false, "ERR CLIENT_TIMEOUT " + port), 1200);

    socket.setNoDelay(true);

    socket.on("connect", () => {
      socket.write(command + "\n", "utf8");
    });

    socket.on("data", (chunk) => {
      buffer += chunk.toString("utf8");
      const lineEnd = buffer.indexOf("\n");
      if (lineEnd >= 0) {
        const response = buffer.slice(0, lineEnd).trim();
        finish(response.toUpperCase().startsWith("OK"), response);
      }
    });

    socket.on("error", (err) => finish(false, "ERR CONNECT " + port + " " + err.message));
    socket.on("close", () => {
      if (!settled) finish(false, "ERR CLOSED " + port);
    });
  });
}

export async function sendPowerPointCommand(command: string): Promise<BridgeResult> {
  let last: BridgeResult = { ok: false, response: "ERR NO_BRIDGE" };

  for (const port of candidatePorts()) {
    const result = await sendToPort(port, command);
    if (result.ok) return result;

    last = result;

    // A reachable bridge may legitimately reject a PowerPoint command because
    // the selection is invalid. In that case, don't continue scanning ports.
    if (!result.response.startsWith("ERR CONNECT") &&
        !result.response.startsWith("ERR CLIENT_TIMEOUT") &&
        !result.response.startsWith("ERR CLOSED")) {
      return result;
    }
  }

  return last;
}
