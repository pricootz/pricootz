import net from "node:net";

const PIPE_PATH = "\\\\.\\pipe\\PricopPowerPointTools";

export type BridgeResult = {
  ok: boolean;
  response: string;
};

export async function sendPowerPointCommand(command: string): Promise<BridgeResult> {
  return new Promise<BridgeResult>((resolve) => {
    const socket = net.createConnection(PIPE_PATH);
    let settled = false;
    let buffer = "";

    const finish = (ok: boolean, response: string) => {
      if (settled) return;
      settled = true;
      clearTimeout(timer);
      socket.destroy();
      resolve({ ok, response });
    };

    const timer = setTimeout(() => finish(false, "ERR CLIENT_TIMEOUT"), 3200);

    socket.on("connect", () => {
      socket.write(command + "\n");
    });

    socket.on("data", (chunk) => {
      buffer += chunk.toString("utf8");
      const lineEnd = buffer.indexOf("\n");
      if (lineEnd >= 0) {
        const response = buffer.slice(0, lineEnd).trim();
        finish(response.toUpperCase().startsWith("OK"), response);
      }
    });

    socket.on("error", (err) => finish(false, "ERR CONNECT " + err.message));
    socket.on("close", () => {
      if (!settled) finish(false, "ERR CLOSED");
    });
  });
}
