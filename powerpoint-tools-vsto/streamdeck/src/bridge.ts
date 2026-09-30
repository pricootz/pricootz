import net from "node:net";

const PIPE_PATH = "\\\\.\\pipe\\PricopPowerPointTools";

export async function sendPowerPointCommand(command: string): Promise<boolean> {
  return new Promise<boolean>((resolve) => {
    const socket = net.createConnection(PIPE_PATH);
    let settled = false;
    let buffer = "";

    const finish = (value: boolean) => {
      if (settled) return;
      settled = true;
      clearTimeout(timer);
      socket.destroy();
      resolve(value);
    };

    const timer = setTimeout(() => finish(false), 1800);

    socket.on("connect", () => {
      socket.write(command + "\n");
    });

    socket.on("data", (chunk) => {
      buffer += chunk.toString("utf8");
      const lineEnd = buffer.indexOf("\n");
      if (lineEnd >= 0) {
        const response = buffer.slice(0, lineEnd).trim().toUpperCase();
        finish(response === "OK");
      }
    });

    socket.on("error", () => finish(false));
    socket.on("close", () => {
      if (!settled) finish(false);
    });
  });
}
