import { execFile } from "node:child_process";
import { fileURLToPath } from "node:url";
import path from "node:path";

const BIN_DIR = path.dirname(fileURLToPath(import.meta.url));
const HELPER_PATH = path.join(BIN_DIR, "Pricop.PowerPointTools.Helper.exe");

export type BridgeResult = {
  ok: boolean;
  response: string;
};

export async function sendPowerPointCommand(command: string): Promise<BridgeResult> {
  return new Promise<BridgeResult>((resolve) => {
    execFile(
      HELPER_PATH,
      [command],
      {
        windowsHide: true,
        timeout: 5000,
        encoding: "utf8"
      },
      (error, stdout, stderr) => {
        const output = (stdout || stderr || "").trim();

        if (!error) {
          resolve({ ok: true, response: output || "OK" });
          return;
        }

        const code = typeof error.code === "number" ? error.code : -1;
        resolve({
          ok: false,
          response: output || `ERR HELPER_EXIT_${code}`
        });
      }
    );
  });
}
