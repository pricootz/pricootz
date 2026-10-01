using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net.WebSockets;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Web.Script.Serialization;

namespace Pricop.PowerPointTools.StreamDeck
{
    internal static class Program
    {
        private static readonly string LogDirectory = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "Pricop", "PowerPointTools");

        private static readonly string LogPath = Path.Combine(LogDirectory, "streamdeck-native.log");

        [STAThread]
        private static int Main(string[] args)
        {
            Directory.CreateDirectory(LogDirectory);
            Log("=== START 4.1.0 ===");
            Log("Args: " + string.Join(" ", args ?? new string[0]));

            try
            {
                if (args != null && args.Any(a => string.Equals(a, "--self-test", StringComparison.OrdinalIgnoreCase)))
                {
                    Log("SELFTEST OK");
                    return 0;
                }

                var parsed = ParseArguments(args ?? new string[0]);
                string port;
                string pluginUuid;
                string registerEvent;

                if (!parsed.TryGetValue("-port", out port) ||
                    !parsed.TryGetValue("-pluginUUID", out pluginUuid) ||
                    !parsed.TryGetValue("-registerEvent", out registerEvent))
                {
                    Log("ERROR missing Stream Deck startup arguments.");
                    return 10;
                }

                return RunStreamDeckAsync(port, pluginUuid, registerEvent).GetAwaiter().GetResult();
            }
            catch (Exception ex)
            {
                Log("FATAL " + ex);
                return 99;
            }
        }

        private static Dictionary<string, string> ParseArguments(string[] args)
        {
            var result = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

            for (int i = 0; i < args.Length; i++)
            {
                string key = args[i];

                if (!key.StartsWith("-", StringComparison.Ordinal))
                    continue;

                string value = "";
                if (i + 1 < args.Length && !args[i + 1].StartsWith("-", StringComparison.Ordinal))
                    value = args[++i];

                result[key] = value;
            }

            return result;
        }

        private static async Task<int> RunStreamDeckAsync(string port, string pluginUuid, string registerEvent)
        {
            var uri = new Uri("ws://127.0.0.1:" + port);
            using (var socket = new ClientWebSocket())
            {
                Log("Connecting to Stream Deck " + uri);
                await socket.ConnectAsync(uri, CancellationToken.None);
                Log("WebSocket connected.");

                await SendJsonAsync(socket, new Dictionary<string, object>
                {
                    ["event"] = registerEvent,
                    ["uuid"] = pluginUuid
                });

                Log("Registered plugin UUID " + pluginUuid);

                while (socket.State == WebSocketState.Open)
                {
                    string json = await ReceiveTextAsync(socket);
                    if (json == null)
                    {
                        Log("WebSocket closed by Stream Deck.");
                        break;
                    }

                    try
                    {
                        HandleMessageAsync(socket, json).GetAwaiter().GetResult();
                    }
                    catch (Exception ex)
                    {
                        Log("MESSAGE ERROR " + ex);
                    }
                }
            }

            return 0;
        }

        private static async Task HandleMessageAsync(ClientWebSocket socket, string json)
        {
            var serializer = new JavaScriptSerializer();
            var message = serializer.Deserialize<Dictionary<string, object>>(json);

            object eventObject;
            if (!message.TryGetValue("event", out eventObject))
                return;

            string eventName = Convert.ToString(eventObject);
            if (!string.Equals(eventName, "keyDown", StringComparison.Ordinal))
                return;

            string action = message.ContainsKey("action") ? Convert.ToString(message["action"]) : "";
            string context = message.ContainsKey("context") ? Convert.ToString(message["context"]) : "";

            string command = CommandFromAction(action);
            Log("KEYDOWN action=" + action + " command=" + command);

            if (string.IsNullOrEmpty(command))
            {
                Log("Unknown action UUID: " + action);
                await SendFeedbackAsync(socket, context, false);
                return;
            }

            var result = await PowerPointAutomation.ExecuteAsync(command);
            Log("RESULT " + command + " => " + result.Code + " " + result.Message);

            await SendFeedbackAsync(socket, context, result.Code == 0);
        }

        private static string CommandFromAction(string action)
        {
            const string prefix = "com.pricop.powerpoint-tools.";

            if (string.IsNullOrEmpty(action) ||
                !action.StartsWith(prefix, StringComparison.OrdinalIgnoreCase))
                return null;

            return action.Substring(prefix.Length);
        }

        private static async Task SendFeedbackAsync(ClientWebSocket socket, string context, bool ok)
        {
            if (string.IsNullOrEmpty(context))
                return;

            await SendJsonAsync(socket, new Dictionary<string, object>
            {
                ["event"] = ok ? "showOk" : "showAlert",
                ["context"] = context
            });
        }

        private static async Task SendJsonAsync(ClientWebSocket socket, object value)
        {
            var serializer = new JavaScriptSerializer();
            string json = serializer.Serialize(value);
            byte[] bytes = Encoding.UTF8.GetBytes(json);

            await socket.SendAsync(
                new ArraySegment<byte>(bytes),
                WebSocketMessageType.Text,
                true,
                CancellationToken.None);
        }

        private static async Task<string> ReceiveTextAsync(ClientWebSocket socket)
        {
            var buffer = new byte[8192];

            using (var stream = new MemoryStream())
            {
                while (true)
                {
                    WebSocketReceiveResult result;

                    try
                    {
                        result = await socket.ReceiveAsync(
                            new ArraySegment<byte>(buffer),
                            CancellationToken.None);
                    }
                    catch (WebSocketException ex)
                    {
                        Log("WebSocket receive error: " + ex.Message);
                        return null;
                    }

                    if (result.MessageType == WebSocketMessageType.Close)
                    {
                        try
                        {
                            await socket.CloseOutputAsync(
                                WebSocketCloseStatus.NormalClosure,
                                "Closing",
                                CancellationToken.None);
                        }
                        catch { }

                        return null;
                    }

                    stream.Write(buffer, 0, result.Count);

                    if (result.EndOfMessage)
                        break;
                }

                return Encoding.UTF8.GetString(stream.ToArray());
            }
        }

        internal static void Log(string message)
        {
            try
            {
                File.AppendAllText(
                    LogPath,
                    DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff") +
                    " " + message + Environment.NewLine);
            }
            catch { }
        }
    }

    internal sealed class CommandResult
    {
        public int Code { get; set; }
        public string Message { get; set; }

        public static CommandResult Ok(string message = "OK")
            => new CommandResult { Code = 0, Message = message };

        public static CommandResult Fail(int code, string message)
            => new CommandResult { Code = code, Message = message };
    }

    internal static class PowerPointAutomation
    {
        private const uint OBJID_NATIVEOM = 0xFFFFFFF0;
        private static readonly Guid IID_IDispatch =
            new Guid("00020400-0000-0000-C000-000000000046");

        public static Task<CommandResult> ExecuteAsync(string command)
        {
            var tcs = new TaskCompletionSource<CommandResult>();

            var thread = new Thread(() =>
            {
                OleMessageFilter.Register();

                object appObject = null;

                try
                {
                    appObject = GetPowerPointApplication();

                    if (appObject == null)
                    {
                        tcs.SetResult(CommandResult.Fail(2, "POWERPOINT_NOT_FOUND"));
                        return;
                    }

                    dynamic app = appObject;

                    if (string.Equals(command, "test-connection", StringComparison.OrdinalIgnoreCase))
                    {
                        string version = "";
                        try { version = Convert.ToString(app.Version); } catch { }

                        tcs.SetResult(CommandResult.Ok("POWERPOINT " + version));
                        return;
                    }

                    tcs.SetResult(ExecuteWithRetry(() => ExecuteCommand(app, command)));
                }
                catch (Exception ex)
                {
                    Program.Log("POWERPOINT FATAL " + command + ": " + ex);
                    tcs.SetResult(CommandResult.Fail(5, ex.GetType().Name + ": " + ex.Message));
                }
                finally
                {
                    try
                    {
                        if (appObject != null && Marshal.IsComObject(appObject))
                            Marshal.FinalReleaseComObject(appObject);
                    }
                    catch { }

                    OleMessageFilter.Revoke();
                }
            });

            thread.IsBackground = true;
            thread.Name = "PricopTools.PowerPointSTA";
            thread.SetApartmentState(ApartmentState.STA);
            thread.Start();

            return tcs.Task;
        }

        private static CommandResult ExecuteWithRetry(Func<CommandResult> action)
        {
            const int RPC_E_CALL_REJECTED = unchecked((int)0x80010001);
            const int RPC_E_SERVERCALL_RETRYLATER = unchecked((int)0x8001010A);

            for (int attempt = 0; attempt < 12; attempt++)
            {
                try
                {
                    return action();
                }
                catch (COMException ex) when (
                    ex.HResult == RPC_E_CALL_REJECTED ||
                    ex.HResult == RPC_E_SERVERCALL_RETRYLATER)
                {
                    Thread.Sleep(100 + attempt * 25);
                }
            }

            return CommandResult.Fail(5, "POWERPOINT_BUSY");
        }

        private static object GetPowerPointApplication()
        {
            for (int attempt = 0; attempt < 4; attempt++)
            {
                try
                {
                    object active = Marshal.GetActiveObject("PowerPoint.Application");
                    if (active != null)
                    {
                        Program.Log("ATTACH ROT");
                        return active;
                    }
                }
                catch (COMException)
                {
                    Thread.Sleep(75);
                }
            }

            object native = GetPowerPointViaNativeObjectModel();

            if (native != null)
            {
                Program.Log("ATTACH OBJID_NATIVEOM");
                return native;
            }

            Program.Log("ATTACH FAILED");
            return null;
        }

        private static object GetPowerPointViaNativeObjectModel()
        {
            var powerPointPids = new HashSet<uint>(
                Process.GetProcessesByName("POWERPNT")
                    .Select(p => unchecked((uint)p.Id)));

            if (powerPointPids.Count == 0)
                return null;

            object foundApplication = null;

            EnumWindows((top, _) =>
            {
                if (foundApplication != null)
                    return false;

                uint pid;
                GetWindowThreadProcessId(top, out pid);

                if (!powerPointPids.Contains(pid))
                    return true;

                EnumChildWindows(top, (child, __) =>
                {
                    if (foundApplication != null)
                        return false;

                    var className = new StringBuilder(256);

                    if (GetClassName(child, className, className.Capacity) <= 0)
                        return true;

                    if (!string.Equals(
                        className.ToString(),
                        "paneClassDC",
                        StringComparison.Ordinal))
                        return true;

                    try
                    {
                        object nativeObject = null;
                        Guid iid = IID_IDispatch;

                        int hr = AccessibleObjectFromWindow(
                            child,
                            OBJID_NATIVEOM,
                            ref iid,
                            ref nativeObject);

                        if (hr < 0 || nativeObject == null)
                            return true;

                        dynamic documentWindow = nativeObject;
                        object application = documentWindow.Application;

                        if (application != null)
                        {
                            foundApplication = application;

                            try
                            {
                                if (Marshal.IsComObject(nativeObject))
                                    Marshal.FinalReleaseComObject(nativeObject);
                            }
                            catch { }

                            return false;
                        }
                    }
                    catch (Exception ex)
                    {
                        Program.Log(
                            "NATIVEOM candidate failed: " +
                            ex.GetType().Name + " " + ex.Message);
                    }

                    return true;
                }, IntPtr.Zero);

                return foundApplication == null;
            }, IntPtr.Zero);

            return foundApplication;
        }

        private static CommandResult ExecuteCommand(dynamic app, string command)
        {
            switch ((command ?? "").ToLowerInvariant())
            {
                case "align-left":
                    return WithShapes((object)app, 2, r => r.Align(0, 0));

                case "align-center":
                    return WithShapes((object)app, 2, r => r.Align(1, 0));

                case "align-right":
                    return WithShapes((object)app, 2, r => r.Align(2, 0));

                case "align-top":
                    return WithShapes((object)app, 2, r => r.Align(3, 0));

                case "align-middle":
                    return WithShapes((object)app, 2, r => r.Align(4, 0));

                case "align-bottom":
                    return WithShapes((object)app, 2, r => r.Align(5, 0));

                case "distribute-horizontal":
                    return WithShapes((object)app, 3, r => r.Distribute(0, 0));

                case "distribute-vertical":
                    return WithShapes((object)app, 3, r => r.Distribute(1, 0));

                case "same-width":
                    return WithShapes((object)app, 2, r =>
                    {
                        float value = Convert.ToSingle(r[1].Width);

                        for (int i = 2; i <= Convert.ToInt32(r.Count); i++)
                            r[i].Width = value;
                    });

                case "same-height":
                    return WithShapes((object)app, 2, r =>
                    {
                        float value = Convert.ToSingle(r[1].Height);

                        for (int i = 2; i <= Convert.ToInt32(r.Count); i++)
                            r[i].Height = value;
                    });

                case "same-size":
                    return WithShapes((object)app, 2, r =>
                    {
                        float width = Convert.ToSingle(r[1].Width);
                        float height = Convert.ToSingle(r[1].Height);

                        for (int i = 2; i <= Convert.ToInt32(r.Count); i++)
                        {
                            r[i].Width = width;
                            r[i].Height = height;
                        }
                    });

                case "rectangle":
                    return WithShapes((object)app, 1, r =>
                    {
                        for (int i = 1; i <= Convert.ToInt32(r.Count); i++)
                            if (Convert.ToInt32(r[i].Type) == 1)
                                r[i].AutoShapeType = 1;
                    });

                case "rounded-rectangle":
                    return WithShapes((object)app, 1, r =>
                    {
                        for (int i = 1; i <= Convert.ToInt32(r.Count); i++)
                            if (Convert.ToInt32(r[i].Type) == 1)
                                r[i].AutoShapeType = 5;
                    });

                case "shadow-off":
                    return WithShapes((object)app, 1, r =>
                    {
                        for (int i = 1; i <= Convert.ToInt32(r.Count); i++)
                            r[i].Shadow.Visible = 0;
                    });

                case "shadow-on":
                    return WithShapes((object)app, 1, r =>
                    {
                        for (int i = 1; i <= Convert.ToInt32(r.Count); i++)
                        {
                            dynamic shadow = r[i].Shadow;
                            shadow.Visible = -1;
                            try { shadow.Transparency = 0.55f; } catch { }
                            try { shadow.Blur = 7f; } catch { }
                            try { shadow.OffsetX = 1.5f; shadow.OffsetY = 2f; } catch { }
                        }
                    });

                case "border-off":
                    return WithShapes((object)app, 1, r =>
                    {
                        for (int i = 1; i <= Convert.ToInt32(r.Count); i++)
                            r[i].Line.Visible = 0;
                    });

                case "border-on":
                    return WithShapes((object)app, 1, r =>
                    {
                        for (int i = 1; i <= Convert.ToInt32(r.Count); i++)
                        {
                            r[i].Line.Visible = -1;
                            r[i].Line.Weight = 1f;
                        }
                    });

                case "match-style":
                    return WithShapes((object)app, 2, r =>
                    {
                        r[1].PickUp();

                        for (int i = 2; i <= Convert.ToInt32(r.Count); i++)
                            r[i].Apply();
                    });

                case "clean-boxes":
                    return WithShapes((object)app, 1, r =>
                    {
                        float height = Convert.ToSingle(r[1].Height);

                        for (int i = 1; i <= Convert.ToInt32(r.Count); i++)
                        {
                            if (Convert.ToInt32(r[i].Type) == 1)
                                r[i].AutoShapeType = 1;

                            r[i].Shadow.Visible = 0;
                            r[i].Line.Visible = 0;

                            if (i > 1)
                                r[i].Height = height;
                        }
                    });

                default:
                    return CommandResult.Fail(4, "UNKNOWN_COMMAND " + command);
            }
        }

        private static CommandResult WithShapes(
            object appObject,
            int minimum,
            Action<dynamic> action)
        {
            dynamic app = appObject;
            dynamic window;

            try
            {
                window = app.ActiveWindow;
            }
            catch
            {
                return CommandResult.Fail(3, "NO_ACTIVE_WINDOW");
            }

            if (window == null)
                return CommandResult.Fail(3, "NO_ACTIVE_WINDOW");

            dynamic selection = window.Selection;

            if (selection == null || Convert.ToInt32(selection.Type) != 2)
                return CommandResult.Fail(3, "SELECT_SHAPES");

            dynamic range = selection.ShapeRange;
            int count = Convert.ToInt32(range.Count);

            if (count < minimum)
                return CommandResult.Fail(3, "NEED_" + minimum + "_SHAPES");

            action(range);
            return CommandResult.Ok();
        }

        private delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

        [DllImport("user32.dll")]
        private static extern bool EnumWindows(
            EnumWindowsProc lpEnumFunc,
            IntPtr lParam);

        [DllImport("user32.dll")]
        private static extern bool EnumChildWindows(
            IntPtr hWndParent,
            EnumWindowsProc lpEnumFunc,
            IntPtr lParam);

        [DllImport("user32.dll")]
        private static extern uint GetWindowThreadProcessId(
            IntPtr hWnd,
            out uint lpdwProcessId);

        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        private static extern int GetClassName(
            IntPtr hWnd,
            StringBuilder lpClassName,
            int nMaxCount);

        [DllImport("oleacc.dll")]
        private static extern int AccessibleObjectFromWindow(
            IntPtr hwnd,
            uint dwId,
            ref Guid riid,
            [In, Out, MarshalAs(UnmanagedType.IUnknown)] ref object ppvObject);

        private static class OleMessageFilter
        {
            [ComImport]
            [Guid("00000016-0000-0000-C000-000000000046")]
            [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
            private interface IOleMessageFilter
            {
                [PreserveSig]
                int HandleInComingCall(
                    int dwCallType,
                    IntPtr hTaskCaller,
                    int dwTickCount,
                    IntPtr lpInterfaceInfo);

                [PreserveSig]
                int RetryRejectedCall(
                    IntPtr hTaskCallee,
                    int dwTickCount,
                    int dwRejectType);

                [PreserveSig]
                int MessagePending(
                    IntPtr hTaskCallee,
                    int dwTickCount,
                    int dwPendingType);
            }

            private sealed class Filter : IOleMessageFilter
            {
                public int HandleInComingCall(
                    int dwCallType,
                    IntPtr hTaskCaller,
                    int dwTickCount,
                    IntPtr lpInterfaceInfo) => 0;

                public int RetryRejectedCall(
                    IntPtr hTaskCallee,
                    int dwTickCount,
                    int dwRejectType) =>
                    dwRejectType == 2 ? 100 : -1;

                public int MessagePending(
                    IntPtr hTaskCallee,
                    int dwTickCount,
                    int dwPendingType) => 2;
            }

            [DllImport("Ole32.dll")]
            private static extern int CoRegisterMessageFilter(
                IOleMessageFilter newFilter,
                out IOleMessageFilter oldFilter);

            public static void Register()
            {
                try
                {
                    IOleMessageFilter ignored;
                    CoRegisterMessageFilter(new Filter(), out ignored);
                }
                catch { }
            }

            public static void Revoke()
            {
                try
                {
                    IOleMessageFilter ignored;
                    CoRegisterMessageFilter(null, out ignored);
                }
                catch { }
            }
        }
    }
}
