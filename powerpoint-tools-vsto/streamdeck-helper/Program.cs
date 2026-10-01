using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;

namespace Pricop.PowerPointTools.Helper
{
    internal static class Program
    {
        private const uint OBJID_NATIVEOM = 0xFFFFFFF0;
        private static readonly Guid IID_IDispatch = new Guid("00020400-0000-0000-C000-000000000046");

        private static readonly string LogDirectory = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "Pricop", "PowerPointTools");

        private static readonly string LogPath = Path.Combine(LogDirectory, "streamdeck-helper.log");

        [STAThread]
        private static int Main(string[] args)
        {
            Directory.CreateDirectory(LogDirectory);

            if (args.Length == 0)
                return Fail(4, "ERR NO_COMMAND");

            string command = (args[0] ?? "").Trim().ToLowerInvariant();

            if (command == "self-test")
            {
                Console.WriteLine("OK SELFTEST");
                return 0;
            }

            OleMessageFilter.Register();

            try
            {
                dynamic app = GetPowerPointApplication();
                if (app == null)
                    return Fail(2, "ERR POWERPOINT_NOT_FOUND");

                if (command == "ping" || command == "test-connection")
                {
                    string version = "";
                    try { version = Convert.ToString(app.Version); } catch { }
                    Console.WriteLine("OK POWERPOINT " + version);
                    return 0;
                }

                int result = ExecuteWithRetry(() => ExecuteCommand(app, command));
                return result;
            }
            catch (Exception ex)
            {
                Log("FATAL " + command + ": " + ex);
                Console.WriteLine("ERR " + ex.GetType().Name);
                return 5;
            }
            finally
            {
                OleMessageFilter.Revoke();
            }
        }

        private static int ExecuteWithRetry(Func<int> action)
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
                    Thread.Sleep(80 + attempt * 20);
                }
            }

            return Fail(5, "ERR POWERPOINT_BUSY");
        }

        private static dynamic GetPowerPointApplication()
        {
            // Fast path: PowerPoint is already registered in the Running Object Table.
            for (int attempt = 0; attempt < 3; attempt++)
            {
                try
                {
                    object active = Marshal.GetActiveObject("PowerPoint.Application");
                    if (active != null)
                    {
                        Log("ATTACH ROT");
                        return active;
                    }
                }
                catch (COMException)
                {
                    Thread.Sleep(60);
                }
            }

            // Microsoft documents that Office may not register in the ROT until it
            // loses focus. PowerPoint also exposes its native object model from its
            // paneClassDC document window via OBJID_NATIVEOM, so use that as a
            // focus-independent fallback.
            dynamic native = GetPowerPointViaNativeObjectModel();
            if (native != null)
            {
                Log("ATTACH OBJID_NATIVEOM");
                return native;
            }

            Log("ATTACH FAILED");
            return null;
        }

        private static dynamic GetPowerPointViaNativeObjectModel()
        {
            var pids = new HashSet<uint>(
                Process.GetProcessesByName("POWERPNT")
                    .Select(p => unchecked((uint)p.Id)));

            if (pids.Count == 0)
                return null;

            object foundApp = null;

            EnumWindows((top, _) =>
            {
                if (foundApp != null)
                    return false;

                GetWindowThreadProcessId(top, out uint pid);
                if (!pids.Contains(pid))
                    return true;

                EnumChildWindows(top, (child, __) =>
                {
                    if (foundApp != null)
                        return false;

                    var className = new StringBuilder(256);
                    if (GetClassName(child, className, className.Capacity) <= 0)
                        return true;

                    if (!string.Equals(className.ToString(), "paneClassDC", StringComparison.Ordinal))
                        return true;

                    try
                    {
                        object native = null;
                        Guid iid = IID_IDispatch;
                        int hr = AccessibleObjectFromWindow(child, OBJID_NATIVEOM, ref iid, ref native);
                        if (hr >= 0 && native != null)
                        {
                            dynamic documentWindow = native;
                            object app = documentWindow.Application;
                            if (app != null)
                            {
                                foundApp = app;
                                return false;
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        Log("NATIVEOM candidate failed: " + ex.GetType().Name + " " + ex.Message);
                    }

                    return true;
                }, IntPtr.Zero);

                return foundApp == null;
            }, IntPtr.Zero);

            return foundApp;
        }

        private static int ExecuteCommand(dynamic app, string command)
        {
            switch (command)
            {
                case "align-left":
                    return WithShapes(app, 2, r => r.Align(0, 0));
                case "align-center":
                    return WithShapes(app, 2, r => r.Align(1, 0));
                case "align-right":
                    return WithShapes(app, 2, r => r.Align(2, 0));
                case "align-top":
                    return WithShapes(app, 2, r => r.Align(3, 0));
                case "align-middle":
                    return WithShapes(app, 2, r => r.Align(4, 0));
                case "align-bottom":
                    return WithShapes(app, 2, r => r.Align(5, 0));

                case "distribute-horizontal":
                    return WithShapes(app, 3, r => r.Distribute(0, 0));
                case "distribute-vertical":
                    return WithShapes(app, 3, r => r.Distribute(1, 0));

                case "same-width":
                    return WithShapes(app, 2, r =>
                    {
                        float width = Convert.ToSingle(r[1].Width);
                        for (int i = 2; i <= Convert.ToInt32(r.Count); i++)
                            r[i].Width = width;
                    });

                case "same-height":
                    return WithShapes(app, 2, r =>
                    {
                        float height = Convert.ToSingle(r[1].Height);
                        for (int i = 2; i <= Convert.ToInt32(r.Count); i++)
                            r[i].Height = height;
                    });

                case "same-size":
                    return WithShapes(app, 2, r =>
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
                    return WithShapes(app, 1, r =>
                    {
                        for (int i = 1; i <= Convert.ToInt32(r.Count); i++)
                            if (Convert.ToInt32(r[i].Type) == 1)
                                r[i].AutoShapeType = 1;
                    });

                case "rounded-rectangle":
                    return WithShapes(app, 1, r =>
                    {
                        for (int i = 1; i <= Convert.ToInt32(r.Count); i++)
                            if (Convert.ToInt32(r[i].Type) == 1)
                                r[i].AutoShapeType = 5;
                    });

                case "shadow-off":
                    return WithShapes(app, 1, r =>
                    {
                        for (int i = 1; i <= Convert.ToInt32(r.Count); i++)
                            r[i].Shadow.Visible = 0;
                    });

                case "shadow-on":
                    return WithShapes(app, 1, r =>
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
                    return WithShapes(app, 1, r =>
                    {
                        for (int i = 1; i <= Convert.ToInt32(r.Count); i++)
                            r[i].Line.Visible = 0;
                    });

                case "border-on":
                    return WithShapes(app, 1, r =>
                    {
                        for (int i = 1; i <= Convert.ToInt32(r.Count); i++)
                        {
                            r[i].Line.Visible = -1;
                            r[i].Line.Weight = 1f;
                        }
                    });

                case "match-style":
                    return WithShapes(app, 2, r =>
                    {
                        r[1].PickUp();
                        for (int i = 2; i <= Convert.ToInt32(r.Count); i++)
                            r[i].Apply();
                    });

                case "clean-boxes":
                    return WithShapes(app, 1, r =>
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
                    return Fail(4, "ERR UNKNOWN_COMMAND " + command);
            }
        }

        private static int WithShapes(dynamic app, int minimum, Action<dynamic> action)
        {
            dynamic window;
            try { window = app.ActiveWindow; }
            catch { return Fail(3, "ERR NO_ACTIVE_WINDOW"); }

            if (window == null)
                return Fail(3, "ERR NO_ACTIVE_WINDOW");

            dynamic selection = window.Selection;
            if (selection == null || Convert.ToInt32(selection.Type) != 2)
                return Fail(3, "ERR SELECT_SHAPES");

            dynamic range = selection.ShapeRange;
            int count = Convert.ToInt32(range.Count);
            if (count < minimum)
                return Fail(3, "ERR NEED_" + minimum + "_SHAPES");

            action(range);

            Console.WriteLine("OK");
            return 0;
        }

        private static int Fail(int exitCode, string message)
        {
            Log(message);
            Console.WriteLine(message);
            return exitCode;
        }

        private static void Log(string message)
        {
            try
            {
                File.AppendAllText(
                    LogPath,
                    DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff") + " " + message + Environment.NewLine);
            }
            catch { }
        }

        private delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

        [DllImport("user32.dll")]
        private static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);

        [DllImport("user32.dll")]
        private static extern bool EnumChildWindows(IntPtr hWndParent, EnumWindowsProc lpEnumFunc, IntPtr lParam);

        [DllImport("user32.dll")]
        private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        private static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);

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
                int HandleInComingCall(int dwCallType, IntPtr hTaskCaller, int dwTickCount, IntPtr lpInterfaceInfo);

                [PreserveSig]
                int RetryRejectedCall(IntPtr hTaskCallee, int dwTickCount, int dwRejectType);

                [PreserveSig]
                int MessagePending(IntPtr hTaskCallee, int dwTickCount, int dwPendingType);
            }

            private sealed class Filter : IOleMessageFilter
            {
                public int HandleInComingCall(int dwCallType, IntPtr hTaskCaller, int dwTickCount, IntPtr lpInterfaceInfo)
                    => 0;

                public int RetryRejectedCall(IntPtr hTaskCallee, int dwTickCount, int dwRejectType)
                    => dwRejectType == 2 ? 100 : -1;

                public int MessagePending(IntPtr hTaskCallee, int dwTickCount, int dwPendingType)
                    => 2;
            }

            [DllImport("Ole32.dll")]
            private static extern int CoRegisterMessageFilter(
                IOleMessageFilter newFilter,
                out IOleMessageFilter oldFilter);

            public static void Register()
            {
                try
                {
                    CoRegisterMessageFilter(new Filter(), out _);
                }
                catch { }
            }

            public static void Revoke()
            {
                try
                {
                    CoRegisterMessageFilter(null, out _);
                }
                catch { }
            }
        }
    }
}
