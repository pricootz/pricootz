using System;
using System.Collections.Concurrent;
using System.IO;
using System.IO.Pipes;
using System.Threading;
using System.Windows.Forms;

namespace Pricop.PowerPointTools
{
    internal sealed class PipeCommandServer : IDisposable
    {
        public const string PipeName = "PricopPowerPointTools";

        private readonly CancellationTokenSource _cts = new CancellationTokenSource();
        private readonly Thread _thread;
        private readonly Control _dispatcher;
        private readonly string _logPath;

        public PipeCommandServer()
        {
            _logPath = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "Pricop", "PowerPointTools", "bridge.log");

            Directory.CreateDirectory(Path.GetDirectoryName(_logPath));

            // Constructed on the PowerPoint UI thread during ThisAddIn_Startup.
            // CreateControl gives us a stable Win32 handle so BeginInvoke can marshal
            // commands back to the Office UI thread.
            _dispatcher = new Control();
            _dispatcher.CreateControl();

            _thread = new Thread(ListenLoop)
            {
                IsBackground = true,
                Name = "PricopTools.NamedPipe"
            };
        }

        public void Start()
        {
            Log("Bridge START. Pipe=" + PipeName);
            _thread.Start();
        }

        private void ListenLoop()
        {
            while (!_cts.IsCancellationRequested)
            {
                try
                {
                    using (var server = new NamedPipeServerStream(
                        PipeName,
                        PipeDirection.InOut,
                        1,
                        PipeTransmissionMode.Byte,
                        PipeOptions.Asynchronous))
                    {
                        server.WaitForConnection();
                        if (_cts.IsCancellationRequested) return;

                        using (var reader = new StreamReader(server))
                        using (var writer = new StreamWriter(server) { AutoFlush = true })
                        {
                            var command = (reader.ReadLine() ?? "").Trim();
                            Log("RX " + command);

                            if (string.IsNullOrWhiteSpace(command))
                            {
                                writer.WriteLine("ERR EMPTY");
                                continue;
                            }

                            if (string.Equals(command, "ping", StringComparison.OrdinalIgnoreCase))
                            {
                                writer.WriteLine("OK PONG");
                                Log("TX OK PONG");
                                continue;
                            }

                            string response = "ERR INTERNAL";
                            using (var completed = new ManualResetEventSlim(false))
                            {
                                try
                                {
                                    _dispatcher.BeginInvoke(new Action(() =>
                                    {
                                        try
                                        {
                                            bool ok = PowerPointCommands.Execute(command, false);
                                            response = ok ? "OK" : "ERR COMMAND";
                                            Log("EXEC " + command + " => " + response);
                                        }
                                        catch (Exception ex)
                                        {
                                            response = "ERR " + ex.GetType().Name;
                                            Log("EXEC ERROR " + command + ": " + ex);
                                        }
                                        finally
                                        {
                                            completed.Set();
                                        }
                                    }));

                                    if (!completed.Wait(2500))
                                    {
                                        response = "ERR TIMEOUT";
                                        Log("TIMEOUT " + command);
                                    }
                                }
                                catch (Exception ex)
                                {
                                    response = "ERR DISPATCH";
                                    Log("DISPATCH ERROR " + command + ": " + ex);
                                }
                            }

                            writer.WriteLine(response);
                            Log("TX " + response);
                        }
                    }
                }
                catch (Exception ex)
                {
                    if (_cts.IsCancellationRequested) return;
                    Log("PIPE ERROR: " + ex);
                    Thread.Sleep(100);
                }
            }
        }

        private void Log(string message)
        {
            try
            {
                File.AppendAllText(
                    _logPath,
                    DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff") + " " + message + Environment.NewLine);
            }
            catch { }
        }

        public void Dispose()
        {
            Log("Bridge STOP");
            _cts.Cancel();

            try
            {
                using (var client = new NamedPipeClientStream(".", PipeName, PipeDirection.Out))
                    client.Connect(100);
            }
            catch { }

            try
            {
                if (_thread.IsAlive) _thread.Join(750);
            }
            catch { }

            try { _dispatcher.Dispose(); } catch { }
            _cts.Dispose();
        }
    }
}
