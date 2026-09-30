using System;
using System.IO;
using System.Net;
using System.Net.Sockets;
using System.Text;
using System.Threading;
using System.Windows.Forms;

namespace Pricop.PowerPointTools
{
    internal sealed class PipeCommandServer : IDisposable
    {
        public const int Port = 32145;

        private readonly CancellationTokenSource _cts = new CancellationTokenSource();
        private readonly Thread _thread;
        private readonly Control _dispatcher;
        private readonly TcpListener _listener;
        private readonly string _logPath;

        public PipeCommandServer()
        {
            _logPath = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "Pricop", "PowerPointTools", "bridge.log");

            Directory.CreateDirectory(Path.GetDirectoryName(_logPath));

            _dispatcher = new Control();
            _dispatcher.CreateControl();

            _listener = new TcpListener(IPAddress.Loopback, Port);

            _thread = new Thread(ListenLoop)
            {
                IsBackground = true,
                Name = "PricopTools.LoopbackBridge"
            };
        }

        public void Start()
        {
            try
            {
                _listener.Start();
                Log("Bridge START TCP 127.0.0.1:" + Port);
                _thread.Start();
            }
            catch (Exception ex)
            {
                Log("Bridge START ERROR: " + ex);
                throw;
            }
        }

        private void ListenLoop()
        {
            while (!_cts.IsCancellationRequested)
            {
                TcpClient client = null;

                try
                {
                    client = _listener.AcceptTcpClient();
                    client.ReceiveTimeout = 4000;
                    client.SendTimeout = 4000;

                    using (client)
                    using (var stream = client.GetStream())
                    using (var reader = new StreamReader(stream, Encoding.UTF8, false, 1024, true))
                    using (var writer = new StreamWriter(stream, new UTF8Encoding(false), 1024, true) { AutoFlush = true })
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

                                if (!completed.Wait(3000))
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
                catch (SocketException)
                {
                    if (_cts.IsCancellationRequested) return;
                    Thread.Sleep(100);
                }
                catch (Exception ex)
                {
                    if (_cts.IsCancellationRequested) return;
                    Log("TCP ERROR: " + ex);
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

            try { _listener.Stop(); } catch { }

            try
            {
                using (var client = new TcpClient())
                    client.Connect(IPAddress.Loopback, Port);
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
