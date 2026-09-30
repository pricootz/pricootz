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
        private sealed class Request : IDisposable
        {
            public readonly string Command;
            public readonly ManualResetEventSlim Completed = new ManualResetEventSlim(false);
            public bool Success;

            public Request(string command) { Command = command; }
            public void Dispose() { Completed.Dispose(); }
        }

        public const string PipeName = "PricopPowerPointTools";
        private readonly ConcurrentQueue<Request> _requests = new ConcurrentQueue<Request>();
        private readonly CancellationTokenSource _cts = new CancellationTokenSource();
        private readonly Thread _thread;
        private readonly Timer _timer;

        public PipeCommandServer()
        {
            _timer = new Timer { Interval = 30 };
            _timer.Tick += Timer_Tick;

            _thread = new Thread(ListenLoop)
            {
                IsBackground = true,
                Name = "PricopTools.NamedPipe"
            };
        }

        public void Start()
        {
            _timer.Start();
            _thread.Start();
        }

        private void Timer_Tick(object sender, EventArgs e)
        {
            while (_requests.TryDequeue(out var request))
            {
                try
                {
                    request.Success = PowerPointCommands.Execute(request.Command, false);
                }
                catch
                {
                    request.Success = false;
                }
                finally
                {
                    request.Completed.Set();
                }
            }
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
                        4,
                        PipeTransmissionMode.Byte,
                        PipeOptions.None))
                    {
                        server.WaitForConnection();
                        if (_cts.IsCancellationRequested) return;

                        using (var reader = new StreamReader(server))
                        using (var writer = new StreamWriter(server) { AutoFlush = true })
                        {
                            var command = reader.ReadLine();
                            if (string.IsNullOrWhiteSpace(command))
                            {
                                writer.WriteLine("ERR");
                                continue;
                            }

                            using (var request = new Request(command.Trim()))
                            {
                                _requests.Enqueue(request);

                                if (request.Completed.Wait(1500))
                                    writer.WriteLine(request.Success ? "OK" : "ERR");
                                else
                                    writer.WriteLine("TIMEOUT");
                            }
                        }
                    }
                }
                catch
                {
                    if (_cts.IsCancellationRequested) return;
                    Thread.Sleep(80);
                }
            }
        }

        public void Dispose()
        {
            _timer.Stop();
            _timer.Dispose();
            _cts.Cancel();

            try
            {
                using (var client = new NamedPipeClientStream(".", PipeName, PipeDirection.Out))
                    client.Connect(100);
            }
            catch { }

            try
            {
                if (_thread.IsAlive) _thread.Join(500);
            }
            catch { }

            while (_requests.TryDequeue(out var request))
            {
                request.Success = false;
                request.Completed.Set();
            }

            _cts.Dispose();
        }
    }
}
