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
        private readonly ConcurrentQueue<string> _commands = new ConcurrentQueue<string>();
        private readonly CancellationTokenSource _cts = new CancellationTokenSource();
        private readonly Thread _thread;
        private readonly Timer _timer;

        public PipeCommandServer()
        {
            _timer = new Timer { Interval = 40 };
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
            while (_commands.TryDequeue(out var command))
                PowerPointCommands.Execute(command, false);
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
                        PipeOptions.None))
                    {
                        server.WaitForConnection();
                        if (_cts.IsCancellationRequested) return;

                        using (var reader = new StreamReader(server))
                        using (var writer = new StreamWriter(server) { AutoFlush = true })
                        {
                            var command = reader.ReadLine();
                            if (!string.IsNullOrWhiteSpace(command))
                            {
                                _commands.Enqueue(command.Trim());
                                writer.WriteLine("OK");
                            }
                            else
                            {
                                writer.WriteLine("ERR");
                            }
                        }
                    }
                }
                catch
                {
                    if (_cts.IsCancellationRequested) return;
                    Thread.Sleep(100);
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
                {
                    client.Connect(100);
                }
            }
            catch { }

            try
            {
                if (_thread.IsAlive) _thread.Join(500);
            }
            catch { }

            _cts.Dispose();
        }
    }
}
