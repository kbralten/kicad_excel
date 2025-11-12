using System;
using System.IO;
using System.Net;
using System.Threading.Tasks;

namespace KiCadExcelBridge
{
    public class HttpServer
    {
        private readonly HttpListener _listener;
        private readonly Func<HttpListenerContext, Task> _handler;

        public HttpServer(string prefix, Func<HttpListenerContext, Task> handler)
        {
            _listener = new HttpListener();
            _listener.Prefixes.Add(prefix);
            _handler = handler;
        }

        public async Task Start()
        {
            _listener.Start();
            while (_listener.IsListening)
            {
                try
                {
                    var context = await _listener.GetContextAsync();
                    await _handler(context);
                }
                catch (HttpListenerException)
                {
                    // Listener stopped.
                }
            }
        }

        public void Stop()
        {
            _listener.Stop();
        }
    }
}
