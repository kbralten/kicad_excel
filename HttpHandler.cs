using System;
using System.IO;
using System.Net;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;
using System.Collections.Generic;

namespace KiCadExcelBridge
{
    public class HttpHandler
    {
        private const string ApiVersion = "v1";
        private readonly ExcelManager _excelManager;

        public HttpHandler(ExcelManager excelManager)
        {
            _excelManager = excelManager;
        }

        public async Task HandleRequestAsync(HttpListenerContext context)
        {
            var request = context.Request;
            var response = context.Response;
            var url = request.Url;

            // Log incoming request for troubleshooting (best-effort)
            try
            {
                var basePath = AppDomain.CurrentDomain.BaseDirectory ?? ".";
                var logPath = System.IO.Path.Combine(basePath, "http_requests.log");
                var segs = url?.Segments != null ? string.Join("|", url.Segments) : string.Empty;
                var line = $"[{DateTime.UtcNow:u}] Request: {url?.AbsoluteUri} Path: {url?.AbsolutePath} Segments: {segs}\r\n";
                System.IO.File.AppendAllText(logPath, line);
            }
            catch
            {
                // ignore logging failures
            }

            if (url == null)
            {
                response.StatusCode = (int)HttpStatusCode.BadRequest;
                response.Close();
                return;
            }

            var segments = url.Segments;

            try
            {
                // Find the position of the API version segment (e.g. "v1") so the handler works
                // regardless of the HttpListener prefix (which may include extra path segments).
                int baseIndex = -1;
                for (int i = 0; i < segments.Length; i++)
                {
                    if (string.Equals(segments[i].Trim('/'), ApiVersion, StringComparison.OrdinalIgnoreCase))
                    {
                        baseIndex = i;
                        break;
                    }
                }

                if (baseIndex == -1)
                {
                    response.StatusCode = (int)HttpStatusCode.NotFound;
                }
                else
                {
                    // Build relative segments after the version
                    var rel = new List<string>();
                    for (int i = baseIndex + 1; i < segments.Length; i++) rel.Add(segments[i]);

                    if (rel.Count == 0 && request.HttpMethod == "GET")
                    {
                        await HandleEndpointValidationAsync(context);
                    }
                    else if (rel.Count == 1 && rel[0].Equals("categories.json", StringComparison.OrdinalIgnoreCase) && request.HttpMethod == "GET")
                    {
                        await HandleGetCategoriesAsync(context);
                    }
                    else if (rel.Count == 3 && rel[0].Equals("parts/", StringComparison.OrdinalIgnoreCase) && rel[1].Equals("category/", StringComparison.OrdinalIgnoreCase) && request.HttpMethod == "GET")
                    {
                        var categoryId = rel[2].Replace(".json", "").Trim('/');
                        await HandleGetPartsForCategoryAsync(context, categoryId);
                    }
                    else if (rel.Count == 2 && rel[0].Equals("parts/", StringComparison.OrdinalIgnoreCase) && request.HttpMethod == "GET")
                    {
                        var partId = rel[1].Replace(".json", "").Trim('/');
                        await HandleGetPartDetailsAsync(context, partId);
                    }
                    else
                    {
                        response.StatusCode = (int)HttpStatusCode.NotFound;
                    }
                }
            }
            catch (Exception ex)
            {
                response.StatusCode = (int)HttpStatusCode.InternalServerError;
                var buffer = Encoding.UTF8.GetBytes(ex.Message);
                response.ContentLength64 = buffer.Length;
                response.OutputStream.Write(buffer, 0, buffer.Length);
            }
            finally
            {
                response.Close();
            }
        }

        private async Task HandleEndpointValidationAsync(HttpListenerContext context)
        {
            var response = context.Response;
            // Return empty strings to align with expected KiCad API contract
            var data = new { categories = string.Empty, parts = string.Empty };
            var json = JsonSerializer.Serialize(data);
            var buffer = Encoding.UTF8.GetBytes(json);

            response.ContentType = "application/json";
            response.ContentLength64 = buffer.Length;
            await response.OutputStream.WriteAsync(buffer, 0, buffer.Length);
            response.StatusCode = (int)HttpStatusCode.OK;
        }

        private async Task HandleGetCategoriesAsync(HttpListenerContext context)
        {
            var categories = _excelManager.GetCategories();
            var json = JsonSerializer.Serialize(categories);
            var buffer = Encoding.UTF8.GetBytes(json);
            var response = context.Response;

            response.ContentType = "application/json";
            response.ContentLength64 = buffer.Length;
            await response.OutputStream.WriteAsync(buffer, 0, buffer.Length);
            response.StatusCode = (int)HttpStatusCode.OK;
        }

        private async Task HandleGetPartsForCategoryAsync(HttpListenerContext context, string categoryId)
        {
            var parts = _excelManager.GetPartsForCategory(categoryId);
            var json = JsonSerializer.Serialize(parts);
            var buffer = Encoding.UTF8.GetBytes(json);
            var response = context.Response;

            response.ContentType = "application/json";
            response.ContentLength64 = buffer.Length;
            await response.OutputStream.WriteAsync(buffer, 0, buffer.Length);
            response.StatusCode = (int)HttpStatusCode.OK;
        }

        private async Task HandleGetPartDetailsAsync(HttpListenerContext context, string partId)
        {
            var partDetails = _excelManager.GetPartDetails(partId);
            if (partDetails == null)
            {
                context.Response.StatusCode = (int)HttpStatusCode.NotFound;
                return;
            }

            var json = JsonSerializer.Serialize(partDetails);
            var buffer = Encoding.UTF8.GetBytes(json);
            var response = context.Response;

            response.ContentType = "application/json";
            response.ContentLength64 = buffer.Length;
            await response.OutputStream.WriteAsync(buffer, 0, buffer.Length);
            response.StatusCode = (int)HttpStatusCode.OK;
        }
    }
}
