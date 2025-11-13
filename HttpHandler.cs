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
                        // Decode URL-encoded IDs so spaces and other characters are handled
                        var categoryId = WebUtility.UrlDecode(rel[2].Replace(".json", "").Trim('/'));
                        await HandleGetPartsForCategoryAsync(context, categoryId);
                    }
                    else if (rel.Count == 2 && rel[0].Equals("parts/", StringComparison.OrdinalIgnoreCase) && request.HttpMethod == "GET")
                    {
                        var partId = WebUtility.UrlDecode(rel[1].Replace(".json", "").Trim('/'));
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
                try
                {
                    // Log exception details to the http_requests.log for debugging
                    var basePath = AppDomain.CurrentDomain.BaseDirectory ?? ".";
                    var logPath = System.IO.Path.Combine(basePath, "http_requests.log");
                    var segs = request?.Url?.Segments != null ? string.Join("|", request.Url.Segments) : string.Empty;
                    var exLine = $"[{DateTime.UtcNow:u}] ERROR handling request {request?.HttpMethod} {request?.Url?.AbsoluteUri} Segments: {segs} Exception: {ex}\r\n";
                    System.IO.File.AppendAllText(logPath, exLine);
                }
                catch
                {
                    // ignore logging failures
                }

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

        private void LogAndWriteResponse(HttpListenerContext context, string json, int statusCode = (int)HttpStatusCode.OK, string contentType = "application/json")
        {
            var response = context.Response;
            try
            {
                var buffer = Encoding.UTF8.GetBytes(json ?? string.Empty);
                response.ContentType = contentType;
                response.ContentLength64 = buffer.Length;
                response.StatusCode = statusCode;
                // write synchronously to ensure content is sent before logging
                response.OutputStream.Write(buffer, 0, buffer.Length);

                try
                {
                    var basePath = AppDomain.CurrentDomain.BaseDirectory ?? ".";
                    var logPath = System.IO.Path.Combine(basePath, "http_requests.log");
                    var req = context.Request;
                    var segs = req?.Url?.Segments != null ? string.Join("|", req.Url.Segments) : string.Empty;
                    var bodyPreview = string.IsNullOrEmpty(json) ? "<empty>" : (json.Length > 2000 ? json.Substring(0, 2000) + "...(truncated)" : json);
                    var line = $"[{DateTime.UtcNow:u}] Response: {req?.HttpMethod} {req?.Url?.AbsoluteUri} Status: {statusCode} ContentType: {contentType} BodyLen: {buffer.Length}\r\n{bodyPreview}\r\n";
                    System.IO.File.AppendAllText(logPath, line);
                }
                catch
                {
                    // ignore logging failures
                }
            }
            catch
            {
                // If writing fails, attempt to at least set status code
                try { response.StatusCode = statusCode; } catch { }
            }
        }

        private Task HandleEndpointValidationAsync(HttpListenerContext context)
        {
            // Return empty strings to align with expected KiCad API contract
            var data = new { categories = string.Empty, parts = string.Empty };
            var json = JsonSerializer.Serialize(data);
            LogAndWriteResponse(context, json, (int)HttpStatusCode.OK);
            return Task.CompletedTask;
        }

        private Task HandleGetCategoriesAsync(HttpListenerContext context)
        {
            var categories = _excelManager.GetCategories();
            var json = JsonSerializer.Serialize(categories);
            LogAndWriteResponse(context, json, (int)HttpStatusCode.OK);
            return Task.CompletedTask;
        }

        private Task HandleGetPartsForCategoryAsync(HttpListenerContext context, string categoryId)
        {
            var parts = _excelManager.GetPartsForCategory(categoryId);
            var json = JsonSerializer.Serialize(parts);
            LogAndWriteResponse(context, json, (int)HttpStatusCode.OK);
            return Task.CompletedTask;
        }

        private Task HandleGetPartDetailsAsync(HttpListenerContext context, string partId)
        {
            var partDetails = _excelManager.GetPartDetails(partId);
            if (partDetails == null)
            {
                LogAndWriteResponse(context, "", (int)HttpStatusCode.NotFound);
                return Task.CompletedTask;
            }

            var json = JsonSerializer.Serialize(partDetails);
            LogAndWriteResponse(context, json, (int)HttpStatusCode.OK);
            return Task.CompletedTask;
        }
    }
}
