// -----------------------------------------------------------------------------
//  ChromeBridgeServer.cs – lightweight long‑poll bridge for the Chrome extension
//  (camel‑case JSON throughout – no more “Action” / “ReqId” mismatch)
// -----------------------------------------------------------------------------
using System;
using System.Collections.Concurrent;
using System.IO;
using System.Net;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;

internal static class ChromeBridgeServer
{
    // -------------------------------------------------------------------------
    //  Shared JSON options (camelCase response + element serialization)
    // -------------------------------------------------------------------------
    private static readonly JsonSerializerOptions JsonLc = new()
    {
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        DictionaryKeyPolicy  = JsonNamingPolicy.CamelCase
    };

    // -------------------------------------------------------------------------
    //  Internal data structures
    // -------------------------------------------------------------------------
    private static readonly HttpListener _http = new();
    private static readonly ConcurrentQueue<Command> _queue = new();
    private static readonly ConcurrentDictionary<Guid, TaskCompletionSource<JsonElement>> _waits = new();

    private record Command(Guid ReqId, string Action, JsonElement Args);

    // -------------------------------------------------------------------------
    //  Lifecycle
    // -------------------------------------------------------------------------
    public static void Start()
    {
        if (_http.IsListening) return;

        _http.Prefixes.Add("http://127.0.0.1:9234/");
        _http.Prefixes.Add("http://localhost:9234/");
        _http.Start();
        _ = Task.Run(Loop);
        Console.WriteLine("Chrome bridge listening ➜  http://127.0.0.1:9234/");
    }

    public static void Stop()
    {
        if (_http.IsListening) _http.Stop();
    }

    // -------------------------------------------------------------------------
    //  Public API – called by Program.cs
    // -------------------------------------------------------------------------
    public static async Task<JsonElement> RequestAsync(string action,
                                                       object? args      = null,
                                                       int     timeoutMs = 8_000)
    {
        var tcs  = new TaskCompletionSource<JsonElement>(
                       TaskCreationOptions.RunContinuationsAsynchronously);
        var elem = JsonSerializer.SerializeToElement(args ?? new { }, JsonLc);
        var cmd  = new Command(Guid.NewGuid(), action, elem);

        _waits[cmd.ReqId] = tcs;
        _queue.Enqueue(cmd);

        using var cts = new CancellationTokenSource(timeoutMs);
        await using var _ = cts.Token.Register(() => tcs.TrySetCanceled(),
                                               useSynchronizationContext: false);
        return await tcs.Task;        // may throw TaskCanceledException on timeout
    }

    // -------------------------------------------------------------------------
    //  HttpListener loop
    // -------------------------------------------------------------------------
    private static async Task Loop()
    {
        while (_http.IsListening)
        {
            try
            {
                var ctx = await _http.GetContextAsync(); // blocks
                _ = Task.Run(() => Handle(ctx));
            }
            catch { /* listener stopped */ }
        }
    }

    private static async Task Handle(HttpListenerContext ctx)
    {
        try
        {
            string path = ctx.Request.Url!.AbsolutePath.ToLowerInvariant();

            switch (path)
            {
                // -------------------------------------------------------------
                //  Called by the extension (poll for next command)
                // -------------------------------------------------------------
                case "/pending":
                    if (_queue.TryDequeue(out var cmd))
                        await RespondJson(ctx, new
                        {
                            action = cmd.Action,
                            reqId  = cmd.ReqId,
                            args   = cmd.Args
                        });
                    else
                        await RespondJson(ctx, new { }); // empty → nothing queued
                    break;

                // -------------------------------------------------------------
                //  Called by the extension (deliver result)
                // -------------------------------------------------------------
                case "/deliver" when ctx.Request.HttpMethod == "POST":
                {
                    var data = await JsonSerializer.DeserializeAsync<JsonElement>(ctx.Request.InputStream, JsonLc);
                    Guid id  = data.GetProperty("reqId").GetGuid();
                    if (_waits.TryRemove(id, out var tcs))
                        tcs.TrySetResult(data.GetProperty("data"));
                    ctx.Response.StatusCode = 204;
                    break;
                }

                // -------------------------------------------------------------
                //  Convenience proxy endpoints (optional)
                // -------------------------------------------------------------
                case "/active-tabs":
                    await Proxy("getTabs", null, ctx);
                    break;

                case "/activate" when ctx.Request.HttpMethod == "POST":
                {
                    var payload = await JsonSerializer.DeserializeAsync<JsonElement>(ctx.Request.InputStream, JsonLc);
                    await Proxy("activate", payload, ctx);
                    break;
                }

                default:
                    ctx.Response.StatusCode = 404;
                    break;
            }
        }
        catch (Exception ex)
        {
            await RespondText(ctx, 500, ex.ToString());
        }
        finally
        {
            ctx.Response.Close();
        }
    }

    // -------------------------------------------------------------------------
    //  Helpers
    // -------------------------------------------------------------------------
    private static async Task Proxy(string action, object? args, HttpListenerContext ctx)
    {
        var res = await RequestAsync(action, args);
        await RespondJson(ctx, res);
    }

    private static async Task RespondJson(HttpListenerContext ctx, object obj)
    {
        ctx.Response.ContentType = "application/json";
        await JsonSerializer.SerializeAsync(ctx.Response.OutputStream, obj, JsonLc);
    }

    private static async Task RespondText(HttpListenerContext ctx, int code, string txt)
    {
        ctx.Response.StatusCode = code;
        await using var w = new StreamWriter(ctx.Response.OutputStream);
        await w.WriteAsync(txt);
    }
}
