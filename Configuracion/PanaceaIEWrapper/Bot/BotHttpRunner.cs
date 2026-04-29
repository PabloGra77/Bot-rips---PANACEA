using System;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;

namespace PanaceaIEWrapper.Bot
{
    internal sealed class BotHttpRunner
    {
        private readonly BotDefinition _definition;
        private readonly BotContext _context;
        private readonly BotLogger _logger;

        public BotHttpRunner(BotDefinition definition, BotContext context, BotLogger logger)
        {
            _definition = definition;
            _context = context;
            _logger = logger;
        }

        public async Task RunAsync(CancellationToken cancellationToken)
        {
            var cookieContainer = new CookieContainer();
            var handler = new HttpClientHandler
            {
                UseCookies = true,
                CookieContainer = cookieContainer,
                AutomaticDecompression = DecompressionMethods.GZip | DecompressionMethods.Deflate
            };

            using (handler)
            using (var client = new HttpClient(handler) { BaseAddress = new Uri(_definition.BaseUrl, UriKind.Absolute) })
            {
                foreach (BotStepDefinition step in _definition.Steps)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    await RunStepAsync(client, step, cancellationToken).ConfigureAwait(false);

                    if (step.DelayMs > 0)
                    {
                        await Task.Delay(step.DelayMs, cancellationToken).ConfigureAwait(false);
                    }
                }
            }
        }

        private async Task RunStepAsync(HttpClient client, BotStepDefinition step, CancellationToken cancellationToken)
        {
            if (string.IsNullOrWhiteSpace(step.Path))
            {
                throw new InvalidOperationException($"El paso '{step.Name}' no tiene path.");
            }

            string resolvedPath = TemplateEngine.Resolve(step.Path, _context);
            var targetUri = Uri.TryCreate(resolvedPath, UriKind.Absolute, out Uri absolute)
                ? absolute
                : new Uri(client.BaseAddress, resolvedPath);

            var method = new HttpMethod((step.Method ?? "GET").ToUpperInvariant());
            using (var request = new HttpRequestMessage(method, targetUri))
            {
                AddHeaders(request, _definition.DefaultHeaders);
                AddHeaders(request, step.Headers);

                if (!string.IsNullOrWhiteSpace(step.BodyTemplate) && method != HttpMethod.Get)
                {
                    bool isFormUrlEncoded = !string.IsNullOrWhiteSpace(step.ContentType)
                        && step.ContentType.IndexOf("application/x-www-form-urlencoded", StringComparison.OrdinalIgnoreCase) >= 0;
                    string resolvedBody = isFormUrlEncoded
                        ? TemplateEngine.ResolveUrlEncoded(step.BodyTemplate, _context)
                        : TemplateEngine.Resolve(step.BodyTemplate, _context);
                    request.Content = new StringContent(resolvedBody, Encoding.UTF8, step.ContentType ?? "application/json");
                }

                _logger.Info($"Ejecutando paso '{step.Name}' -> {method} {targetUri}");
                using (HttpResponseMessage response = await client.SendAsync(request, cancellationToken).ConfigureAwait(false))
                {
                    string responseBody = await response.Content.ReadAsStringAsync().ConfigureAwait(false);
                    _logger.Info($"Paso '{step.Name}' estado {(int)response.StatusCode} {response.ReasonPhrase}");

                    if (!response.IsSuccessStatusCode)
                    {
                        throw new HttpRequestException($"Paso '{step.Name}' fallo con codigo {(int)response.StatusCode}. Respuesta: {Truncate(responseBody)}");
                    }

                    ValidateFailureMarkers(step, responseBody);
                    ValidateContains(step, responseBody);
                    SaveResponse(step, responseBody);
                    ExtractVariables(step, responseBody);
                }
            }
        }

        private static void AddHeaders(HttpRequestMessage request, System.Collections.Generic.Dictionary<string, string> headers)
        {
            if (headers == null)
            {
                return;
            }

            foreach (var pair in headers)
            {
                if (string.IsNullOrWhiteSpace(pair.Key) || pair.Value == null)
                {
                    continue;
                }

                request.Headers.Remove(pair.Key);
                request.Headers.TryAddWithoutValidation(pair.Key, pair.Value);
            }
        }

        private void SaveResponse(BotStepDefinition step, string responseBody)
        {
            if (string.IsNullOrWhiteSpace(step.SaveResponseAs))
            {
                return;
            }

            _context.Variables[step.SaveResponseAs] = responseBody ?? string.Empty;
        }

        private void ExtractVariables(BotStepDefinition step, string responseBody)
        {
            if (step.ExtractRegex == null || step.ExtractRegex.Count == 0 || string.IsNullOrEmpty(responseBody))
            {
                return;
            }

            foreach (var pair in step.ExtractRegex)
            {
                var regex = new Regex(pair.Value, RegexOptions.IgnoreCase | RegexOptions.Singleline);
                Match match = regex.Match(responseBody);
                if (!match.Success)
                {
                    continue;
                }

                string value = match.Groups.Count > 1 ? match.Groups[1].Value : match.Value;
                _context.Variables[pair.Key] = WebUtility.HtmlDecode(value ?? string.Empty);
                _logger.Info($"Variable extraida '{pair.Key}'.");
            }
        }

        private static void ValidateContains(BotStepDefinition step, string responseBody)
        {
            if (step.SuccessWhenContains == null || step.SuccessWhenContains.Count == 0)
            {
                return;
            }

            bool allFound = step.SuccessWhenContains.All(token =>
                !string.IsNullOrWhiteSpace(token) && (responseBody?.IndexOf(token, StringComparison.OrdinalIgnoreCase) >= 0));

            if (!allFound)
            {
                throw new InvalidOperationException($"Paso '{step.Name}' no contiene los marcadores de exito esperados.");
            }
        }

        private static void ValidateFailureMarkers(BotStepDefinition step, string responseBody)
        {
            if (step.FailWhenContains == null || step.FailWhenContains.Count == 0 || string.IsNullOrEmpty(responseBody))
            {
                return;
            }

            string matched = step.FailWhenContains.FirstOrDefault(token =>
                !string.IsNullOrWhiteSpace(token) && responseBody.IndexOf(token, StringComparison.OrdinalIgnoreCase) >= 0);

            if (!string.IsNullOrWhiteSpace(matched))
            {
                throw new InvalidOperationException($"Paso '{step.Name}' detecto marcador de fallo: '{matched}'.");
            }
        }

        private static string Truncate(string text)
        {
            if (string.IsNullOrEmpty(text) || text.Length <= 300)
            {
                return text;
            }

            return text.Substring(0, 300) + "...";
        }
    }
}
