using System.Text.RegularExpressions;
using System;

namespace PanaceaIEWrapper.Bot
{
    internal static class TemplateEngine
    {
        private static readonly Regex TokenRegex = new Regex(@"\{\{(?<name>[^\}]+)\}\}", RegexOptions.Compiled);

        public static string Resolve(string template, BotContext context)
        {
            if (string.IsNullOrEmpty(template))
            {
                return template;
            }

            return TokenRegex.Replace(template, match =>
            {
                string key = match.Groups["name"].Value.Trim();
                if (context.Variables.TryGetValue(key, out string value))
                {
                    return value;
                }

                return string.Empty;
            });
        }

        public static string ResolveUrlEncoded(string template, BotContext context)
        {
            if (string.IsNullOrEmpty(template))
            {
                return template;
            }

            return TokenRegex.Replace(template, match =>
            {
                string key = match.Groups["name"].Value.Trim();
                if (context.Variables.TryGetValue(key, out string value))
                {
                    return Uri.EscapeDataString(value ?? string.Empty);
                }

                return string.Empty;
            });
        }
    }
}
