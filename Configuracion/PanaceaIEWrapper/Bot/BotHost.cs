using System;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;

namespace PanaceaIEWrapper.Bot
{
    internal static class BotHost
    {
        public static async Task<int> RunAsync(string[] args)
        {
            string baseDir = AppDomain.CurrentDomain.BaseDirectory;
            string configPath = GetConfigPath(args, baseDir);

            string logDir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Panacea", "bot");
            string logPath = Path.Combine(logDir, $"bot-{DateTime.Now:yyyyMMdd-HHmmss}.log");
            var logger = new BotLogger(logPath);

            logger.Info("Inicio de ejecucion del bot.");
            logger.Info("Archivo de configuracion: " + configPath);

            try
            {
                BotDefinition definition = BotConfigLoader.Load(configPath);
                var context = new BotContext();
                foreach (var pair in definition.Variables)
                {
                    context.Variables[pair.Key] = pair.Value ?? string.Empty;
                }

                string userFromEnv = Environment.GetEnvironmentVariable("PANACEA_USER");
                if (!string.IsNullOrWhiteSpace(userFromEnv))
                {
                    context.Variables["username"] = userFromEnv;
                }

                string passFromEnv = Environment.GetEnvironmentVariable("PANACEA_PASS");
                if (!string.IsNullOrWhiteSpace(passFromEnv))
                {
                    context.Variables["password"] = passFromEnv;
                }

                var runner = new BotHttpRunner(definition, context, logger);
                await runner.RunAsync(CancellationToken.None).ConfigureAwait(false);
                logger.Info("Bot finalizado correctamente.");
                return 0;
            }
            catch (Exception ex)
            {
                logger.Error("Fallo en la ejecucion del bot: " + ex);
                return 1;
            }
        }

        private static string GetConfigPath(string[] args, string baseDir)
        {
            string arg = args?.FirstOrDefault(a => a.StartsWith("--bot-config=", StringComparison.OrdinalIgnoreCase));
            if (string.IsNullOrWhiteSpace(arg))
            {
                return Path.Combine(baseDir, "bot-config.json");
            }

            string value = arg.Substring("--bot-config=".Length).Trim('"');
            if (Path.IsPathRooted(value))
            {
                return value;
            }

            return Path.Combine(baseDir, value);
        }
    }
}
