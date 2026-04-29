using System;
using System.IO;
using System.Runtime.Serialization.Json;

namespace PanaceaIEWrapper.Bot
{
    internal static class BotConfigLoader
    {
        public static BotDefinition Load(string configPath)
        {
            if (!File.Exists(configPath))
            {
                throw new FileNotFoundException("No se encontro el archivo de configuracion del bot.", configPath);
            }

            using (FileStream stream = File.OpenRead(configPath))
            {
                var settings = new DataContractJsonSerializerSettings
                {
                    UseSimpleDictionaryFormat = true
                };
                var serializer = new DataContractJsonSerializer(typeof(BotDefinition), settings);
                var definition = serializer.ReadObject(stream) as BotDefinition;
                if (definition == null)
                {
                    throw new InvalidOperationException("No se pudo interpretar la configuracion JSON del bot.");
                }

                if (string.IsNullOrWhiteSpace(definition.BaseUrl))
                {
                    throw new InvalidOperationException("La propiedad baseUrl es obligatoria en la configuracion del bot.");
                }

                if (definition.Steps == null || definition.Steps.Count == 0)
                {
                    throw new InvalidOperationException("La configuracion del bot debe contener al menos un paso en steps.");
                }

                return definition;
            }
        }
    }
}
