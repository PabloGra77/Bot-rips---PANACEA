using System;
using System.Collections.Generic;
using System.Runtime.Serialization;

namespace PanaceaIEWrapper.Bot
{
    [DataContract]
    public sealed class BotDefinition
    {
        [DataMember(Name = "baseUrl")]
        public string BaseUrl { get; set; }

        [DataMember(Name = "variables")]
        public Dictionary<string, string> Variables { get; set; } = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        [DataMember(Name = "defaultHeaders")]
        public Dictionary<string, string> DefaultHeaders { get; set; } = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        [DataMember(Name = "steps")]
        public List<BotStepDefinition> Steps { get; set; } = new List<BotStepDefinition>();
    }

    [DataContract]
    public sealed class BotStepDefinition
    {
        [DataMember(Name = "name")]
        public string Name { get; set; }

        [DataMember(Name = "method")]
        public string Method { get; set; } = "GET";

        [DataMember(Name = "path")]
        public string Path { get; set; }

        [DataMember(Name = "headers")]
        public Dictionary<string, string> Headers { get; set; } = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        [DataMember(Name = "bodyTemplate")]
        public string BodyTemplate { get; set; }

        [DataMember(Name = "contentType")]
        public string ContentType { get; set; } = "application/json";

        [DataMember(Name = "saveResponseAs")]
        public string SaveResponseAs { get; set; }

        [DataMember(Name = "extractRegex")]
        public Dictionary<string, string> ExtractRegex { get; set; } = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        [DataMember(Name = "successWhenContains")]
        public List<string> SuccessWhenContains { get; set; } = new List<string>();

        [DataMember(Name = "failWhenContains")]
        public List<string> FailWhenContains { get; set; } = new List<string>();

        [DataMember(Name = "delayMs")]
        public int DelayMs { get; set; }
    }

    public sealed class BotContext
    {
        public BotContext()
        {
            Variables = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        }

        public Dictionary<string, string> Variables { get; }
    }

    public sealed class RipsRecord
    {
        public string CC { get; set; }
        public string FechaInicio { get; set; }  // formato dd-MM-yyyy
        public string FechaFin { get; set; }     // formato dd-MM-yyyy
        public string TipoConvenio { get; set; } // ej: FOMAG-1
        public string Finalidad { get; set; }
        public string Diagnostico { get; set; }
        public string CausaExterna { get; set; }
    }
}
