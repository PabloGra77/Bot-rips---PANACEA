using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Json;

namespace PanaceaIEWrapper
{
    [DataContract]
    internal sealed class ProgressState
    {
        [DataMember(Name = "excelPath")]
        public string ExcelPath { get; set; } = string.Empty;

        [DataMember(Name = "totalRecords")]
        public int TotalRecords { get; set; }

        [DataMember(Name = "currentIndex")]
        public int CurrentIndex { get; set; }

        [DataMember(Name = "processedRecords")]
        public List<ProcessedRecord> ProcessedRecords { get; set; } = new List<ProcessedRecord>();

        [DataMember(Name = "lastUpdate")]
        public string LastUpdate { get; set; } = string.Empty;

        [DataMember(Name = "machineName")]
        public string MachineName { get; set; } = string.Empty;

        private static string StatePath =>
            Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "PanaceaRIPS", "progress.json");

        public static ProgressState Load()
        {
            try
            {
                if (!File.Exists(StatePath)) return new ProgressState();
                using (var fs = File.OpenRead(StatePath))
                {
                    var s = new DataContractJsonSerializer(typeof(ProgressState),
                        new DataContractJsonSerializerSettings { UseSimpleDictionaryFormat = true });
                    return (s.ReadObject(fs) as ProgressState) ?? new ProgressState();
                }
            }
            catch { return new ProgressState(); }
        }

        public void Save()
        {
            try
            {
                string dir = Path.GetDirectoryName(StatePath);
                if (!string.IsNullOrWhiteSpace(dir)) Directory.CreateDirectory(dir);
                LastUpdate = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                MachineName = Environment.MachineName;
                using (var fs = File.Open(StatePath, FileMode.Create, FileAccess.Write))
                {
                    var s = new DataContractJsonSerializer(typeof(ProgressState),
                        new DataContractJsonSerializerSettings { UseSimpleDictionaryFormat = true });
                    s.WriteObject(fs, this);
                }
            }
            catch { }
        }

        public static void Clear()
        {
            try { if (File.Exists(StatePath)) File.Delete(StatePath); } catch { }
        }
    }

    [DataContract]
    internal sealed class ProcessedRecord
    {
        [DataMember(Name = "cc")]
        public string CC { get; set; } = string.Empty;

        [DataMember(Name = "fechaInicio")]
        public string FechaInicio { get; set; } = string.Empty;

        [DataMember(Name = "fechaFin")]
        public string FechaFin { get; set; } = string.Empty;

        [DataMember(Name = "convenio")]
        public string Convenio { get; set; } = string.Empty;

        [DataMember(Name = "estado")]
        public string Estado { get; set; } = string.Empty;

        [DataMember(Name = "timestamp")]
        public string Timestamp { get; set; } = string.Empty;
    }
}
