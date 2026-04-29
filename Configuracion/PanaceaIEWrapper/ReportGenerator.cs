using System;
using System.Drawing;
using System.IO;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace PanaceaIEWrapper
{
    internal static class ReportGenerator
    {
        public static string Generate(ProgressState state)
        {
            string reportPath = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory),
                "INFORME_PANACEA_" + DateTime.Now.ToString("yyyyMMdd_HHmmss") + ".xlsx");

            using (var pkg = new ExcelPackage())
            {
                var ws = pkg.Workbook.Worksheets.Add("Informe RIPS");

                // Titulo
                ws.Cells[1, 1].Value = "INFORME DE PROCESAMIENTO — PANACEA RIPS";
                ws.Cells[1, 1, 1, 7].Merge = true;
                ws.Cells[1, 1].Style.Font.Bold = true;
                ws.Cells[1, 1].Style.Font.Size = 14;
                ws.Cells[1, 1].Style.Font.Color.SetColor(Color.White);
                ws.Cells[1, 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                ws.Cells[1, 1].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(0, 84, 166));
                ws.Cells[1, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                ws.Row(1).Height = 28;

                // Resumen
                WriteKV(ws, 3, "Generado el:", DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss"));
                WriteKV(ws, 4, "Equipo:", string.IsNullOrWhiteSpace(state.MachineName) ? Environment.MachineName : state.MachineName);
                WriteKV(ws, 5, "Archivo base:", string.IsNullOrWhiteSpace(state.ExcelPath) ? "—" : state.ExcelPath);
                WriteKV(ws, 6, "Total registros:", state.TotalRecords.ToString());
                int done = state.ProcessedRecords?.Count ?? 0;
                WriteKV(ws, 7, "Procesados:", done.ToString());
                WriteKV(ws, 8, "Pendientes:", (state.TotalRecords - done).ToString());

                // Cabeceras tabla
                string[] headers = { "#", "Cedula (CC)", "Fecha Inicio", "Fecha Fin", "Convenio", "Estado", "Fecha/Hora" };
                for (int c = 0; c < headers.Length; c++)
                {
                    ws.Cells[10, c + 1].Value = headers[c];
                    ws.Cells[10, c + 1].Style.Font.Bold = true;
                    ws.Cells[10, c + 1].Style.Font.Color.SetColor(Color.White);
                    ws.Cells[10, c + 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    ws.Cells[10, c + 1].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(0, 84, 166));
                    ws.Cells[10, c + 1].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                }

                // Datos
                int row = 11;
                if (state.ProcessedRecords != null)
                {
                    int n = 1;
                    foreach (var r in state.ProcessedRecords)
                    {
                        ws.Cells[row, 1].Value = n++;
                        ws.Cells[row, 2].Value = r.CC;
                        ws.Cells[row, 3].Value = r.FechaInicio;
                        ws.Cells[row, 4].Value = r.FechaFin;
                        ws.Cells[row, 5].Value = r.Convenio;
                        ws.Cells[row, 6].Value = r.Estado;
                        ws.Cells[row, 7].Value = r.Timestamp;

                        Color bg = r.Estado == "Completado" ? Color.FromArgb(198, 239, 206)
                                 : r.Estado == "Error"      ? Color.FromArgb(255, 199, 206)
                                 :                            Color.FromArgb(255, 235, 156);
                        for (int c = 1; c <= 7; c++)
                        {
                            ws.Cells[row, c].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            ws.Cells[row, c].Style.Fill.BackgroundColor.SetColor(bg);
                            ws.Cells[row, c].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                        }
                        row++;
                    }
                }

                if (ws.Dimension != null)
                    ws.Cells[ws.Dimension.Address].AutoFitColumns();

                pkg.SaveAs(new FileInfo(reportPath));
            }

            return reportPath;
        }

        private static void WriteKV(ExcelWorksheet ws, int row, string key, string value)
        {
            ws.Cells[row, 1].Value = key;
            ws.Cells[row, 1].Style.Font.Bold = true;
            ws.Cells[row, 2].Value = value;
        }
    }
}
