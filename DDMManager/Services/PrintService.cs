using System.Management;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Word;

namespace DDMManager.Services;

/// <summary>
/// Impressão direta de DOCX via Word sem abrir janela.
/// </summary>
public class PrintService
{
    public List<(string Arquivo, bool Sucesso, string Erro)> ImprimirDocxs(
        IEnumerable<string> docxPaths,
        string? impressora = null)
    {
        var resultados = new List<(string, bool, string)>();
        Application? word = null;

        try
        {
            word = new Application { Visible = false };
            word.DisplayAlerts = WdAlertLevel.wdAlertsNone;

            var impressoraOriginal = word.ActivePrinter;
            if (!string.IsNullOrWhiteSpace(impressora))
                word.ActivePrinter = impressora;

            foreach (var path in docxPaths)
            {
                Document? doc = null;
                try
                {
                    doc = word.Documents.Open(
                        FileName: path, ReadOnly: true, AddToRecentFiles: false);
                    doc.PrintOut(Background: false);
                    resultados.Add((path, true, ""));
                }
                catch (Exception ex)
                {
                    resultados.Add((path, false, ex.Message));
                }
                finally
                {
                    if (doc != null)
                    {
                        try { doc.Close(SaveChanges: false); } catch { }
                        Marshal.ReleaseComObject(doc);
                    }
                }
            }

            if (!string.IsNullOrWhiteSpace(impressora))
                word.ActivePrinter = impressoraOriginal;
        }
        finally
        {
            if (word != null)
            {
                try { word.Quit(); } catch { }
                Marshal.ReleaseComObject(word);
            }
        }

        return resultados;
    }

    /// <summary>Lista impressoras via WMI — sem dependência de Windows.Forms.</summary>
    public static IEnumerable<string> ListarImpressoras()
    {
        var impressoras = new List<string>();
        try
        {
            using var searcher = new ManagementObjectSearcher(
                "SELECT Name FROM Win32_Printer");
            foreach (var obj in searcher.Get())
                impressoras.Add(obj["Name"]?.ToString() ?? "");
        }
        catch
        {
            // Fallback: retorna lista vazia se WMI não disponível
        }
        return impressoras.Where(p => !string.IsNullOrWhiteSpace(p)).OrderBy(p => p);
    }
}
