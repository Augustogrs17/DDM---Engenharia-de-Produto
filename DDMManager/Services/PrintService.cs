using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Word;

namespace DDMManager.Services;

/// <summary>
/// Impressão direta de DOCX via Word sem abrir janela.
/// </summary>
public class PrintService
{
    /// <summary>
    /// Imprime uma lista de DOCX na impressora padrão sem abrir o Word visualmente.
    /// </summary>
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

            // Define impressora se especificada
            var impressoraOriginal = word.ActivePrinter;
            if (!string.IsNullOrWhiteSpace(impressora))
                word.ActivePrinter = impressora;

            foreach (var path in docxPaths)
            {
                Document? doc = null;
                try
                {
                    doc = word.Documents.Open(
                        FileName:         path,
                        ReadOnly:         true,
                        AddToRecentFiles: false);

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

            // Restaura impressora original
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

    /// <summary>Retorna lista de impressoras disponíveis no sistema.</summary>
    public static IEnumerable<string> ListarImpressoras()
        => System.Drawing.Printing.PrinterSettings.InstalledPrinters.Cast<string>();
}
