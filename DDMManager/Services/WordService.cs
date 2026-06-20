using System.IO;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Word;

namespace DDMManager.Services;

/// <summary>
/// Conversão DOCX → PDF e mesclagem de DOCX via Microsoft Word (COM Interop).
/// Abre o Word uma única vez para todos os arquivos do lote.
/// </summary>
public class WordService
{
    /// <summary>
    /// Converte uma lista de arquivos DOCX em PDF.
    /// Retorna lista de (pdfPath | null, erro | "") para cada entrada.
    /// </summary>
    public List<(string? Pdf, string Erro)> ConverterParaPdf(IEnumerable<string> docxPaths)
    {
        var resultados = new List<(string?, string)>();
        Application? word = null;

        try
        {
            word = new Application { Visible = false };
            word.DisplayAlerts = WdAlertLevel.wdAlertsNone;

            foreach (var docxPath in docxPaths)
            {
                var pdfPath = Path.ChangeExtension(docxPath, ".pdf");
                Document? doc = null;
                try
                {
                    doc = word.Documents.Open(
                        FileName:           docxPath,
                        ReadOnly:           true,
                        AddToRecentFiles:   false);

                    doc.SaveAs2(
                        FileName:   pdfPath,
                        FileFormat: WdSaveFormat.wdFormatPDF);

                    resultados.Add((pdfPath, ""));
                }
                catch (Exception ex)
                {
                    resultados.Add((null, ex.Message));
                }
                finally
                {
                    if (doc != null)
                    {
                        try { doc.Close(SaveChanges: false); }
                        catch { /* ignora */ }
                        Marshal.ReleaseComObject(doc);
                    }
                }
            }
        }
        finally
        {
            if (word != null)
            {
                try { word.Quit(); }
                catch { /* ignora */ }
                Marshal.ReleaseComObject(word);
            }
        }

        return resultados;
    }

    /// <summary>
    /// Combina vários DOCX em um único arquivo usando InsertFile do Word.
    /// Cada documento é separado por uma quebra de página.
    /// </summary>
    public string CombinarDocx(IList<string> docxPaths, string pastaDestino)
    {
        if (docxPaths.Count == 0)
            throw new ArgumentException("Nenhum arquivo para combinar.");

        var saida   = Path.Combine(pastaDestino, "DDM_SEMANA_COMBINADO.docx");
        Application? word = null;
        Document?    docBase = null;

        try
        {
            word = new Application { Visible = false };
            word.DisplayAlerts = WdAlertLevel.wdAlertsNone;

            // Abre o primeiro doc como base
            docBase = word.Documents.Open(
                FileName:         docxPaths[0],
                ReadOnly:         false,
                AddToRecentFiles: false);

            foreach (var path in docxPaths.Skip(1))
            {
                // Move cursor para o final e insere quebra de página
                var rng = docBase.Range();
                rng.Collapse(WdCollapseDirection.wdCollapseEnd);
                rng.InsertBreak(WdBreakType.wdPageBreak);

                // Move para o final novamente e insere o arquivo
                rng = docBase.Range();
                rng.Collapse(WdCollapseDirection.wdCollapseEnd);
                rng.InsertFile(path);
            }

            docBase.SaveAs2(saida);
            return saida;
        }
        finally
        {
            if (docBase != null)
            {
                try { docBase.Close(SaveChanges: false); }
                catch { }
                Marshal.ReleaseComObject(docBase);
            }
            if (word != null)
            {
                try { word.Quit(); }
                catch { }
                Marshal.ReleaseComObject(word);
            }
        }
    }

    /// <summary>
    /// Mescla PDFs em um único arquivo usando o Word como conversor intermediário.
    /// Estratégia: combina os DOCX e converte o resultado para PDF.
    /// </summary>
    public string MesclarPdfs(IList<string> pdfPaths, IList<string> docxOrigens,
                               string pastaDestino)
    {
        // Combina os DOCX → converte para PDF
        var docxCombinado = CombinarDocx(docxOrigens, pastaDestino);
        var pdfSaida      = Path.Combine(pastaDestino, "DDM_SEMANA_COMBINADO.pdf");

        var resultado = ConverterParaPdf(new[] { docxCombinado });
        if (resultado[0].Pdf == null)
            throw new InvalidOperationException(resultado[0].Erro);

        // Renomeia para o nome correto
        if (File.Exists(pdfSaida)) File.Delete(pdfSaida);
        File.Move(resultado[0].Pdf!, pdfSaida);
        return pdfSaida;
    }
}
