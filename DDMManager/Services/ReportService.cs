using System.IO;
using System.IO.Packaging;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace DDMManager.Services;

/// <summary>
/// Gera relatório semanal .xlsx com os DDMs realizados.
/// Usa OpenXml puro — sem dependência extra além do DocumentFormat.OpenXml já presente.
/// </summary>
public class ReportService
{
    public string GerarRelatorioSemanal(
        IEnumerable<(string DiaNome, string Arquivo, string Tema, string Data, int Participantes)> itens,
        string pastaDestino,
        DateTime semanaRef)
    {
        var (seg, sex) = DdmService.SemanaAtual(semanaRef);
        var nomeArq = $"Relatorio_DDM_{seg:dd-MM-yyyy}.xlsx";
        var caminho = Path.Combine(pastaDestino, nomeArq);

        using var doc = SpreadsheetDocument.Create(caminho, SpreadsheetDocumentType.Workbook);

        var wbPart = doc.AddWorkbookPart();
        wbPart.Workbook = new Workbook();

        // Estilos
        var stylesPart = wbPart.AddNewPart<WorkbookStylesPart>();
        stylesPart.Stylesheet = CriarEstilos();

        var wsPart = wbPart.AddNewPart<WorksheetPart>();
        var sheetData = new SheetData();
        wsPart.Worksheet = new Worksheet(sheetData);

        var sheets = doc.WorkbookPart!.Workbook.AppendChild(new Sheets());
        sheets.Append(new Sheet
        {
            Id      = doc.WorkbookPart.GetIdOfPart(wsPart),
            SheetId = 1,
            Name    = "DDMs da Semana"
        });

        // Larguras de coluna
        wsPart.Worksheet.InsertBefore(new Columns(
            new Column { Min = 1, Max = 1, Width = 18, CustomWidth = true },
            new Column { Min = 2, Max = 2, Width = 30, CustomWidth = true },
            new Column { Min = 3, Max = 3, Width = 40, CustomWidth = true },
            new Column { Min = 4, Max = 4, Width = 14, CustomWidth = true },
            new Column { Min = 5, Max = 5, Width = 16, CustomWidth = true }
        ), sheetData);

        // Título
        sheetData.AppendChild(CriarLinha(1,
            new[] { $"Relatório DDM — Semana {seg:dd/MM} a {sex:dd/MM/yyyy}" },
            styleIdx: 2));

        sheetData.AppendChild(CriarLinha(2,
            new[] { $"Gerado em: {DateTime.Now:dd/MM/yyyy HH:mm}" },
            styleIdx: 0));

        sheetData.AppendChild(CriarLinha(3, Array.Empty<string>()));

        // Cabeçalho
        sheetData.AppendChild(CriarLinha(4,
            new[] { "Dia", "Arquivo", "Tema", "Data DDM", "Participantes" },
            styleIdx: 1));

        // Dados
        int rowIdx = 5;
        foreach (var (diaNome, arquivo, tema, data, part) in itens)
        {
            sheetData.AppendChild(CriarLinha(rowIdx++,
                new[] { diaNome, Path.GetFileName(arquivo), tema, data, part.ToString() }));
        }

        // Linha de total
        sheetData.AppendChild(CriarLinha(rowIdx, Array.Empty<string>()));
        sheetData.AppendChild(CriarLinha(rowIdx + 1,
            new[] { "Total de DDMs realizados:", itens.Count().ToString() },
            styleIdx: 1));

        wbPart.Workbook.Save();
        return caminho;
    }

    // ── Helpers OpenXml ───────────────────────────────────────────────────

    private static Row CriarLinha(int rowIndex, string[] valores, uint styleIdx = 0)
    {
        var row = new Row { RowIndex = (uint)rowIndex };
        string[] cols = { "A", "B", "C", "D", "E" };

        for (int i = 0; i < valores.Length; i++)
        {
            var cell = new Cell
            {
                CellReference = $"{cols[i]}{rowIndex}",
                DataType      = CellValues.String,
                CellValue     = new CellValue(valores[i]),
                StyleIndex    = styleIdx,
            };
            row.AppendChild(cell);
        }
        return row;
    }

    private static Stylesheet CriarEstilos()
    {
        // Normal, Cabeçalho (negrito azul), Título (negrito grande)
        return new Stylesheet(
            new Fonts(
                new Font(),   // 0 normal
                new Font(     // 1 negrito
                    new Bold(),
                    new Color { Rgb = "FFFFFFFF" }),
                new Font(     // 2 título
                    new Bold(),
                    new FontSize { Val = 14 },
                    new Color { Rgb = "FF003F7F" })
            ),
            new Fills(
                new Fill(new PatternFill { PatternType = PatternValues.None }),
                new Fill(new PatternFill { PatternType = PatternValues.Gray125 }),
                new Fill(new PatternFill   // 2 azul cabeçalho
                {
                    PatternType      = PatternValues.Solid,
                    ForegroundColor  = new ForegroundColor { Rgb = "FF003F7F" },
                    BackgroundColor  = new BackgroundColor { Indexed = 64 },
                })
            ),
            new Borders(new Border()),
            new CellStyleFormats(new CellFormat()),
            new CellFormats(
                new CellFormat(),                                      // 0 normal
                new CellFormat { FontId = 1, FillId = 2, ApplyFont = true, ApplyFill = true },  // 1 cabeçalho
                new CellFormat { FontId = 2, ApplyFont = true }        // 2 título
            )
        );
    }
}
