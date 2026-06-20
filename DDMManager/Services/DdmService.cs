using System.IO;
using DDMManager.Models;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DDMManager.Services;

/// <summary>
/// Varredura semanal de DDMs e preenchimento de cabeçalho/participantes.
/// </summary>
public class DdmService
{
    // ── Mapeamentos ───────────────────────────────────────────────────────

    private static readonly Dictionary<int, (string Num, string Nome, string Abrev)> DiaSemana = new()
    {
        [1] = ("2", "Segunda-feira", "SEG"),
        [2] = ("3", "Terça-feira",   "TER"),
        [3] = ("4", "Quarta-feira",  "QUA"),
        [4] = ("5", "Quinta-feira",  "QUI"),
        [5] = ("6", "Sexta-feira",   "SEX"),
        [6] = ("7", "Sábado",        "SÁB"),
        [0] = ("7", "Sábado",        "SÁB"),
    };

    private static readonly Dictionary<int, string> NomeMes = new()
    {
        [1]  = "1 - JANEIRO",   [2]  = "2 - FEVEREIRO", [3]  = "3 - MARÇO",
        [4]  = "4 - ABRIL",     [5]  = "5 - MAIO",      [6]  = "6 - JUNHO",
        [7]  = "7 - JULHO",     [8]  = "8 - AGOSTO",    [9]  = "9 - SETEMBRO",
        [10] = "10 - OUTUBRO",  [11] = "11 - NOVEMBRO", [12] = "12 - DEZEMBRO",
    };

    // Participantes fixos (RE → Nome)
    public static readonly (string Re, string Nome)[] Participantes =
    {
        ("10347", "ADRIANO GARCIA"),
        ("63363", "AGATHA LAÍS SALGUEIRO MAROTTI"),
        ("63051", "AMANDA FLORÊNCIO"),
        ("63175", "ANA EDUARDA SANO SILVA"),
        ("10934", "ANDRÉ LUIZ HIGA FREITAS"),
        ("63538", "AUGUSTO GRAÇA SILVESTRIN"),
        ("63231", "BARBARA RAMIRES IANHES"),
        ("62334", "CLEYTON FARIA ASSIS"),
        ("62397", "JOÃO VITOR RODRIGUES SOUZA"),
        ("184",   "JOEL BORGES DE CAMPOS JUNIOR"),
        ("61510", "KAROLINE PEDROSO SANTOS"),
        ("10437", "LETICIA FERNANDES GONÇALVES"),
        ("63186", "LUDIMILA SANTOS SOUZA"),
        ("63516", "LUIS GUILHERME DUENHAS SILVA"),
        ("63391", "RENAN DA SILVA MANTOVANI"),
        ("63999", "RODRIGO ARIAS SILVA"),
        ("11461", "VANESSA DIAS DA SILVA"),
        ("9018",  "WAGNER JUNIOR ALMEIDA"),
    };

    // ── Semana ────────────────────────────────────────────────────────────

    public static (DateTime Seg, DateTime Sex) SemanaAtual(DateTime data)
    {
        int diff = ((int)data.DayOfWeek - (int)DayOfWeek.Monday + 7) % 7;
        var seg = data.Date.AddDays(-diff);
        return (seg, seg.AddDays(4));
    }

    /// <summary>Varre a pasta raiz e retorna lista de DDMs para a semana de <paramref name="data"/>.</summary>
    public List<DdmItem> VarrerSemana(string raiz, DateTime data)
    {
        var resultado = new List<DdmItem>();
        var dir = new DirectoryInfo(raiz);
        if (!dir.Exists) return resultado;

        var (seg, sex) = SemanaAtual(data);
        int diaSemanaHoje = ((int)data.DayOfWeek + 6) % 7 + 1; // 1=Seg..7=Dom

        foreach (var pastaDir in dir.GetDirectories().OrderBy(d => d.Name))
        {
            var partes = pastaDir.Name.Split(" - ", 2);
            if (partes.Length < 2) continue;
            if (!int.TryParse(partes[0].Trim(), out int num)) continue;

            // Mapeia número para dia da semana (1-indexed Mon=1)
            var diaInfo = DiaSemana.Values.FirstOrDefault(d => d.Num == partes[0].Trim());
            if (diaInfo == default) continue;

            // Encontra pasta do mês
            var nomeMes = NomeMes[data.Month];
            var pastaMes = pastaDir.GetDirectories()
                .FirstOrDefault(d => d.Name.Equals(nomeMes, StringComparison.OrdinalIgnoreCase)
                                  || d.Name.StartsWith(data.Month + " -"));

            string? docxPath = null;
            string tema = "—";
            string dataDdm = "";
            string erro = "";

            if (pastaMes == null)
            {
                erro = $"Pasta '{nomeMes}' não encontrada";
            }
            else
            {
                // Encontra o .docx da semana (prefixo DD-MM dentro do range seg-sex)
                var candidatos = pastaMes.GetFiles("*.docx")
                    .Where(f => !f.Name.Contains("_PREENCHIDO"))
                    .OrderBy(f => f.Name)
                    .ToList();

                FileInfo? melhor = null;
                int melhorDelta = int.MaxValue;

                foreach (var f in candidatos)
                {
                    if (f.Name.Length < 5) continue;
                    if (!int.TryParse(f.Name[..2], out int dd)) continue;
                    if (f.Name[2] != '-') continue;
                    if (!int.TryParse(f.Name[3..5], out int mm)) continue;

                    try
                    {
                        var dArq = new DateTime(data.Year, mm, dd);
                        if (dArq >= seg && dArq <= sex)
                        {
                            int delta = Math.Abs((dArq - data).Days);
                            if (delta < melhorDelta)
                            {
                                melhor = f;
                                melhorDelta = delta;
                            }
                        }
                    }
                    catch { continue; }
                }

                if (melhor != null)
                {
                    docxPath = melhor.FullName;
                    tema     = ExtrairTema(melhor.Name);
                    dataDdm  = ExtrairDataDdm(melhor.Name, data.Year);
                }
                else
                {
                    erro = $"Sem DDM {seg:dd/MM}–{sex:dd/MM}";
                }
            }

            resultado.Add(new DdmItem
            {
                NumPasta  = num,
                DiaNome   = diaInfo.Nome,
                DiaAbrev  = diaInfo.Abrev,
                DocxPath  = docxPath,
                Tema      = tema,
                DataDdm   = dataDdm,
                Erro      = erro,
                EhHoje    = num == diaSemanaHoje,
                Selecionado = docxPath != null,
            });
        }

        return resultado.OrderBy(d => d.NumPasta).ToList();
    }

    // ── Helpers ───────────────────────────────────────────────────────────

    public static string ExtrairTema(string nomeArq)
    {
        var stem = Path.GetFileNameWithoutExtension(nomeArq);
        var idx  = stem.IndexOf(' ');
        return idx >= 0 ? stem[(idx + 1)..].Trim() : stem;
    }

    public static string ExtrairDataDdm(string nomeArq, int ano)
    {
        try
        {
            int dd = int.Parse(nomeArq[..2]);
            int mm = int.Parse(nomeArq[3..5]);
            return $"{dd:D2}/{mm:D2}/{ano}";
        }
        catch { return ""; }
    }

    // ── Processamento DOCX ────────────────────────────────────────────────

    public record CamposCabecalho(
        string Setor,
        string Subarea,
        string Turno,
        string Facilitador,
        string Data);

    /// <summary>
    /// Preenche o cabeçalho e injeta participantes num DOCX.
    /// Salva em <paramref name="destino"/>.
    /// Retorna número de nomes injetados.
    /// </summary>
    public int ProcessarDdm(string origem, string destino, CamposCabecalho campos)
    {
        File.Copy(origem, destino, overwrite: true);

        using var doc = WordprocessingDocument.Open(destino, isEditable: true);
        var body = doc.MainDocumentPart!.Document.Body!;

        // Labels do cabeçalho → valor a preencher
        var labels = new Dictionary<string, string>
        {
            ["SETOR/LINHA:"] = campos.Setor,
            ["TURNO:"]       = campos.Turno,
            ["SUBÁREA:"]     = campos.Subarea,
            ["FACILITADOR:"] = campos.Facilitador,
        };

        // DATA: só preenche se estiver vazia no documento
        bool dataJaPreenchida = body.Descendants<TableCell>()
            .Any(c => c.InnerText.StartsWith("DATA:") &&
                      c.InnerText.Replace("DATA:", "").Trim().Length > 0);

        if (!dataJaPreenchida)
            labels["DATA:"] = campos.Data;

        // Visita todas as células da tabela
        var visitadas = new HashSet<string>();

        foreach (var celula in body.Descendants<TableCell>())
        {
            var cellId = celula.GetHashCode().ToString();
            if (!visitadas.Add(cellId)) continue;

            foreach (var para in celula.Elements<Paragraph>())
            {
                var texto = para.InnerText.Trim();

                foreach (var (label, valor) in labels)
                {
                    if (!texto.StartsWith(label)) continue;
                    // Só preenche se vazio após o label
                    if (texto[label.Length..].Trim().Length > 0) continue;

                    // Preserva formatação do primeiro run e adiciona run com valor
                    var runs = para.Elements<Run>().ToList();
                    if (!runs.Any()) break;

                    // Remove runs extras
                    foreach (var r in runs.Skip(1)) r.Remove();

                    var novoRun = (Run)runs[0].Clone();
                    // Limpa o texto do run clonado e define o novo valor
                    foreach (var t in novoRun.Elements<Text>()) t.Remove();
                    novoRun.AppendChild(new Text(" " + valor)
                        { Space = SpaceProcessingModeValues.Preserve });
                    para.AppendChild(novoRun);
                    break;
                }
            }
        }

        // ── Participantes ─────────────────────────────────────────────────
        int nomesInjetados = 0;
        int partIdx = 0;

        foreach (var linha in body.Descendants<TableRow>())
        {
            var celulas = linha.Elements<TableCell>().ToList();
            if (celulas.Count < 3) continue;

            // Linha de participante: primeira célula = número inteiro
            var txtPrimeira = celulas[0].InnerText.Trim();
            if (!int.TryParse(txtPrimeira, out _)) continue;

            if (partIdx < Participantes.Length)
            {
                var (re, nome) = Participantes[partIdx];
                EscreverCelula(celulas[1], re,   centralizar: true);
                EscreverCelula(celulas[2], nome,  centralizar: false);
                nomesInjetados++;
            }
            partIdx++;
        }

        doc.Save();
        return nomesInjetados;
    }

    private static void EscreverCelula(TableCell celula, string valor, bool centralizar)
    {
        // Alinhamento vertical
        var tcPr = celula.GetFirstChild<TableCellProperties>()
                   ?? celula.PrependChild(new TableCellProperties());
        var vAlign = tcPr.GetFirstChild<TableCellVerticalAlignment>()
                     ?? tcPr.AppendChild(new TableCellVerticalAlignment());
        vAlign.Val = TableVerticalAlignmentValues.Center;

        var para = celula.GetFirstChild<Paragraph>()
                   ?? celula.AppendChild(new Paragraph());

        // Alinhamento horizontal
        var pPr = para.GetFirstChild<ParagraphProperties>()
                  ?? para.PrependChild(new ParagraphProperties());
        var jc = pPr.GetFirstChild<Justification>()
                 ?? pPr.AppendChild(new Justification());
        jc.Val = centralizar ? JustificationValues.Center : JustificationValues.Left;

        // Remove runs existentes e adiciona novo
        foreach (var r in para.Elements<Run>().ToList()) r.Remove();

        var run = new Run();
        var rPr = new RunProperties();
        var fonts = new RunFonts
        {
            Ascii    = "Arial",
            HighAnsi = "Arial",
            EastAsia = "Batang",
            ComplexScript = "Arial"
        };
        rPr.AppendChild(fonts);
        rPr.AppendChild(new FontSize { Val = "18" });      // 9pt
        rPr.AppendChild(new FontSizeComplexScript { Val = "18" });
        run.AppendChild(rPr);
        run.AppendChild(new Text(valor) { Space = SpaceProcessingModeValues.Preserve });
        para.AppendChild(run);
    }
}
