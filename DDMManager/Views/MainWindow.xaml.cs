using System.Collections.ObjectModel;
using System.Diagnostics;
using System.IO;
using System.Windows;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Threading;
using DDMManager.Models;
using DDMManager.Services;
using Microsoft.Win32;

namespace DDMManager.Views;

public partial class MainWindow : Window
{
    // ── Serviços ──────────────────────────────────────────────────────────
    private readonly DdmService      _ddmSvc   = new();
    private readonly WordService     _wordSvc  = new();
    private readonly PrintService    _printSvc = new();
    private readonly ReportService   _reportSvc = new();
    private readonly SettingsService _settingsSvc = new();
    private readonly LogService      _logSvc  = new();
    private AppSettings              _settings = new();

    // ── Estado ────────────────────────────────────────────────────────────
    private readonly ObservableCollection<DdmItem> _itens = new();
    private int    _semanaOffset = 0;   // 0=atual, -1=anterior, +1=próxima
    private string _pastaSaida  = "";
    private bool   _operacaoEmAndamento = false;

    // ── Relógio ───────────────────────────────────────────────────────────
    private readonly DispatcherTimer _timer = new() { Interval = TimeSpan.FromSeconds(1) };

    // ── Log colors ────────────────────────────────────────────────────────
    private static readonly SolidColorBrush LogSuccess = new(Color.FromRgb(0x10, 0xB9, 0x81));
    private static readonly SolidColorBrush LogWarn    = new(Color.FromRgb(0xF5, 0x9E, 0x0B));
    private static readonly SolidColorBrush LogError   = new(Color.FromRgb(0xEF, 0x44, 0x44));
    private static readonly SolidColorBrush LogAccent  = new(Color.FromRgb(0x00, 0xD4, 0xFF));
    private static readonly SolidColorBrush LogMuted   = new(Color.FromRgb(0x94, 0xA3, 0xB8));

    public MainWindow()
    {
        InitializeComponent();
        ListaDdms.ItemsSource      = _itens;
        RunNParticipantes.Text     = DdmService.Participantes.Length.ToString();

        // Registra conversor no recursos
        Resources["BoolToCursorConverter"] = new BoolToCursorConverter();
    }

    // ════════════════════════════════════════════════════════════════════════
    // Inicialização
    // ════════════════════════════════════════════════════════════════════════

    private void Window_Loaded(object sender, RoutedEventArgs e)
    {
        CarregarConfiguracoes();
        IniciarRelogio();
        CarregarImpressoras();
        AtualizarStatusNotificacao();
        BtnHistorico.IsEnabled = LogService.HistoricoExiste();
        VarrerSemana();
    }

    private void CarregarConfiguracoes()
    {
        _settings = _settingsSvc.Load();
        TxtRaiz.Text        = _settings.PastaRaiz;
        TxtSetor.Text       = _settings.Setor;
        TxtSubarea.Text     = _settings.Subarea;
        TxtTurno.Text       = _settings.Turno;
        TxtFacilitador.Text = _settings.Facilitador;
        _semanaOffset       = _settings.SemanaOffset;

        if (!string.IsNullOrWhiteSpace(_settings.PastaSaida))
        {
            _pastaSaida      = _settings.PastaSaida;
            TxtSaida.Text    = _settings.PastaSaida;
            TxtSaida.Opacity = 1;
        }
    }

    private void SalvarConfiguracoes()
    {
        _settings.PastaRaiz    = TxtRaiz.Text;
        _settings.Setor        = TxtSetor.Text;
        _settings.Subarea      = TxtSubarea.Text;
        _settings.Turno        = TxtTurno.Text;
        _settings.Facilitador  = TxtFacilitador.Text;
        _settings.PastaSaida   = _pastaSaida;
        _settings.SemanaOffset = _semanaOffset;
        _settingsSvc.Save(_settings);
    }

    private void IniciarRelogio()
    {
        string[] dias = { "Domingo","Segunda","Terça","Quarta","Quinta","Sexta","Sábado" };
        _timer.Tick += (_, _) =>
        {
            var n = DateTime.Now;
            TxtRelogio.Text = $"{dias[(int)n.DayOfWeek]}, {n:dd/MM/yyyy  HH:mm:ss}";
        };
        _timer.Start();
    }

    private void CarregarImpressoras()
    {
        CmbImpressora.Items.Add("(Impressora padrão do sistema)");
        foreach (var p in PrintService.ListarImpressoras())
            CmbImpressora.Items.Add(p);
        CmbImpressora.SelectedIndex = 0;
    }

    private void AtualizarStatusNotificacao()
    {
        var ativa = NotificationService.TarefaRegistrada();
        TxtStatusNotif.Text      = ativa
            ? "✔  Lembrete ativo — toda segunda às 8h"
            : "Sem lembrete configurado";
        TxtStatusNotif.Foreground = ativa ? LogSuccess : LogMuted;
        BtnAtivarNotif.IsEnabled    = !ativa;
        BtnDesativarNotif.IsEnabled = ativa;
    }

    // ════════════════════════════════════════════════════════════════════════
    // Varredura
    // ════════════════════════════════════════════════════════════════════════

    private void VarrerSemana()
    {
        LimparLog();
        var raiz     = TxtRaiz.Text.Trim();
        var dataRef  = DateTime.Today.AddDays(_semanaOffset * 7);
        var (seg, sex) = DdmService.SemanaAtual(dataRef);

        // Label da semana
        TxtSemana.Text = $"  {seg:dd/MM} – {sex:dd/MM}";
        TxtSemanaLabel.Text = _semanaOffset == 0 ? "(semana atual)"
                            : _semanaOffset <  0 ? $"({Math.Abs(_semanaOffset)} semana(s) atrás)"
                                                 : $"({_semanaOffset} semana(s) à frente)";

        Log($"Semana: {seg:dd/MM/yyyy} – {sex:dd/MM/yyyy}", LogMuted);
        Log($"Raiz:   {raiz}\n", LogMuted);

        _itens.Clear();
        var lista = _ddmSvc.VarrerSemana(raiz, dataRef);
        foreach (var item in lista) _itens.Add(item);

        // Subscribe preview on selection change
        foreach (var item in _itens)
            item.PropertyChanged += (s, _) =>
            {
                if (s is DdmItem d && d.Selecionado)
                    AtualizarPreview(d);
            };

        var ok  = lista.Count(d => d.TemArquivo);
        var err = lista.Count(d => !d.TemArquivo);

        Log($"✔  {ok} DDM(s) encontrado(s)", LogSuccess);
        foreach (var d in lista.Where(d => d.TemArquivo))
            Log($"   {d.DiaNome}: {Path.GetFileName(d.DocxPath!)}", LogMuted);

        if (err > 0)
        {
            Log($"\n⚠  {err} pasta(s) sem DDM:", LogWarn);
            foreach (var d in lista.Where(d => !d.TemArquivo))
                Log($"   {d.DiaNome}: {d.Erro}", LogMuted);
        }

        AtualizarChecklist();
        BtnHistorico.IsEnabled = LogService.HistoricoExiste();
    }

    // ════════════════════════════════════════════════════════════════════════
    // Checklist pré-geração
    // ════════════════════════════════════════════════════════════════════════

    private void AtualizarChecklist()
    {
        // Pasta raiz acessível
        var pastaOk = Directory.Exists(TxtRaiz.Text.Trim());
        ChkPastaRaiz.Text       = pastaOk ? "✔  Pasta raiz" : "✖  Pasta raiz inacessível";
        ChkPastaRaiz.Foreground = pastaOk ? LogSuccess : LogError;

        // Word instalado (verifica no registry)
        var wordOk = VerificarWord();
        ChkWord.Text       = wordOk ? "✔  Word instalado" : "✖  Word não encontrado";
        ChkWord.Foreground = wordOk ? LogSuccess : LogWarn;

        // DDMs disponíveis
        var nDdms = _itens.Count(d => d.TemArquivo);
        var ddmsOk = nDdms > 0;
        ChkDdms.Text       = ddmsOk ? $"✔  {nDdms} DDM(s)" : "✖  Sem DDMs";
        ChkDdms.Foreground = ddmsOk ? LogSuccess : LogWarn;

        // Campos preenchidos
        var camposOk = !string.IsNullOrWhiteSpace(TxtSetor.Text)
                    && !string.IsNullOrWhiteSpace(TxtFacilitador.Text);
        ChkCampos.Text       = camposOk ? "✔  Campos preenchidos" : "✖  Campos incompletos";
        ChkCampos.Foreground = camposOk ? LogSuccess : LogWarn;
    }

    private static bool VerificarWord()
    {
        try
        {
            using var key = Microsoft.Win32.Registry.LocalMachine
                .OpenSubKey(@"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\WINWORD.EXE");
            return key != null;
        }
        catch { return false; }
    }

    // ════════════════════════════════════════════════════════════════════════
    // Preview
    // ════════════════════════════════════════════════════════════════════════

    private void AtualizarPreview(DdmItem? item)
    {
        if (item == null || !item.TemArquivo)
        {
            PreviewVazio();
            return;
        }

        TxtPreviewVazio.Visibility    = Visibility.Collapsed;
        PreviewConteudo.Visibility    = Visibility.Visible;

        PreviewDia.Text       = item.DiaNome;
        PreviewTema.Text      = item.Tema;
        PreviewData.Text      = item.DataDdm;
        PreviewArquivo.Text   = Path.GetFileName(item.DocxPath!);
        PrvSetor.Text         = TxtSetor.Text;
        PrvTurno.Text         = TxtTurno.Text;
        PrvFacilitador.Text   = TxtFacilitador.Text;
        PrvData.Text          = !string.IsNullOrEmpty(item.DataDdm)
                                ? item.DataDdm
                                : $"{DateTime.Today:dd/MM/yyyy} (hoje)";
    }

    private void PreviewVazio()
    {
        TxtPreviewVazio.Visibility  = Visibility.Visible;
        PreviewConteudo.Visibility  = Visibility.Collapsed;
    }

    // ════════════════════════════════════════════════════════════════════════
    // Geração de arquivos
    // ════════════════════════════════════════════════════════════════════════

    private List<DdmItem> Selecionados()
    {
        var sel = _itens.Where(d => d.Selecionado && d.TemArquivo).ToList();
        if (!sel.Any())
            MessageBox.Show("Selecione ao menos um DDM da lista.",
                "Nenhum selecionado", MessageBoxButton.OK, MessageBoxImage.Information);
        return sel;
    }

    private string PastaEfetiva(string fallback)
    {
        if (!string.IsNullOrWhiteSpace(_pastaSaida))
        {
            Directory.CreateDirectory(_pastaSaida);
            return _pastaSaida;
        }
        return fallback;
    }

    private DdmService.CamposCabecalho MontarCampos(DdmItem d) => new(
        Setor:       TxtSetor.Text,
        Subarea:     TxtSubarea.Text,
        Turno:       TxtTurno.Text,
        Facilitador: TxtFacilitador.Text,
        Data:        !string.IsNullOrEmpty(d.DataDdm)
                     ? d.DataDdm : DateTime.Today.ToString("dd/MM/yyyy"));

    private async Task<List<string>> GerarDocxs(List<DdmItem> sel)
    {
        var gerados = new List<string>();
        await Task.Run(() =>
        {
            foreach (var d in sel)
            {
                var pasta = PastaEfetiva(Path.GetDirectoryName(d.DocxPath!)!);
                var dst   = Path.Combine(pasta,
                    Path.GetFileNameWithoutExtension(d.DocxPath!) + "_PREENCHIDO.docx");
                try
                {
                    int nomes = _ddmSvc.ProcessarDdm(d.DocxPath!, dst, MontarCampos(d));
                    Dispatch(() => Log($"✔  {d.DiaNome}: {Path.GetFileName(dst)} ({nomes} nomes)", LogSuccess));
                    _logSvc.Registrar(d.DiaNome, dst, "DOCX", nomes, true);
                    gerados.Add(dst);
                }
                catch (Exception ex)
                {
                    Dispatch(() => Log($"✖  {d.DiaNome}: {ex.Message}", LogError));
                    _logSvc.Registrar(d.DiaNome, d.DocxPath!, "DOCX", 0, false, ex.Message);
                }
            }
        });
        return gerados;
    }

    // ── Gerar DOCX individuais ────────────────────────────────────────────

    private async void BtnGerarDocx_Click(object sender, RoutedEventArgs e)
    {
        var sel = Selecionados();
        if (!sel.Any() || _operacaoEmAndamento) return;
        IniciarOperacao($"Gerando {sel.Count} DOCX(s)...");
        Log($"\n─── Gerando {sel.Count} DOCX(s)  [Ctrl+D] ───", LogAccent);

        var gerados = await GerarDocxs(sel);

        FinalizarOperacao();
        BtnHistorico.IsEnabled = true;

        if (gerados.Any())
        {
            var r = MessageBox.Show($"{gerados.Count} DOCX(s) gerado(s).\nAbrir pasta?",
                "Concluído", MessageBoxButton.YesNo, MessageBoxImage.Information);
            if (r == MessageBoxResult.Yes)
                Process.Start("explorer.exe", Path.GetDirectoryName(gerados[0])!);
        }
    }

    // ── Gerar PDF ─────────────────────────────────────────────────────────

    private async void BtnGerarPdf_Click(object sender, RoutedEventArgs e)
    {
        var sel = Selecionados();
        if (!sel.Any() || _operacaoEmAndamento) return;
        IniciarOperacao($"Gerando {sel.Count} PDF(s)...");
        Log($"\n─── Gerando {sel.Count} PDF(s)  [Ctrl+P] ───", LogAccent);

        var docxs = await GerarDocxs(sel);
        if (!docxs.Any()) { FinalizarOperacao(); return; }

        Log($"\n   Convertendo via Word...", LogMuted);
        List<(string? Pdf, string Erro)> resultados = null!;
        await Task.Run(() => { resultados = _wordSvc.ConverterParaPdf(docxs); });

        var pdfs = new List<string>();
        for (int i = 0; i < resultados.Count; i++)
        {
            var (pdf, erro) = resultados[i];
            if (pdf != null)
            {
                Log($"✔  {sel[i].DiaNome}: {Path.GetFileName(pdf)}", LogSuccess);
                _logSvc.Registrar(sel[i].DiaNome, pdf, "PDF", 0, true);
                pdfs.Add(pdf);
            }
            else
            {
                Log($"✖  {sel[i].DiaNome}: {erro}", LogError);
                _logSvc.Registrar(sel[i].DiaNome, docxs[i], "PDF", 0, false, erro);
            }
        }

        FinalizarOperacao();
        BtnHistorico.IsEnabled = true;
        if (!pdfs.Any()) return;

        if (pdfs.Count == 1)
        {
            Process.Start(new ProcessStartInfo(pdfs[0]) { UseShellExecute = true });
        }
        else
        {
            Log($"\n   Mesclando {pdfs.Count} PDFs...", LogMuted);
            string? combinado = null;
            await Task.Run(() =>
            {
                try
                {
                    var pasta = PastaEfetiva(Path.GetDirectoryName(pdfs[0])!);
                    combinado = _wordSvc.MesclarPdfs(pdfs, docxs, pasta);
                }
                catch (Exception ex)
                {
                    Dispatch(() => Log($"⚠  Mesclagem: {ex.Message}", LogWarn));
                }
            });

            if (combinado != null)
            {
                Log($"✔  Combinado: {Path.GetFileName(combinado)}", LogSuccess);
                Process.Start(new ProcessStartInfo(combinado) { UseShellExecute = true });
            }
            else
                foreach (var p in pdfs)
                    Process.Start(new ProcessStartInfo(p) { UseShellExecute = true });
        }
    }

    // ── Combinar DOCX ─────────────────────────────────────────────────────

    private async void BtnCombinarDocx_Click(object sender, RoutedEventArgs e)
    {
        var sel = Selecionados();
        if (!sel.Any() || _operacaoEmAndamento) return;
        if (sel.Count < 2)
        {
            MessageBox.Show("Selecione ao menos 2 DDMs para combinar.",
                "Atenção", MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }

        IniciarOperacao($"Combinando {sel.Count} DOCX(s)...");
        Log($"\n─── Combinando {sel.Count} DOCX(s) ───", LogAccent);

        var docxs = await GerarDocxs(sel);
        if (docxs.Count < 2) { FinalizarOperacao(); return; }

        string? combinado = null;
        await Task.Run(() =>
        {
            try
            {
                var pasta = PastaEfetiva(Path.GetDirectoryName(docxs[0])!);
                combinado = _wordSvc.CombinarDocx(docxs, pasta);
            }
            catch (Exception ex)
            {
                Dispatch(() => Log($"✖  {ex.Message}", LogError));
            }
        });

        FinalizarOperacao();
        BtnHistorico.IsEnabled = true;

        if (combinado != null)
        {
            Log($"✔  Combinado: {Path.GetFileName(combinado)}", LogSuccess);
            _logSvc.Registrar("SEMANA", combinado, "DOCX_COMBINADO", 0, true);
            var r = MessageBox.Show($"DOCX combinado:\n{Path.GetFileName(combinado)}\n\nAbrir?",
                "Concluído", MessageBoxButton.YesNo, MessageBoxImage.Information);
            if (r == MessageBoxResult.Yes)
                Process.Start(new ProcessStartInfo(combinado) { UseShellExecute = true });
        }
    }

    // ── Imprimir direto ───────────────────────────────────────────────────

    private async void BtnImprimirDireto_Click(object sender, RoutedEventArgs e)
    {
        var sel = Selecionados();
        if (!sel.Any() || _operacaoEmAndamento) return;

        var impressora = CmbImpressora.SelectedIndex > 0
            ? CmbImpressora.SelectedItem?.ToString()
            : null;

        var confirmar = MessageBox.Show(
            $"Imprimir {sel.Count} DDM(s) diretamente?\n" +
            $"Impressora: {impressora ?? "padrão do sistema"}",
            "Confirmar impressão", MessageBoxButton.YesNo, MessageBoxImage.Question);
        if (confirmar != MessageBoxResult.Yes) return;

        IniciarOperacao("Gerando e imprimindo...");
        Log($"\n─── Impressão direta: {sel.Count} DDM(s) ───", LogAccent);

        var docxs = await GerarDocxs(sel);
        if (!docxs.Any()) { FinalizarOperacao(); return; }

        Log("   Enviando para impressora...", LogMuted);
        List<(string Arquivo, bool Sucesso, string Erro)> resultados = null!;
        await Task.Run(() =>
        {
            resultados = _printSvc.ImprimirDocxs(docxs, impressora);
        });

        foreach (var (arq, ok, erro) in resultados)
        {
            if (ok)
            {
                Log($"✔  {Path.GetFileName(arq)}: enviado", LogSuccess);
                _logSvc.Registrar("IMPRESSÃO", arq, "PRINT", 0, true);
            }
            else
            {
                Log($"✖  {Path.GetFileName(arq)}: {erro}", LogError);
                _logSvc.Registrar("IMPRESSÃO", arq, "PRINT", 0, false, erro);
            }
        }
        FinalizarOperacao();
        BtnHistorico.IsEnabled = true;
    }

    // ── Relatório semanal ─────────────────────────────────────────────────

    private async void BtnRelatorio_Click(object sender, RoutedEventArgs e)
    {
        if (_operacaoEmAndamento) return;
        var ddmsOk = _itens.Where(d => d.TemArquivo).ToList();
        if (!ddmsOk.Any())
        {
            MessageBox.Show("Nenhum DDM disponível para gerar relatório.",
                "Sem dados", MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }

        IniciarOperacao("Gerando relatório...");
        Log("\n─── Gerando relatório semanal ───", LogAccent);

        var dataRef = DateTime.Today.AddDays(_semanaOffset * 7);
        var itensRel = ddmsOk.Select(d =>
            (d.DiaNome, d.DocxPath!, d.Tema, d.DataDdm,
             DdmService.Participantes.Length));

        string? caminho = null;
        await Task.Run(() =>
        {
            try
            {
                var pasta = !string.IsNullOrWhiteSpace(_pastaSaida)
                    ? _pastaSaida
                    : AppContext.BaseDirectory;
                caminho = _reportSvc.GerarRelatorioSemanal(itensRel, pasta, dataRef);
            }
            catch (Exception ex)
            {
                Dispatch(() => Log($"✖  {ex.Message}", LogError));
            }
        });

        FinalizarOperacao();
        if (caminho != null)
        {
            Log($"✔  Relatório: {Path.GetFileName(caminho)}", LogSuccess);
            Process.Start(new ProcessStartInfo(caminho) { UseShellExecute = true });
        }
    }

    // ════════════════════════════════════════════════════════════════════════
    // Progresso
    // ════════════════════════════════════════════════════════════════════════

    private void IniciarOperacao(string msg)
    {
        _operacaoEmAndamento    = true;
        TxtProgresso.Text       = msg;
        ProgressBar.Visibility  = Visibility.Visible;
    }

    private void FinalizarOperacao()
    {
        _operacaoEmAndamento   = false;
        ProgressBar.Visibility = Visibility.Collapsed;
    }

    // ════════════════════════════════════════════════════════════════════════
    // Event handlers UI
    // ════════════════════════════════════════════════════════════════════════

    private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        => SalvarConfiguracoes();

    private void BtnAtualizar_Click(object sender, RoutedEventArgs e) => VarrerSemana();
    private void BtnSair_Click(object sender, RoutedEventArgs e) => Application.Current.Shutdown();
    private void BtnLimparLog_Click(object sender, RoutedEventArgs e) => LimparLog();

    private void BtnSemanaAnterior_Click(object sender, RoutedEventArgs e)
    { _semanaOffset--; VarrerSemana(); }

    private void BtnSemanaProxima_Click(object sender, RoutedEventArgs e)
    { _semanaOffset++; VarrerSemana(); }

    private void BtnBrowseRaiz_Click(object sender, RoutedEventArgs e)
    {
        var dlg = new OpenFolderDialog { Title = "Pasta raiz DDM 2026" };
        if (dlg.ShowDialog(this) == true)
        { TxtRaiz.Text = dlg.FolderName; VarrerSemana(); }
    }

    private void BtnBrowseSaida_Click(object sender, RoutedEventArgs e)
    {
        var dlg = new OpenFolderDialog { Title = "Onde salvar os arquivos gerados" };
        if (dlg.ShowDialog(this) == true)
        { _pastaSaida = dlg.FolderName; TxtSaida.Text = dlg.FolderName; TxtSaida.Opacity = 1; }
    }

    private void BtnLimparSaida_Click(object sender, RoutedEventArgs e)
    {
        _pastaSaida      = "";
        TxtSaida.Text    = "(padrão: mesma pasta do DDM)";
        TxtSaida.Opacity = 0.45;
    }

    private void BtnSelecionarTudo_Click(object sender, RoutedEventArgs e)
    {
        foreach (var d in _itens.Where(d => d.TemArquivo)) d.Selecionado = true;
    }

    private void BtnLimparSelecao_Click(object sender, RoutedEventArgs e)
    {
        foreach (var d in _itens) d.Selecionado = false;
        PreviewVazio();
    }

    private void Campo_Changed(object sender, System.Windows.Controls.TextChangedEventArgs e)
    {
        AtualizarChecklist();
        // Atualiza preview se houver item selecionado
        var sel = _itens.FirstOrDefault(d => d.Selecionado && d.TemArquivo);
        if (sel != null) AtualizarPreview(sel);
    }

    private void BtnHistorico_Click(object sender, RoutedEventArgs e)
        => LogService.AbrirHistorico();

    private void BtnAtivarNotif_Click(object sender, RoutedEventArgs e)
    {
        var (ok, erro) = NotificationService.RegistrarTarefa();
        if (ok)
        {
            Log("✔  Lembrete semanal ativado (toda segunda às 8h)", LogSuccess);
            AtualizarStatusNotificacao();
        }
        else
            Log($"✖  Falha ao ativar lembrete: {erro}", LogError);
    }

    private void BtnDesativarNotif_Click(object sender, RoutedEventArgs e)
    {
        var (ok, erro) = NotificationService.RemoverTarefa();
        if (ok)
        {
            Log("✔  Lembrete semanal desativado", LogMuted);
            AtualizarStatusNotificacao();
        }
        else
            Log($"✖  Falha ao desativar: {erro}", LogError);
    }

    // ════════════════════════════════════════════════════════════════════════
    // Commands (para key bindings)
    // ════════════════════════════════════════════════════════════════════════

    public ICommand AtualizarCommand        => new RelayCommand(_ => VarrerSemana());
    public ICommand GerarPdfCommand         => new RelayCommand(_ => BtnGerarPdf_Click(this, new()));
    public ICommand GerarDocxCommand        => new RelayCommand(_ => BtnGerarDocx_Click(this, new()));
    public ICommand SemanaAnteriorCommand   => new RelayCommand(_ => { _semanaOffset--; VarrerSemana(); });
    public ICommand SemanaProximaCommand    => new RelayCommand(_ => { _semanaOffset++; VarrerSemana(); });
    public ICommand ToggleSelecaoCommand    => new RelayCommand<DdmItem>(item =>
    {
        if (item is { TemArquivo: true })
        {
            item.Selecionado = !item.Selecionado;
            if (item.Selecionado) AtualizarPreview(item);
        }
    });

    // ════════════════════════════════════════════════════════════════════════
    // Log
    // ════════════════════════════════════════════════════════════════════════

    private void Log(string msg, SolidColorBrush? cor = null)
    {
        var para = new Paragraph(new Run(msg))
        {
            Foreground = cor ?? (SolidColorBrush)RtbLog.Foreground,
            Margin     = new Thickness(0, 1, 0, 1),
        };
        RtbLog.Document.Blocks.Add(para);
        RtbLog.ScrollToEnd();
    }

    private void LimparLog() => RtbLog.Document.Blocks.Clear();

    private void Dispatch(Action a) => Dispatcher.Invoke(a);
}

// ── Helpers ───────────────────────────────────────────────────────────────────

public class RelayCommand<T>(Action<T?> execute) : ICommand
{
    public bool CanExecute(object? p) => true;
    public void Execute(object? p) => execute(p is T t ? t : default);
    public event EventHandler? CanExecuteChanged;
}

public class RelayCommand(Action<object?> execute) : ICommand
{
    public bool CanExecute(object? p) => true;
    public void Execute(object? p) => execute(p);
    public event EventHandler? CanExecuteChanged;
}

public class BoolToCursorConverter : System.Windows.Data.IValueConverter
{
    public object Convert(object v, Type t, object p, System.Globalization.CultureInfo c)
        => (bool)v ? Cursors.Hand : Cursors.Arrow;
    public object ConvertBack(object v, Type t, object p, System.Globalization.CultureInfo c)
        => throw new NotImplementedException();
}
