using System.IO;
using System.Text;

namespace DDMManager.Services;

/// <summary>
/// Registra histórico de gerações em ddm_historico.csv ao lado do .exe.
/// </summary>
public class LogService
{
    private static readonly string FilePath = Path.Combine(
        AppContext.BaseDirectory, "ddm_historico.csv");

    private static readonly string Header =
        "Data;Hora;DiaNome;Arquivo;Tipo;Participantes;Status;Erro";

    public LogService() => EnsureHeader();

    private void EnsureHeader()
    {
        if (!File.Exists(FilePath))
            File.WriteAllText(FilePath, Header + Environment.NewLine, Encoding.UTF8);
    }

    public void Registrar(
        string diaNome,
        string arquivo,
        string tipo,          // "DOCX" | "PDF" | "DOCX_COMBINADO" | "PDF_COMBINADO"
        int participantes,
        bool sucesso,
        string erro = "")
    {
        try
        {
            var agora = DateTime.Now;
            var linha = string.Join(";",
                agora.ToString("dd/MM/yyyy"),
                agora.ToString("HH:mm:ss"),
                diaNome,
                Path.GetFileName(arquivo),
                tipo,
                participantes.ToString(),
                sucesso ? "OK" : "ERRO",
                erro.Replace(";", ",").Replace("\n", " ")
            );
            File.AppendAllText(FilePath, linha + Environment.NewLine, Encoding.UTF8);
        }
        catch { /* não interrompe o fluxo principal */ }
    }

    /// <summary>Abre o CSV no Excel/Notepad.</summary>
    public static void AbrirHistorico()
    {
        if (File.Exists(FilePath))
            System.Diagnostics.Process.Start(
                new System.Diagnostics.ProcessStartInfo(FilePath)
                { UseShellExecute = true });
    }

    public static bool HistoricoExiste() => File.Exists(FilePath);
}
