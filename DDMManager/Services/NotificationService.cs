using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;

namespace DDMManager.Services;

/// <summary>
/// Registra o app no Agendador de Tarefas do Windows para
/// notificação toast toda segunda-feira às 8h.
/// Usa um script PowerShell via schtasks.exe — sem dependências extras.
/// </summary>
public class NotificationService
{
    private const string TaskName = "DDMManager_LembreteSemanal";

    /// <summary>Verifica se a tarefa já está registrada.</summary>
    public static bool TarefaRegistrada()
    {
        try
        {
            var r = Run("schtasks", $"/Query /TN \"{TaskName}\"");
            return r.exitCode == 0;
        }
        catch { return false; }
    }

    /// <summary>
    /// Registra tarefa no Agendador para exibir toast toda segunda às 8h.
    /// Usa msg.exe (disponível em qualquer Windows) como fallback ao PowerShell toast.
    /// </summary>
    public static (bool Sucesso, string Erro) RegistrarTarefa()
    {
        try
        {
            // Script PowerShell que exibe notificação toast nativa do Windows 10/11
            var script = """
                $app  = 'DDM Manager - Metalfrio'
                $msg  = 'Não esqueça de gerar os DDMs da semana!'
                [Windows.UI.Notifications.ToastNotificationManager,
                 Windows.UI.Notifications, ContentType=WindowsRuntime] | Out-Null
                [Windows.Data.Xml.Dom.XmlDocument,
                 Windows.Data.Xml.Dom, ContentType=WindowsRuntime] | Out-Null
                $xml = [Windows.UI.Notifications.ToastNotificationManager]::GetTemplateContent(
                    [Windows.UI.Notifications.ToastTemplateType]::ToastText02)
                $xml.GetElementsByTagName('text')[0].AppendChild(
                    $xml.CreateTextNode($app)) | Out-Null
                $xml.GetElementsByTagName('text')[1].AppendChild(
                    $xml.CreateTextNode($msg)) | Out-Null
                $toast = [Windows.UI.Notifications.ToastNotification]::new($xml)
                [Windows.UI.Notifications.ToastNotificationManager]::CreateToastNotifier(
                    'DDMManager').Show($toast)
                """;

            // Salva o script ao lado do exe
            var scriptPath = Path.Combine(AppContext.BaseDirectory, "lembrete_ddm.ps1");
            File.WriteAllText(scriptPath, script);

            // Registra no Agendador: toda segunda-feira às 08:00
            var cmd = $"/Create /F /TN \"{TaskName}\" " +
                      $"/TR \"powershell.exe -WindowStyle Hidden -ExecutionPolicy Bypass -File \\\"{scriptPath}\\\"\" " +
                      "/SC WEEKLY /D MON /ST 08:00";

            var (exitCode, _, stderr) = Run("schtasks", cmd);

            return exitCode == 0
                ? (true, "")
                : (false, stderr);
        }
        catch (Exception ex)
        {
            return (false, ex.Message);
        }
    }

    /// <summary>Remove a tarefa do agendador.</summary>
    public static (bool Sucesso, string Erro) RemoverTarefa()
    {
        try
        {
            var (exitCode, _, stderr) = Run("schtasks", $"/Delete /F /TN \"{TaskName}\"");
            return exitCode == 0 ? (true, "") : (false, stderr);
        }
        catch (Exception ex)
        {
            return (false, ex.Message);
        }
    }

    private static (int exitCode, string stdout, string stderr) Run(string cmd, string args)
    {
        var psi = new ProcessStartInfo(cmd, args)
        {
            RedirectStandardOutput = true,
            RedirectStandardError  = true,
            UseShellExecute        = false,
            CreateNoWindow         = true,
        };
        using var p = Process.Start(psi)!;
        var stdout = p.StandardOutput.ReadToEnd();
        var stderr = p.StandardError.ReadToEnd();
        p.WaitForExit();
        return (p.ExitCode, stdout, stderr);
    }
}
