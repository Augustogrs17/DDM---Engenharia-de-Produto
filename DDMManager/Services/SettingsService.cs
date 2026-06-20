using System.IO;
using System.Text.Json;

namespace DDMManager.Services;

/// <summary>
/// Persiste preferências do usuário em appsettings.json ao lado do .exe.
/// </summary>
public class AppSettings
{
    public string PastaRaiz      { get; set; } = @"Q:\Transferencia\DDM 2026";
    public string PastaSaida     { get; set; } = "";
    public string Setor          { get; set; } = "Engenharia Industrial";
    public string Subarea        { get; set; } = "Engenharia de Produto";
    public string Turno          { get; set; } = "ADM";
    public string Facilitador    { get; set; } = "Leitura Individual";
    public int    SemanaOffset   { get; set; } = 0;   // 0=atual, -1=anterior, +1=próxima
    public string UltimaAtualizacao { get; set; } = "";
}

public class SettingsService
{
    private static readonly string FilePath = Path.Combine(
        AppContext.BaseDirectory, "appsettings.json");

    private static readonly JsonSerializerOptions JsonOpts = new()
    {
        WriteIndented = true,
        PropertyNameCaseInsensitive = true,
    };

    public AppSettings Load()
    {
        try
        {
            if (File.Exists(FilePath))
            {
                var json = File.ReadAllText(FilePath);
                return JsonSerializer.Deserialize<AppSettings>(json, JsonOpts)
                       ?? new AppSettings();
            }
        }
        catch { /* retorna padrão se corrompido */ }
        return new AppSettings();
    }

    public void Save(AppSettings settings)
    {
        try
        {
            settings.UltimaAtualizacao = DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss");
            var json = JsonSerializer.Serialize(settings, JsonOpts);
            File.WriteAllText(FilePath, json);
        }
        catch { /* silencia erros de IO */ }
    }
}
