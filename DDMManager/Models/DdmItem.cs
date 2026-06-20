using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace DDMManager.Models;

/// <summary>
/// Representa um DDM detectado na pasta de rede para a semana corrente.
/// </summary>
public class DdmItem : INotifyPropertyChanged
{
    public int    NumPasta  { get; init; }
    public string DiaNome   { get; init; } = "";
    public string DiaAbrev  { get; init; } = "";
    public string? DocxPath { get; init; }
    public string Tema      { get; init; } = "";
    public string DataDdm   { get; init; } = "";
    public string Erro      { get; init; } = "";
    public bool   EhHoje    { get; init; }

    public bool TemArquivo => DocxPath != null;

    private bool _selecionado;
    public bool Selecionado
    {
        get => _selecionado;
        set { _selecionado = value; OnPropertyChanged(); }
    }

    public event PropertyChangedEventHandler? PropertyChanged;
    protected void OnPropertyChanged([CallerMemberName] string? name = null)
        => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
}
