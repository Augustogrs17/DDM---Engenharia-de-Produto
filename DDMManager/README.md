# DDM Manager — C# / WPF

Reescrita em C# WPF com Fluent Design (ModernWPF).  
Funcionalidades idênticas à versão Python, com UI nativa do Windows 11.

## Estrutura do projeto

```
DDMManager/
├── DDMManager.csproj          ← configurações e dependências NuGet
├── App.xaml / App.xaml.cs     ← entrada + tema global
├── Models/
│   └── DdmItem.cs             ← modelo de dados (DDM + Usuário)
├── Services/
│   ├── AuthService.cs         ← login, cadastro de usuários, SHA-256
│   ├── DdmService.cs          ← varredura de pastas + preenchimento DOCX
│   └── WordService.cs         ← PDF e DOCX combinado via Word COM
├── Views/
│   ├── LoginWindow.xaml/.cs   ← tela de login
│   └── MainWindow.xaml/.cs    ← tela principal
└── .github/workflows/
    └── build.yml              ← CI/CD: gera .exe no GitHub Actions
```

## Gerar o .exe via GitHub Actions

1. Crie um repositório no GitHub (pode ser privado)
2. Faça upload de **todos os arquivos** mantendo a estrutura de pastas
3. Vá em **Actions → Build DDM Manager (C#) → Run workflow**
4. Após ~2 minutos, baixe o `.exe` em **Artifacts → DDM_Manager_Metalfrio_CSharp**

## Dependências NuGet (instaladas automaticamente)

| Pacote | Uso |
|--------|-----|
| `ModernWpfUI` | Fluent Design, tema do Windows 11 |
| `DocumentFormat.OpenXml` | Manipulação de .docx (substitui python-docx) |
| `Newtonsoft.Json` | Persistência de usuários em JSON |
| `Microsoft.Office.Interop.Word` | Conversão PDF e mesclagem via Word |

## Requisito em produção

**Microsoft Word instalado** — necessário para geração de PDF e combinação de DOCX.

## Login inicial

| Campo | Valor |
|-------|-------|
| RE    | 184   |
| Senha | metalfrio |

## Fase 2 (próximos passos)

- [ ] AdminWindow: gestão de usuários (adicionar, remover, redefinir senha)
- [ ] Animações de transição entre telas
- [ ] Notificações toast nativas do Windows 11
- [ ] Barra de progresso durante conversão PDF
- [ ] Salvar preferências do usuário (pasta raiz, campos) entre sessões
