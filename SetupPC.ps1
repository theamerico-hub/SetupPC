<#
.SYNOPSIS
    SetupPC - Pos-formatacao automatica para Windows 10/11.
    Autor: Danilo Américo
    1) Cria ponto de restauracao
    2) Garante que o winget esteja funcionando
    3) Instala os programas padrao (winget + instaladores locais)
    4) Faz o debloat do sistema (apps, telemetria, anuncios, Copilot, etc.)

.EXAMPLE
    .\SetupPC.ps1                 # faz tudo
    .\SetupPC.ps1 -SemDebloat     # so instala os programas
    .\SetupPC.ps1 -SemProgramas   # so faz o debloat
    .\SetupPC.ps1 -Silencioso     # nao pergunta nada (nem reinicia no final)
    .\SetupPC.ps1 -Office 5       # ja escolhe o Office 2024 Pro Plus sem mostrar o menu
    .\SetupPC.ps1 -Console        # usa o modo texto (sem interface grafica)
    .\SetupPC.ps1 -Gamer -Office 0 -Otimizar -DesativarWindowsUpdate   # gamer completo, sem Office
#>
[CmdletBinding()]
param(
    [switch]$SemProgramas,
    [switch]$SemDebloat,
    [switch]$SemPontoRestauracao,
    [switch]$Silencioso,
    [switch]$Gamer,                   # inclui os programas da seção gamer
    [switch]$DesativarWindowsUpdate,  # desativa o Windows Update no final da instalação
    [switch]$Otimizar,                # aplica a otimização avançada no final da instalação
    [switch]$Console,                 # modo texto (sem interface gráfica)
    [string]$Office          # número da versão do menu (pula a pergunta). '0' = não instalar
)

# =====================================================================
#  CONFIGURAÇÃO  -  edite aqui
# =====================================================================

# Programas instalados pelo winget (com o Chocolatey como reserva).
# Id:     ID no winget       ->  descubra com:  winget search nome-do-programa
# Choco:  nome no Chocolatey ->  descubra em:   https://community.chocolatey.org/packages
# Fonte:  'winget' (padrão) ou 'msstore' (apps da Microsoft Store).
# Locale: tenta instalar naquele idioma; se não existir, instala o padrão.
$Programas = @(
    @{ Nome = 'Google Chrome';          Id = 'Google.Chrome';                Choco = 'googlechrome' }
    @{ Nome = 'WinRAR';                 Id = 'RARLab.WinRAR';                Choco = 'winrar'; Locale = 'pt-BR' }
    @{ Nome = 'AnyDesk';                Id = 'AnyDesk.AnyDesk';              Choco = 'anydesk.portable'; ChocoParams = '/INSTALL' }
    @{ Nome = 'Adobe Acrobat Reader';   Id = 'Adobe.Acrobat.Reader.64-bit';  Choco = 'adobereader'; Locale = 'pt-BR' }

)

# Programas opcionais: aparecem DESMARCADOS na interface gráfica, para marcar quando quiser.
$ProgramasOpcionais = @(
    @{ Nome = 'Mozilla Firefox';      Id = 'Mozilla.Firefox.pt-BR';        Choco = 'firefox'; ChocoParams = '/l:pt-BR' }
    @{ Nome = '7-Zip';                Id = '7zip.7zip';                    Choco = '7zip.install' }
    @{ Nome = 'VLC';                  Id = 'VideoLAN.VLC';                 Choco = 'vlc' }
    @{ Nome = 'Notepad++';            Id = 'Notepad++.Notepad++';          Choco = 'notepadplusplus.install' }
    @{ Nome = 'VC++ Redist x64';      Id = 'Microsoft.VCRedist.2015+.x64'; Choco = 'vcredist140' }
    @{ Nome = 'VC++ Redist x86';      Id = 'Microsoft.VCRedist.2015+.x86'; Choco = 'vcredist140' }
    @{ Nome = '.NET Desktop Runtime 8'; Id = 'Microsoft.DotNet.DesktopRuntime.8'; Choco = 'dotnet-8.0-desktopruntime' }
    @{ Nome = 'Java (JRE)';           Id = 'Oracle.JavaRuntimeEnvironment'; Choco = 'javaruntime' }
    @{ Nome = 'LibreOffice';          Id = 'TheDocumentFoundation.LibreOffice'; Choco = 'libreoffice-fresh' }
    @{ Nome = 'TeamViewer';           Id = 'TeamViewer.TeamViewer';        Choco = 'teamviewer' }
    @{ Nome = 'Zoom';                 Id = 'Zoom.Zoom';                    Choco = 'zoom' }
    @{ Nome = 'WhatsApp';             Id = '9NKSQGP7F2NH'; Fonte = 'msstore' }
)

# Programas extras das opções "Gamer" do menu (somados à lista acima)
$ProgramasGamer = @(
    @{ Nome = 'Steam';                  Id = 'Valve.Steam';                  Choco = 'steam' }
    @{ Nome = 'Epic Games Launcher';    Id = 'EpicGames.EpicGamesLauncher';  Choco = 'epicgameslauncher' }
    @{ Nome = 'Discord';                Id = 'Discord.Discord';              Choco = 'discord' }
)

# ---------------- Microsoft Office ----------------
# No início o script mostra um menu para escolher a versão.
# Instalação oficial pela Office Deployment Tool, direto dos servidores da Microsoft.
# O Office é instalado sem licença: ative com a chave ou a conta Microsoft do cliente.
$VersoesOffice = [ordered]@{
    '1'  = @{ Nome = 'Microsoft 365 Apps para Empresas (completo)';   Produto = 'O365ProPlusRetail';  Canal = 'Current' }
    '2'  = @{ Nome = 'Microsoft 365 - só Word, Excel e PowerPoint';   Produto = 'O365ProPlusRetail';  Canal = 'Current'; Apps = @('Word', 'Excel', 'PowerPoint') }
    '3'  = @{ Nome = 'Microsoft 365 Business (Standard/Premium)';     Produto = 'O365BusinessRetail'; Canal = 'Current' }
    '4'  = @{ Nome = 'Microsoft 365 Personal / Family';               Produto = 'O365HomePremRetail'; Canal = 'Current' }
    '5'  = @{ Nome = 'Office 2024 Professional Plus';                 Produto = 'ProPlus2024Retail';  Canal = 'Current' }
    '6'  = @{ Nome = 'Office 2024 Home & Business';                   Produto = 'HomeBusiness2024Retail'; Canal = 'Current' }
    '7'  = @{ Nome = 'Office LTSC 2024 Professional Plus (Volume)';   Produto = 'ProPlus2024Volume';  Canal = 'PerpetualVL2024' }
    '8'  = @{ Nome = 'Office 2021 Professional Plus';                 Produto = 'ProPlus2021Retail';  Canal = 'Current' }
    '9'  = @{ Nome = 'Office LTSC 2021 Professional Plus (Volume)';   Produto = 'ProPlus2021Volume';  Canal = 'PerpetualVL2021' }
    '10' = @{ Nome = 'Office 2019 Professional Plus';                 Produto = 'ProPlus2019Retail';  Canal = 'Current' }
    '11' = @{ Nome = 'Office 2016 Professional Plus';                 Produto = 'ProPlusRetail';      Canal = 'Current' }
}
# "Apps" (opcional) = instala SÓ esses programas daquela versão; o resto é excluído.
$OfficePadrao      = '5'      # versão pré-selecionada ('0' = não instalar)
$OfficeIdioma      = 'pt-br'
$OfficeArquitetura = '64'     # '64' ou '32'
$OfficeMostrarJanela = $true  # mostra a janela de progresso do Office
# Apps que NÃO serão instalados junto com o Office.
# Possíveis: Access, Excel, Groove, Lync, OneDrive, OneNote, Outlook, PowerPoint, Publisher, Teams, Word, Bing
$OfficeExcluirApps = @('Groove', 'Lync', 'Teams', 'Bing')

# Instaladores que você tem no pendrive (antivírus, drivers, sistemas próprios...).
# Caminho relativo à pasta do script. .exe usa os argumentos informados;
# .msi já roda com /qn /norestart automaticamente.
$InstaladoresLocais = @(
    # @{ Nome = 'Antivírus';  Arquivo = 'Instaladores\antivirus.exe'; Argumentos = '/S' }
    # @{ Nome = 'Sistema X';  Arquivo = 'Instaladores\sistemax.msi';  Argumentos = '' }
)

# Opções do debloat
$RemoverXbox            = $false  # apps Xbox (alguns jogos da Store/Game Pass precisam deles)
$RemoverOneDrive        = $false  # desinstala o OneDrive e impede que volte em novos usuários
$MenuContextoClassico   = $true   # botão direito "antigo" do Windows 10 (só Win11)
$BarraTarefasEsquerda   = $false  # alinha o Iniciar à esquerda (só Win11)
$ExplorerAbreEsteComputador = $true
$MostrarExtensoes       = $true
$AtualizarTudoNoFinal   = $false  # roda "winget upgrade --all" no final
$UsarChocolatey         = $true   # se o winget não funcionar (ou falhar num programa), tenta pelo Chocolatey
$PreferirChocolatey     = $false  # $true = usa o Chocolatey primeiro e o winget como reserva

# Apps pré-instalados que serão removidos (aceita curingas *).
# O que NÃO está aqui não é tocado (Loja, Calculadora, Fotos, Bloco de Notas, Paint, etc.).
$AppsParaRemover = @(
    'Clipchamp.Clipchamp'
    'Microsoft.3DBuilder'
    'Microsoft.549981C3F5F10'               # Cortana
    'Microsoft.BingFinance'
    'Microsoft.BingFoodAndDrink'
    'Microsoft.BingHealthAndFitness'
    'Microsoft.BingNews'
    'Microsoft.BingSearch'
    'Microsoft.BingSports'
    'Microsoft.BingTranslator'
    'Microsoft.BingTravel'
    'Microsoft.BingWeather'
    'Microsoft.Copilot'
    'Microsoft.GetHelp'
    'Microsoft.Getstarted'                  # Dicas
    'Microsoft.Messaging'
    'Microsoft.Microsoft3DViewer'
    'Microsoft.MicrosoftJournal'
    'Microsoft.MicrosoftOfficeHub'          # app "Microsoft 365 (Office)"
    'Microsoft.MicrosoftSolitaireCollection'
    'Microsoft.MixedReality.Portal'
    'Microsoft.NetworkSpeedTest'
    'Microsoft.News'
    'Microsoft.Office.OneNote'
    'Microsoft.Office.Sway'
    'Microsoft.OneConnect'
    'Microsoft.OutlookForWindows'           # "Novo Outlook"
    'Microsoft.People'
    'Microsoft.PowerAutomateDesktop'
    'Microsoft.Print3D'
    'Microsoft.SkypeApp'
    'Microsoft.Todos'
    'Microsoft.Wallet'
    'Microsoft.Windows.DevHome'
    'Microsoft.WindowsFeedbackHub'
    'Microsoft.WindowsMaps'
    'Microsoft.ZuneVideo'                   # Filmes e TV
    'MicrosoftCorporationII.MicrosoftFamily'
    'MicrosoftTeams'                        # Teams pessoal (antigo)
    'MSTeams'                               # Teams pessoal (novo)
    # 'Microsoft.YourPhone'                 # Vincular ao Celular
    # 'Microsoft.WindowsAlarms'             # Alarmes e Relógio
    # 'Microsoft.ZuneMusic'                 # Media Player

    # Apps de terceiros que vêm "patrocinados"
    '*ACGMediaPlayer*'
    '*ActiproSoftwareLLC*'
    '*AdobePhotoshopExpress*'
    '*Amazon.com.Amazon*'
    '*AmazonVideo.PrimeVideo*'
    '*BubbleWitch3Saga*'
    '*CandyCrush*'
    '*Disney*'
    '*Duolingo*'
    '*Facebook*'
    '*Flipboard*'
    '*HiddenCity*'
    '*Instagram*'
    '*king.com.*'
    '*LinkedInforWindows*'
    '*Netflix*'
    '*PandoraMediaInc*'
    '*PicsArt*'
    '*Shazam*'
    '*SpotifyAB.SpotifyMusic*'
    '*TikTok*'
    '*Twitter*'
    '*WinZipUniversal*'
    '*XING*'
)

$AppsXbox = @(
    'Microsoft.GamingApp'
    'Microsoft.XboxApp'
    'Microsoft.Xbox.TCUI'
    'Microsoft.XboxGameOverlay'
    'Microsoft.XboxGamingOverlay'
    'Microsoft.XboxIdentityProvider'
    'Microsoft.XboxSpeechToTextOverlay'
)

# ---------------- Configurações do sistema ----------------
$IconesAreaTrabalho  = $true   # Este Computador, Lixeira, Painel de Controle e pasta do usuário
$ConfigurarRede      = $true   # rede privada + descoberta de rede + compartilhamento de arquivos/impressoras
$DesativarHibernacao = $true   # apaga o hiberfil.sys e libera espaço (só desktop; notebook mantém)

# ---------------- Limpeza (só pelo menu Ferramentas, não roda na instalação) ----------------
$LimpezaProfunda = $true       # inclui limpeza de componentes antigos do Windows (DISM, alguns minutos a mais)

# ---------------- Relatórios ----------------
$GerarRelatorioEntrega         = $true   # relatório HTML do serviço ao fim da instalação
$InstalarDrivers               = $true   # drivers do fabricante (Dell/Lenovo/HP) + Windows Update na instalação
$CopiarEntregaParaAreaTrabalho = $true   # deixa uma cópia na Área de Trabalho para o cliente
$NomeTecnico    = 'Danilo Américo'
$ContatoTecnico = ''                     # ex.: 'WhatsApp (81) 9xxxx-xxxx' - aparece nos relatórios
$MostrarChaveOEMCompleta = $false        # relatório mostra a chave OEM inteira (senão, só o final)

# =====================================================================
#  FIM DA CONFIGURAÇÃO
# =====================================================================

# --- Elevação automática para administrador (sempre no Windows PowerShell 5.1,
#     pois os comandos de Appx não funcionam direito no PowerShell 7) ---
$principal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
if (-not $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator) -or $PSVersionTable.PSEdition -eq 'Core') {
    $argList = @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', "`"$PSCommandPath`"")
    foreach ($k in $PSBoundParameters.Keys) {
        $v = $PSBoundParameters[$k]
        if ($v -is [System.Management.Automation.SwitchParameter]) { $argList += "-$k" }
        else { $argList += @("-$k", "`"$v`"") }
    }
    if ($PSBoundParameters.Count -eq 0) { Start-Process powershell.exe -Verb RunAs -WindowStyle Hidden -ArgumentList $argList }
    else                                { Start-Process powershell.exe -Verb RunAs -ArgumentList $argList }
    exit
}

$ErrorActionPreference = 'Continue'
$ProgressPreference    = 'SilentlyContinue'   # deixa os downloads bem mais rápidos
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
$Host.UI.RawUI.WindowTitle = 'SetupPC'

$PastaScript = $PSScriptRoot
$PastaLogs   = Join-Path $PastaScript 'logs'
New-Item -ItemType Directory -Path $PastaLogs -Force | Out-Null
$ArquivoLog  = Join-Path $PastaLogs ("SetupPC_{0}_{1}.log" -f $env:COMPUTERNAME, (Get-Date -Format 'yyyy-MM-dd_HH-mm'))
Start-Transcript -Path $ArquivoLog -Append | Out-Null

$Resultados = New-Object System.Collections.Generic.List[object]
$Acoes      = New-Object System.Collections.Generic.List[string]   # o que foi feito (vai para o relatório)
$PrecisaReiniciar = $false
$PastaRelatorios  = Join-Path $PastaScript 'relatorios'
$script:ModoNome  = 'Personalizada (linha de comando)'

# ---------------------------------------------------------------------
#  Funções utilitárias
# ---------------------------------------------------------------------
function Write-Etapa([string]$Texto) {
    Write-Host ''
    Write-Host ('=' * 70) -ForegroundColor DarkCyan
    Write-Host "  $Texto" -ForegroundColor Cyan
    Write-Host ('=' * 70) -ForegroundColor DarkCyan
}
# Write-Host "turbinado": além do console, manda cada linha para o log da interface gráfica
$script:Sync = $null
function Write-Host {
    param(
        [Parameter(Position = 0, ValueFromRemainingArguments = $true)] [object[]]$Object,
        [object]$ForegroundColor,
        [switch]$NoNewline
    )
    $txt = ($Object | ForEach-Object { "$_" }) -join ' '
    if ($script:Sync) { $script:Sync.Log.Enqueue([pscustomobject]@{ Texto = $txt; Cor = "$ForegroundColor" }) }
    $p = @{ Object = $txt }
    if ($ForegroundColor) { $p.ForegroundColor = $ForegroundColor }
    if ($NoNewline)       { $p.NoNewline = $true }
    Microsoft.PowerShell.Utility\Write-Host @p
}

# Progresso por etapas (barra da interface gráfica)
$script:PassoAtual  = 0
$script:TotalPassos = 0
function Step-Progresso([string]$Texto) {
    $script:PassoAtual++
    if ($script:Sync) {
        $script:Sync.Etapa       = $Texto
        $script:Sync.PassoAtual  = $script:PassoAtual
        $script:Sync.TotalPassos = $script:TotalPassos
    }
}

# Perguntas: no modo texto usa o console; na interface gráfica abre uma janela
function Invoke-PerguntaGui([string]$Tipo, [string]$Texto) {
    $script:Sync.Resposta   = $null
    $script:Sync.Respondida = $false
    $script:Sync.Pergunta   = @{ Tipo = $Tipo; Texto = $Texto }
    while (-not $script:Sync.Respondida -and -not $script:Sync.Fechado) { Start-Sleep -Milliseconds 100 }
    return $script:Sync.Resposta
}
function Confirm-Usuario([string]$Texto) {
    if ($script:Sync) { return [bool](Invoke-PerguntaGui 'SimNao' $Texto) }
    return ((Read-Host "   $Texto (S/N)") -match '^[sS]')
}
function Show-AvisoUsuario([string]$Texto) {
    if ($script:Sync) { [void](Invoke-PerguntaGui 'Aviso' $Texto); return }
    [void](Read-Host "   $($Texto -replace "`n", ' ') (Enter para continuar)")
}
function Read-TextoUsuario([string]$Texto) {
    if ($script:Sync) { return [string](Invoke-PerguntaGui 'Texto' $Texto) }
    return (Read-Host "   $Texto")
}

# Mostra/esconde a janela preta do console
function Set-Console([bool]$Visivel) {
    try {
        if (-not ('SetupPC.Janela' -as [type])) {
            Add-Type -Namespace SetupPC -Name Janela -MemberDefinition @'
[DllImport("kernel32.dll")] public static extern IntPtr GetConsoleWindow();
[DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
'@
        }
        $h = [SetupPC.Janela]::GetConsoleWindow()
        if ($h -ne [IntPtr]::Zero) { if ($Visivel) { [void][SetupPC.Janela]::ShowWindow($h, 5) } else { [void][SetupPC.Janela]::ShowWindow($h, 0) } }
    } catch {}
}

function Write-Ok([string]$t)    { Write-Host "  [OK]   $t" -ForegroundColor Green }
function Write-Info([string]$t)  { Write-Host "  [..]   $t" -ForegroundColor Gray }
function Write-Aviso([string]$t) { Write-Host "  [!!]   $t" -ForegroundColor Yellow }
function Write-Erro([string]$t)  { Write-Host "  [ERRO] $t" -ForegroundColor Red }

function Set-Reg {
    param([string]$Caminho, [string]$Nome, $Valor, [string]$Tipo = 'DWord')
    try {
        if (-not (Test-Path $Caminho)) { New-Item -Path $Caminho -Force | Out-Null }
        New-ItemProperty -Path $Caminho -Name $Nome -Value $Valor -PropertyType $Tipo -Force -ErrorAction Stop | Out-Null
    } catch {
        Write-Aviso "Não foi possível gravar $Caminho\$Nome ($($_.Exception.Message))"
    }
}

function Add-Acao([string]$Texto) { $script:Acoes.Add($Texto) }

function Format-Tamanho([double]$Bytes) {
    if ([math]::Abs($Bytes) -ge 1GB) { return '{0:N1} GB' -f ($Bytes / 1GB) }
    return '{0:N0} MB' -f ($Bytes / 1MB)
}

# ---------------------------------------------------------------------
#  1. Ponto de restauração
# ---------------------------------------------------------------------
function New-PontoRestauracao {
    Write-Etapa 'Criando ponto de restauração'
    try {
        Enable-ComputerRestore -Drive "$env:SystemDrive\" -ErrorAction Stop
        # Permite criar mais de um ponto em menos de 24h
        Set-Reg 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\SystemRestore' 'SystemRestorePointCreationFrequency' 0
        Checkpoint-Computer -Description 'Antes do SetupPC' -RestorePointType 'MODIFY_SETTINGS' -ErrorAction Stop
        Write-Ok 'Ponto de restauração criado.'
        Add-Acao 'Ponto de restauração criado antes das alterações'
    } catch {
        Write-Aviso "Não deu para criar o ponto de restauração: $($_.Exception.Message)"
    }
}

# ---------------------------------------------------------------------
#  2. Winget
# ---------------------------------------------------------------------
function Get-WingetPath {
    $cmd = Get-Command winget.exe -ErrorAction SilentlyContinue
    if ($cmd) { return $cmd.Source }
    $exe = Get-ChildItem "$env:ProgramFiles\WindowsApps\Microsoft.DesktopAppInstaller_*_x64__8wekyb3d8bbwe\winget.exe" -ErrorAction SilentlyContinue |
           Sort-Object FullName -Descending | Select-Object -First 1
    if ($exe) { return $exe.FullName }
    return $null
}

function Get-WingetVersao([string]$Exe) {
    try {
        $v = (& $Exe --version) 2>$null | Select-Object -First 1
        return [version]($v.Trim().TrimStart('v') -replace '-.*$', '')
    } catch { return [version]'0.0' }
}

function Install-WingetOnline {
    Write-Info 'Baixando o winget (App Installer) e dependências...'
    $tmp = Join-Path $env:TEMP 'SetupPC_winget'
    New-Item -ItemType Directory -Path $tmp -Force | Out-Null
    $arquivos = @(
        @{ Url = 'https://aka.ms/Microsoft.VCLibs.x64.14.00.Desktop.appx'; Nome = 'VCLibs.appx' }
        @{ Url = 'https://github.com/microsoft/microsoft-ui-xaml/releases/download/v2.8.6/Microsoft.UI.Xaml.2.8.x64.appx'; Nome = 'UIXaml.appx' }
        @{ Url = 'https://aka.ms/getwinget'; Nome = 'winget.msixbundle' }
    )
    foreach ($a in $arquivos) {
        $destino = Join-Path $tmp $a.Nome
        Invoke-WebRequest -Uri $a.Url -OutFile $destino -UseBasicParsing -ErrorAction Stop
        try { Add-AppxPackage -Path $destino -ErrorAction Stop }
        catch { Write-Info "$($a.Nome): $($_.Exception.Message.Split([Environment]::NewLine)[0])" }
    }
}

function Initialize-Winget {
    Write-Etapa 'Verificando o winget'
    $exe = Get-WingetPath

    if (-not $exe) {
        Write-Info 'winget não encontrado. Tentando registrar o App Installer...'
        try { Add-AppxPackage -RegisterByFamilyName -MainPackage Microsoft.DesktopAppInstaller_8wekyb3d8bbwe -ErrorAction Stop } catch {}
        for ($i = 0; $i -lt 12 -and -not $exe; $i++) { Start-Sleep -Seconds 5; $exe = Get-WingetPath }
    }

    if (-not $exe -or (Get-WingetVersao $exe) -lt [version]'1.7') {
        try { Install-WingetOnline } catch { Write-Erro "Falha ao baixar o winget: $($_.Exception.Message)" }
        Start-Sleep -Seconds 3
        $exe = Get-WingetPath
    }

    if (-not $exe) {
        Write-Erro 'winget indisponível. Verifique a internet ou atualize o "Instalador de Aplicativo" pela Microsoft Store.'
        return $null
    }

    Write-Ok "winget $(Get-WingetVersao $exe) pronto."
    & $exe source update --disable-interactivity 2>&1 | Out-Null
    return $exe
}

# ---------------------------------------------------------------------
#  3. Instalação dos programas
# ---------------------------------------------------------------------
# Códigos de saída do winget
$WG_JA_INSTALADO   = @(-1978335189, -1978335135)   # 0x8A15002B, 0x8A150061
$WG_REINICIAR      = @(-1978334967, 3010)          # 0x8A150109, ERROR_SUCCESS_REBOOT_REQUIRED

function Install-ProgramaWinget($Winget, $P) {
    $wargs = @('install', '--id', $P.Id, '--exact', '--silent',
               '--accept-package-agreements', '--accept-source-agreements', '--disable-interactivity')
    if ($P.Fonte) { $wargs += @('--source', $P.Fonte) } else { $wargs += @('--source', 'winget') }

    for ($tentativa = 1; $tentativa -le 2; $tentativa++) {
        $extra = @()
        # 1ª tentativa no idioma pedido; a 2ª sem idioma (caso não exista o instalador traduzido)
        if ($P.Locale -and $tentativa -eq 1) { $extra = @('--locale', $P.Locale) }
        Write-Info "Instalando $($P.Nome)$(if ($tentativa -gt 1) { ' (2ª tentativa)' })..."
        & $Winget @wargs @extra 2>&1 | Where-Object { $_ -match '\S' -and $_ -notmatch '^\s*[-\\|/]\s*$' -and $_ -notmatch '█|▒' } | Out-Host
        $code = $LASTEXITCODE

        if ($code -eq 0)                     { Write-Ok "$($P.Nome) instalado.";        return 'Instalado' }
        if ($WG_JA_INSTALADO -contains $code) { Write-Ok "$($P.Nome) já estava instalado."; return 'Já instalado' }
        if ($WG_REINICIAR -contains $code)    { Write-Ok "$($P.Nome) instalado (requer reinício)."; $script:PrecisaReiniciar = $true; return 'Instalado (reiniciar)' }
        Start-Sleep -Seconds 3
    }
    Write-Erro "$($P.Nome) falhou (código $code)."
    return "FALHOU ($code)"
}

function Install-InstaladorLocal($I) {
    $caminho = Join-Path $PastaScript $I.Arquivo
    if (-not (Test-Path $caminho)) { Write-Erro "Arquivo não encontrado: $caminho"; return 'Arquivo não encontrado' }

    Write-Info "Instalando $($I.Nome)..."
    try {
        if ([IO.Path]::GetExtension($caminho) -eq '.msi') {
            $proc = Start-Process msiexec.exe -ArgumentList "/i `"$caminho`" /qn /norestart $($I.Argumentos)" -Wait -PassThru -ErrorAction Stop
        } elseif ($I.Argumentos) {
            $proc = Start-Process $caminho -ArgumentList $I.Argumentos -Wait -PassThru -ErrorAction Stop
        } else {
            $proc = Start-Process $caminho -Wait -PassThru -ErrorAction Stop
        }
    } catch {
        Write-Erro "$($I.Nome): $($_.Exception.Message)"
        return 'FALHOU (não executou)'
    }
    switch ($proc.ExitCode) {
        0       { Write-Ok "$($I.Nome) instalado."; return 'Instalado' }
        3010    { Write-Ok "$($I.Nome) instalado (requer reinício)."; $script:PrecisaReiniciar = $true; return 'Instalado (reiniciar)' }
        default { Write-Erro "$($I.Nome) terminou com código $($proc.ExitCode)."; return "FALHOU ($($proc.ExitCode))" }
    }
}

function Select-Office {
    if ($SemProgramas) { return $null }
    $escolha = $Office
    if (-not $escolha -and $Silencioso) { $escolha = $OfficePadrao }

    if (-not $escolha) {
        Write-Etapa 'Qual Office instalar?'
        foreach ($k in $VersoesOffice.Keys) {
            Write-Host ("   {0,2}  {1}" -f $k, $VersoesOffice[$k].Nome)
        }
        Write-Host '    0  Não instalar o Office'
        Write-Host ''
        do {
            $escolha = (Read-Host "  Digite o número [Enter = $OfficePadrao]").Trim()
            if (-not $escolha) { $escolha = $OfficePadrao }
        } until ($escolha -eq '0' -or $VersoesOffice.Contains($escolha))
    }

    if ($escolha -eq '0' -or -not $VersoesOffice.Contains($escolha)) {
        Write-Info 'Office não será instalado.'
        return $null
    }
    Write-Ok "Office escolhido: $($VersoesOffice[$escolha].Nome)"
    return $VersoesOffice[$escolha]
}

function Install-Office($V) {
    Write-Etapa "Instalando $($V.Nome) ($OfficeIdioma, $OfficeArquitetura bits)"
    $pasta = Join-Path $env:TEMP 'SetupPC_Office'
    New-Item -ItemType Directory -Path $pasta -Force | Out-Null
    $setup = Join-Path $pasta 'setup.exe'

    # setup.exe da Office Deployment Tool (oficial da Microsoft).
    # Sem internet para baixar? Coloque um setup.exe em Instaladores\Office\ no pendrive.
    try {
        Invoke-WebRequest -Uri 'https://officecdn.microsoft.com/pr/wsus/setup.exe' -OutFile $setup -UseBasicParsing -ErrorAction Stop
    } catch {
        $local = Join-Path $PastaScript 'Instaladores\Office\setup.exe'
        if (Test-Path $local) {
            Copy-Item $local $setup -Force
            Write-Aviso 'Não deu para baixar a ODT; usando o setup.exe do pendrive.'
        } else {
            Write-Erro "Não foi possível baixar a Office Deployment Tool: $($_.Exception.Message)"
            return 'FALHOU (sem ODT)'
        }
    }

    $todosApps = 'Access', 'Excel', 'Groove', 'Lync', 'OneDrive', 'OneNote', 'Outlook', 'PowerPoint', 'Publisher', 'Teams', 'Word', 'Bing'
    if ($V.Apps) { $excluir = @($todosApps | Where-Object { $V.Apps -notcontains $_ }) }
    else         { $excluir = @($OfficeExcluirApps) }
    if ($RemoverOneDrive -and $excluir -notcontains 'OneDrive') { $excluir += 'OneDrive' }
    $linhasExcluir = ($excluir | ForEach-Object { "      <ExcludeApp ID=`"$_`" />" }) -join "`r`n"
    if ($OfficeMostrarJanela) { $display = 'Full' } else { $display = 'None' }

    $xml = @"
<Configuration>
  <Add OfficeClientEdition="$OfficeArquitetura" Channel="$($V.Canal)">
    <Product ID="$($V.Produto)">
      <Language ID="$OfficeIdioma" />
$linhasExcluir
    </Product>
  </Add>
  <Property Name="FORCEAPPSHUTDOWN" Value="TRUE" />
  <Updates Enabled="TRUE" />
  <RemoveMSI />
  <Display Level="$display" AcceptEULA="TRUE" />
</Configuration>
"@
    $xmlPath = Join-Path $pasta 'configuracao.xml'
    [IO.File]::WriteAllText($xmlPath, $xml)
    Write-Info "Configuração usada:`r`n$xml"
    Write-Info 'Baixando e instalando o Office (pode levar de 10 a 30 minutos, depende da internet)...'

    $proc = Start-Process $setup -ArgumentList "/configure `"$xmlPath`"" -Wait -PassThru
    if ($proc.ExitCode -eq 0) {
        Write-Ok "$($V.Nome) instalado."
        return 'Instalado'
    }
    Write-Erro "Instalação do Office terminou com código $($proc.ExitCode). Logs da ODT em: $env:TEMP"
    return "FALHOU ($($proc.ExitCode))"
}

function Get-ListaProgramas {
    # Na interface gráfica vale o que foi marcado; no modo texto, a lista padrão (+ gamer)
    if ($script:UsarSelecao) {
        return @(@($Programas) + @($ProgramasGamer) + @($ProgramasOpcionais) | Where-Object { $script:IdsSelecionados -contains $_.Id })
    }
    $l = @($Programas)
    if ($Gamer) { $l += $ProgramasGamer }
    return $l
}

# ---------------- Chocolatey (reserva do winget) ----------------
function Get-ChocoPath {
    $cmd = Get-Command choco.exe -ErrorAction SilentlyContinue
    if ($cmd) { return $cmd.Source }
    $exe = Join-Path $env:ProgramData 'chocolatey\bin\choco.exe'
    if (Test-Path $exe) { return $exe }
    return $null
}

function Initialize-Choco {
    $choco = Get-ChocoPath
    if ($choco) { return $choco }
    Write-Info 'Instalando o Chocolatey (gerenciador de programas reserva)...'
    try {
        [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor 3072
        $instalador = (New-Object System.Net.WebClient).DownloadString('https://community.chocolatey.org/install.ps1')
        & ([scriptblock]::Create($instalador)) *>&1 | Out-Null
    } catch {
        Write-Erro "Não foi possível instalar o Chocolatey: $($_.Exception.Message)"
        return $null
    }
    $choco = Get-ChocoPath
    if ($choco) { Write-Ok 'Chocolatey pronto.' } else { Write-Erro 'O Chocolatey não ficou disponível.' }
    return $choco
}

function Install-ProgramaChoco($Choco, $P) {
    $base = @('install', $P.Choco, '-y', '--no-progress', '--limit-output')
    if ($P.ChocoParams) { $base += @('--params', $P.ChocoParams) }

    $esperou = $false
    for ($tentativa = 1; $tentativa -le 2; $tentativa++) {
        $cargs = @($base)
        # 2ª tentativa: ignora a conferência do arquivo (o Google, por exemplo, troca o instalador
        # do Chrome no mesmo endereço e o pacote do Chocolatey fica com a "impressão digital" velha)
        if ($tentativa -eq 2) { $cargs += '--ignore-checksums' }
        if ($tentativa -eq 1) { Write-Info "Instalando $($P.Nome) pelo Chocolatey..." }
        else                  { Write-Info "Tentando de novo $($P.Nome) sem conferir o checksum..." }

        $saida = @(& $Choco @cargs 2>&1 | ForEach-Object { "$_" })
        $code = $LASTEXITCODE
        $saida | Out-Host

        if ($code -eq 0) {
            $txt = 'Instalado (Chocolatey)'
            if ($tentativa -eq 2) { $txt = 'Instalado (Chocolatey, sem checksum)' }
            Write-Ok "$($P.Nome): $txt."
            return $txt
        }
        if ($code -in 1641, 3010) {
            Write-Ok "$($P.Nome) instalado (Chocolatey, requer reinício)."
            $script:PrecisaReiniciar = $true
            return 'Instalado (Chocolatey, reiniciar)'
        }

        # Mostra no log da interface o motivo da falha
        $erros = @($saida | Where-Object { $_ -match 'ERROR|Error|error|failed|Failed|checksum|Checksum|1603|1618|denied' } | Select-Object -First 4)
        foreach ($linha in $erros) { Write-Aviso "  Chocolatey: $($linha.Trim())" }

        $ehChecksum = ($saida -join ' ') -match 'checksum'
        if ($tentativa -eq 1 -and -not $ehChecksum) {
            if (-not $esperou -and ($saida -join ' ') -match '1618|another installation') {
                Write-Info 'Outra instalação está em andamento; aguardando 30 segundos...'
                Start-Sleep -Seconds 30
                $esperou = $true
                $tentativa = 0; continue   # repete do jeito normal uma única vez
            }
            break
        }
    }
    Write-Erro "$($P.Nome) falhou no Chocolatey (código $code). Detalhes: $env:ProgramData\chocolatey\logs\chocolatey.log"
    return "FALHOU (Chocolatey $code)"
}

function Install-Programas {
    Write-Etapa 'Instalando programas'
    Step-Progresso 'Preparando o instalador de programas'
    $winget = $null
    $choco  = $null
    if ($PreferirChocolatey -and $UsarChocolatey) { $choco = Initialize-Choco }
    if (-not $choco) { $winget = Initialize-Winget }
    if (-not $winget -and -not $choco -and $UsarChocolatey) {
        Write-Aviso 'winget indisponível neste PC: usando o Chocolatey.'
        $choco = Initialize-Choco
    }

    foreach ($p in @(Get-ListaProgramas)) {
        Step-Progresso "Instalando $($p.Nome)"
        $status = 'FALHOU (sem instalador)'
        if ($PreferirChocolatey -and $choco -and $p.Choco) {
            $status = Install-ProgramaChoco $choco $p
            if ($status -like 'FALHOU*' -and -not $winget) { $winget = Initialize-Winget }
            if ($status -like 'FALHOU*' -and $winget) { $status = Install-ProgramaWinget $winget $p }
        } else {
            if ($winget) { $status = Install-ProgramaWinget $winget $p }
            # Reserva: se o winget falhar nesse programa, tenta pelo Chocolatey
            if ($status -like 'FALHOU*' -and $UsarChocolatey -and $p.Choco) {
                if (-not $choco) { $choco = Initialize-Choco }
                if ($choco) { $status = Install-ProgramaChoco $choco $p }
            }
        }
        if ($status -like 'FALHOU*' -and -not $p.Choco -and -not $winget) { $status = 'FALHOU (só existe no winget)' }
        $Resultados.Add([pscustomobject]@{ Programa = $p.Nome; Status = $status })
    }
    foreach ($i in $InstaladoresLocais) {
        Step-Progresso "Instalando $($i.Nome)"
        $status = Install-InstaladorLocal $i
        $Resultados.Add([pscustomobject]@{ Programa = $i.Nome; Status = $status })
    }

    if ($script:OfficeEscolhido) {
        Step-Progresso "Instalando $($script:OfficeEscolhido.Nome) (10 a 30 minutos)"
        $status = Install-Office $script:OfficeEscolhido
        $Resultados.Add([pscustomobject]@{ Programa = $script:OfficeEscolhido.Nome; Status = $status })
    }

    if ($AtualizarTudoNoFinal -and $winget) {
        Write-Info 'Atualizando todos os programas (winget upgrade --all)...'
        & $winget upgrade --all --silent --accept-package-agreements --accept-source-agreements --disable-interactivity | Out-Host
    }
}

# ---------------------------------------------------------------------
#  4. Debloat
# ---------------------------------------------------------------------
function Remove-AppsBloat {
    Write-Etapa 'Removendo apps pré-instalados'
    $lista = @($AppsParaRemover)
    if ($RemoverXbox) { $lista += $AppsXbox }

    $provisionados = Get-AppxProvisionedPackage -Online -ErrorAction SilentlyContinue
    $totalRemovidos = 0

    foreach ($nome in $lista) {
        $removeu = $false
        # Instalado para os usuários atuais
        $pkgs = Get-AppxPackage -AllUsers -Name $nome -ErrorAction SilentlyContinue
        foreach ($pkg in $pkgs) {
            try { Remove-AppxPackage -Package $pkg.PackageFullName -AllUsers -ErrorAction Stop; $removeu = $true }
            catch { try { Remove-AppxPackage -Package $pkg.PackageFullName -ErrorAction Stop; $removeu = $true } catch {} }
        }
        # Provisionado (evita que volte em novos usuários)
        foreach ($prov in ($provisionados | Where-Object { $_.DisplayName -like $nome })) {
            try { Remove-AppxProvisionedPackage -Online -PackageName $prov.PackageName -ErrorAction Stop | Out-Null; $removeu = $true } catch {}
        }
        if ($removeu) { Write-Ok "Removido: $nome"; $totalRemovidos++ }
    }
    if ($totalRemovidos -gt 0) { Add-Acao "Removidos $totalRemovidos aplicativos pré-instalados desnecessários" }

    # Impede o Windows Update de reinstalar o "Novo Outlook" e o Dev Home
    Remove-Item 'HKLM:\SOFTWARE\Microsoft\WindowsUpdate\Orchestrator\UScheduler_Oobe\OutlookUpdate' -Recurse -Force -ErrorAction SilentlyContinue
    Remove-Item 'HKLM:\SOFTWARE\Microsoft\WindowsUpdate\Orchestrator\UScheduler_Oobe\DevHomeUpdate' -Recurse -Force -ErrorAction SilentlyContinue
}

# Ajustes por usuário. $Raiz = 'HKCU:' ou a hive do usuário padrão.
function Set-AjustesUsuario([string]$Raiz) {
    $cdm = "$Raiz\Software\Microsoft\Windows\CurrentVersion\ContentDeliveryManager"
    foreach ($n in 'ContentDeliveryAllowed','OemPreInstalledAppsEnabled','PreInstalledAppsEnabled','PreInstalledAppsEverEnabled',
                   'SilentInstalledAppsEnabled','SoftLandingEnabled','SystemPaneSuggestionsEnabled','RotatingLockScreenOverlayEnabled',
                   'SubscribedContent-310093Enabled','SubscribedContent-338387Enabled','SubscribedContent-338388Enabled',
                   'SubscribedContent-338389Enabled','SubscribedContent-338393Enabled','SubscribedContent-353694Enabled',
                   'SubscribedContent-353696Enabled','SubscribedContent-353698Enabled') {
        Set-Reg $cdm $n 0
    }

    $adv = "$Raiz\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced"
    Set-Reg $adv 'ShowTaskViewButton'          0   # botão Visão de Tarefas
    Set-Reg $adv 'ShowCopilotButton'           0   # botão Copilot
    Set-Reg $adv 'TaskbarMn'                   0   # Chat/Teams na barra (Win11 antigo)
    Set-Reg $adv 'Start_IrisRecommendations'   0   # dicas/recomendações no Iniciar
    Set-Reg $adv 'Start_AccountNotifications'  0   # avisos de conta no Iniciar
    Set-Reg $adv 'ShowSyncProviderNotifications' 0 # anúncios no Explorer
    if ($MostrarExtensoes)            { Set-Reg $adv 'HideFileExt' 0 }
    if ($ExplorerAbreEsteComputador)  { Set-Reg $adv 'LaunchTo' 1 }
    if ($BarraTarefasEsquerda)        { Set-Reg $adv 'TaskbarAl' 0 }

    Set-Reg "$Raiz\Software\Microsoft\Windows\CurrentVersion\Search" 'SearchboxTaskbarMode' 1   # pesquisa só como ícone
    Set-Reg "$Raiz\Software\Microsoft\Windows\CurrentVersion\Search" 'BingSearchEnabled'    0
    Set-Reg "$Raiz\Software\Policies\Microsoft\Windows\Explorer"      'DisableSearchBoxSuggestions' 1  # sem Bing no Iniciar
    Set-Reg "$Raiz\Software\Microsoft\Windows\CurrentVersion\AdvertisingInfo" 'Enabled' 0
    Set-Reg "$Raiz\Software\Microsoft\Windows\CurrentVersion\Privacy" 'TailoredExperiencesWithDiagnosticDataEnabled' 0
    Set-Reg "$Raiz\Software\Microsoft\Windows\CurrentVersion\UserProfileEngagement" 'ScoobeSystemSettingEnabled' 0  # "Vamos concluir a configuração"
    Set-Reg "$Raiz\Software\Microsoft\Windows\CurrentVersion\Feeds"   'ShellFeedsTaskbarViewMode' 2  # Notícias e Interesses (Win10)
    Set-Reg "$Raiz\Software\Policies\Microsoft\Windows\WindowsCopilot" 'TurnOffWindowsCopilot' 1
    Set-Reg "$Raiz\Software\Policies\Microsoft\Windows\WindowsAI"     'DisableAIDataAnalysis' 1      # Recall

    if ($IconesAreaTrabalho) { Set-IconesDesktop $Raiz }
    if ($MenuContextoClassico) {
        Set-Reg "$Raiz\Software\Classes\CLSID\{86ca1aa0-34aa-4e8b-a509-50c905bae2a2}\InprocServer32" '(default)' '' 'String'
    }
    if ($RemoverOneDrive) {
        Remove-ItemProperty "$Raiz\Software\Microsoft\Windows\CurrentVersion\Run" -Name 'OneDriveSetup' -ErrorAction SilentlyContinue
        Remove-ItemProperty "$Raiz\Software\Microsoft\Windows\CurrentVersion\Run" -Name 'OneDrive' -ErrorAction SilentlyContinue
    }
}

function Set-AjustesSistema {
    Write-Etapa 'Desativando telemetria, anúncios, Copilot e Widgets'

    $pol = 'HKLM:\SOFTWARE\Policies\Microsoft'
    Set-Reg "$pol\Windows\DataCollection" 'AllowTelemetry' 0
    Set-Reg "$pol\Windows\DataCollection" 'DoNotShowFeedbackNotifications' 1
    Set-Reg "$pol\Windows\CloudContent"   'DisableWindowsConsumerFeatures' 1
    Set-Reg "$pol\Windows\CloudContent"   'DisableConsumerAccountStateContent' 1
    Set-Reg "$pol\Windows\AdvertisingInfo" 'DisabledByGroupPolicy' 1
    Set-Reg "$pol\Windows\System"          'EnableActivityFeed' 0
    Set-Reg "$pol\Windows\System"          'PublishUserActivities' 0
    Set-Reg "$pol\Windows\System"          'UploadUserActivities' 0
    Set-Reg "$pol\Dsh"                     'AllowNewsAndInterests' 0   # Widgets (Win11)
    Set-Reg "$pol\Windows\Windows Feeds"   'EnableFeeds' 0             # Notícias e Interesses (Win10)
    Set-Reg "$pol\Windows\WindowsCopilot"  'TurnOffWindowsCopilot' 1
    Set-Reg "$pol\Windows\WindowsAI"       'DisableAIDataAnalysis' 1   # Recall
    Set-Reg "$pol\Windows\Windows Search"  'AllowCortana' 0
    Set-Reg "$pol\Windows\Windows Search"  'DisableWebSearch' 1
    Write-Ok 'Políticas aplicadas.'

    # Serviço de telemetria
    foreach ($svc in 'DiagTrack', 'dmwappushservice') {
        try { Stop-Service $svc -Force -ErrorAction Stop } catch {}
        try { Set-Service  $svc -StartupType Disabled -ErrorAction Stop; Write-Ok "Serviço $svc desativado." } catch {}
    }

    # Tarefas agendadas de coleta de dados
    $tarefas = @(
        @{ P = '\Microsoft\Windows\Application Experience\';                N = 'Microsoft Compatibility Appraiser' }
        @{ P = '\Microsoft\Windows\Application Experience\';                N = 'ProgramDataUpdater' }
        @{ P = '\Microsoft\Windows\Customer Experience Improvement Program\'; N = 'Consolidator' }
        @{ P = '\Microsoft\Windows\Customer Experience Improvement Program\'; N = 'UsbCeip' }
        @{ P = '\Microsoft\Windows\Feedback\Siuf\';                         N = 'DmClient' }
        @{ P = '\Microsoft\Windows\Feedback\Siuf\';                         N = 'DmClientOnScenarioDownload' }
    )
    foreach ($t in $tarefas) {
        Disable-ScheduledTask -TaskPath $t.P -TaskName $t.N -ErrorAction SilentlyContinue | Out-Null
    }
    Write-Ok 'Tarefas de telemetria desativadas.'
}

function Set-AjustesTodosUsuarios {
    Write-Etapa 'Ajustando Iniciar, barra de tarefas e Explorer'

    Set-AjustesUsuario 'HKCU:'
    Write-Ok "Aplicado ao usuário atual ($env:USERNAME)."

    # Também aplica ao perfil padrão, para que usuários criados depois já venham limpos
    $hive = "$env:SystemDrive\Users\Default\NTUSER.DAT"
    if (Test-Path $hive) {
        & reg.exe load 'HKU\SetupPC_Default' $hive 2>&1 | Out-Null
        if ($LASTEXITCODE -eq 0) {
            Set-AjustesUsuario 'Registry::HKEY_USERS\SetupPC_Default'
            [gc]::Collect(); [gc]::WaitForPendingFinalizers(); Start-Sleep -Seconds 1
            & reg.exe unload 'HKU\SetupPC_Default' 2>&1 | Out-Null
            Write-Ok 'Aplicado ao perfil padrão (novos usuários).'
        } else {
            Write-Aviso 'Não foi possível carregar o perfil padrão.'
        }
    }
}

function Remove-OneDriveCompleto {
    Write-Etapa 'Removendo OneDrive'
    Get-Process OneDrive -ErrorAction SilentlyContinue | Stop-Process -Force
    $setups = @("$env:SystemRoot\System32\OneDriveSetup.exe", "$env:SystemRoot\SysWOW64\OneDriveSetup.exe")
    $exe = $setups | Where-Object { Test-Path $_ } | Select-Object -First 1
    if ($exe) {
        Start-Process $exe -ArgumentList '/uninstall' -Wait -NoNewWindow
    } else {
        $w = Get-WingetPath
        if ($w) { & $w uninstall --id Microsoft.OneDrive --silent --disable-interactivity --accept-source-agreements | Out-Null }
    }
    Set-Reg 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\OneDrive' 'DisableFileSyncNGSC' 1
    Write-Ok 'OneDrive removido.'
    Add-Acao 'OneDrive removido'
}

function Invoke-Debloat {
    Remove-AppsBloat
    Set-AjustesSistema
    Set-AjustesTodosUsuarios
    if ($RemoverOneDrive) { Remove-OneDriveCompleto }
    Add-Acao 'Telemetria, anúncios, sugestões, Copilot e Widgets desativados'
    Add-Acao 'Menu Iniciar, barra de tarefas e Explorador de Arquivos ajustados'

    Write-Info 'Reiniciando o Explorer para aplicar as mudanças...'
    Stop-Process -Name explorer -Force -ErrorAction SilentlyContinue
    Start-Sleep -Seconds 2
    if (-not (Get-Process explorer -ErrorAction SilentlyContinue)) { Start-Process explorer.exe }
}

# ---------------------------------------------------------------------
#  5. Telas e menus
# ---------------------------------------------------------------------
function Show-Banner {
    Clear-Host
    $c = 'Cyan'
    Write-Host ''
    Write-Host '  ╔══════════════════════════════════════════════════════════╗' -ForegroundColor $c
    Write-Host '  ║                                                          ║' -ForegroundColor $c
    Write-Host '  ║          S E T U P P C   -   Pós-formatação              ║' -ForegroundColor $c
    Write-Host '  ║                                                          ║' -ForegroundColor $c
    Write-Host '  ║             Script feito por Danilo Américo              ║' -ForegroundColor $c
    Write-Host '  ║                                                          ║' -ForegroundColor $c
    Write-Host '  ╚══════════════════════════════════════════════════════════╝' -ForegroundColor $c
    Write-Host ''
    $so = (Get-CimInstance Win32_OperatingSystem -ErrorAction SilentlyContinue).Caption
    Write-Host "   Computador: $env:COMPUTERNAME   |   Usuário: $env:USERNAME" -ForegroundColor DarkGray
    if ($so) { Write-Host "   Sistema:    $so" -ForegroundColor DarkGray }
    Write-Host "   Log:        $ArquivoLog" -ForegroundColor DarkGray
}

function Show-MenuPrincipal {
    Write-Host ''
    Write-Host '   INSTALAÇÃO' -ForegroundColor DarkCyan
    Write-Host '   [1]  Padrão                (programas + Office + debloat)'         -ForegroundColor White
    Write-Host '   [2]  Padrão sem Office     (programas + debloat)'                  -ForegroundColor White
    Write-Host '   [3]  Gamer                 (programas + jogos + Office + debloat)' -ForegroundColor White
    Write-Host '   [4]  Gamer sem Office      (programas + jogos + debloat)'          -ForegroundColor White
    Write-Host ''
    Write-Host '   FERRAMENTAS' -ForegroundColor DarkCyan
    Write-Host '   [5]  Otimização avançada   (desempenho, input lag, plano de energia)' -ForegroundColor White
    Write-Host '   [6]  Windows Update        (desativar / reativar)'                 -ForegroundColor White
    Write-Host '   [7]  Ativação              (Windows e Office com chave)'           -ForegroundColor White
    Write-Host '   [8]  Relatório do PC       (hardware, saúde do disco e bateria)'   -ForegroundColor White
    Write-Host '   [9]  Limpeza               (temporários, cache e arquivos antigos)' -ForegroundColor White
    Write-Host '   [10] Drivers               (verificar, instalar, Driver Booster)'  -ForegroundColor White
    Write-Host '   [0]  Sair'                                                         -ForegroundColor White
    Write-Host ''
    Write-Host "   Jogos: $(($ProgramasGamer | ForEach-Object { $_.Nome }) -join ', ')" -ForegroundColor DarkGray
    Write-Host ''
    do { $op = (Read-Host '   Escolha uma opção').Trim() } until ($op -in '0', '1', '2', '3', '4', '5', '6', '7', '8', '9', '10')
    return $op
}

# ---------------------------------------------------------------------
#  Windows Update (desativa sem quebrar a Microsoft Store)
# ---------------------------------------------------------------------
# Não desligamos o serviço "wuauserv": a Loja e o winget (fonte msstore) dependem dele.
# Em vez disso: política "sem atualizações automáticas" + pausa até daqui ~19 anos
# + trava na versão atual do Windows (não sobe para 24H2/25H2 etc.).
$WU_POL = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate'
$WU_UX  = 'HKLM:\SOFTWARE\Microsoft\WindowsUpdate\UX\Settings'
$WU_PAUSA_NOMES = 'PauseFeatureUpdatesStartTime','PauseFeatureUpdatesEndTime','PauseQualityUpdatesStartTime',
                  'PauseQualityUpdatesEndTime','PauseUpdatesStartTime','PauseUpdatesExpiryTime'

function Disable-WindowsUpdate {
    Write-Etapa 'Desativando o Windows Update (a Loja continua funcionando)'

    # 1) Política: não baixar/instalar atualizações automaticamente
    Set-Reg "$WU_POL\AU" 'NoAutoUpdate' 1
    Set-Reg "$WU_POL\AU" 'AUOptions'    2   # se algo verificar, só avisa; não baixa

    # 2) Trava na versão atual do Windows
    $cv = Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion'
    if ($cv.DisplayVersion) {
        if ([int]$cv.CurrentBuildNumber -ge 22000) { $produto = 'Windows 11' } else { $produto = 'Windows 10' }
        Set-Reg $WU_POL 'TargetReleaseVersion'     1
        Set-Reg $WU_POL 'ProductVersion'           $produto 'String'
        Set-Reg $WU_POL 'TargetReleaseVersionInfo' $cv.DisplayVersion 'String'
        Write-Ok "Versão travada em $produto $($cv.DisplayVersion)."
    }

    # 3) Pausa de atualizações por ~19 anos (funciona também no Windows Home)
    $agora = (Get-Date).ToUniversalTime()
    $ini = $agora.ToString('yyyy-MM-ddTHH:mm:ssZ')
    $fim = $agora.AddDays(7000).ToString('yyyy-MM-ddTHH:mm:ssZ')
    Set-Reg $WU_UX 'FlightSettingsMaxPauseDays' 7000
    foreach ($n in $WU_PAUSA_NOMES) {
        if ($n -like '*Start*') { Set-Reg $WU_UX $n $ini 'String' } else { Set-Reg $WU_UX $n $fim 'String' }
    }
    Write-Ok "Atualizações pausadas até $($agora.AddDays(7000).ToString('dd/MM/yyyy'))."

    Restart-Service wuauserv -Force -ErrorAction SilentlyContinue
    Write-Ok 'Windows Update desativado. Microsoft Store e apps da Loja continuam atualizando.'
    Add-Acao 'Windows Update desativado (Microsoft Store continua funcionando)'
    Write-Aviso 'Sem atualizações o PC não recebe correções de segurança. Para reverter use [6] > [2].'
}

function Enable-WindowsUpdate {
    Write-Etapa 'Reativando o Windows Update'
    foreach ($n in 'NoAutoUpdate', 'AUOptions') {
        Remove-ItemProperty "$WU_POL\AU" -Name $n -ErrorAction SilentlyContinue
    }
    foreach ($n in 'TargetReleaseVersion', 'ProductVersion', 'TargetReleaseVersionInfo') {
        Remove-ItemProperty $WU_POL -Name $n -ErrorAction SilentlyContinue
    }
    foreach ($n in $WU_PAUSA_NOMES + 'FlightSettingsMaxPauseDays') {
        Remove-ItemProperty $WU_UX -Name $n -ErrorAction SilentlyContinue
    }
    Set-Service wuauserv -StartupType Manual -ErrorAction SilentlyContinue
    Restart-Service wuauserv -Force -ErrorAction SilentlyContinue
    Write-Ok 'Windows Update reativado. Abra Configurações > Windows Update e clique em "Verificar".'
}

function Show-StatusWindowsUpdate {
    Write-Etapa 'Status do Windows Update'
    $noAuto = (Get-ItemProperty "$WU_POL\AU" -ErrorAction SilentlyContinue).NoAutoUpdate
    $pausa  = (Get-ItemProperty $WU_UX -ErrorAction SilentlyContinue).PauseUpdatesExpiryTime
    $trava  = Get-ItemProperty $WU_POL -ErrorAction SilentlyContinue
    if ($noAuto -eq 1 -or $pausa) {
        Write-Host '   Windows Update: DESATIVADO' -ForegroundColor Yellow
        if ($pausa) { Write-Host "   Pausado até:    $(([datetime]$pausa).ToString('dd/MM/yyyy'))" }
        if ($trava.TargetReleaseVersion -eq 1) { Write-Host "   Versão travada: $($trava.ProductVersion) $($trava.TargetReleaseVersionInfo)" }
    } else {
        Write-Host '   Windows Update: ATIVO' -ForegroundColor Green
    }
    $svc = Get-Service wuauserv -ErrorAction SilentlyContinue
    if ($svc) { Write-Host "   Serviço wuauserv: $($svc.Status) ($($svc.StartType))  - necessário para a Loja" -ForegroundColor DarkGray }
}

function Invoke-MenuWindowsUpdate {
    do {
        Write-Etapa 'Windows Update'
        Write-Host '   [1]  Desativar (mantém a Microsoft Store funcionando)'
        Write-Host '   [2]  Reativar'
        Write-Host '   [3]  Ver status'
        Write-Host '   [0]  Voltar'
        Write-Host ''
        $op = (Read-Host '   Escolha uma opção').Trim()
        switch ($op) {
            '1' { Disable-WindowsUpdate }
            '2' { Enable-WindowsUpdate }
            '3' { Show-StatusWindowsUpdate }
        }
        if ($op -ne '0') { Write-Host ''; Read-Host '   Enter para continuar' | Out-Null }
    } until ($op -eq '0')
}

# ---------------------------------------------------------------------
#  6. Ativação (com chave de licença legítima)
# ---------------------------------------------------------------------
$SLMGR = "$env:SystemRoot\System32\slmgr.vbs"

function Test-FormatoChave([string]$Chave) {
    return $Chave -match '^([A-Z0-9]{5}-){4}[A-Z0-9]{5}$'
}

function Read-Chave([string]$Texto) {
    do {
        $k = ("" + (Read-TextoUsuario "$Texto (XXXXX-XXXXX-XXXXX-XXXXX-XXXXX). Deixe vazio para cancelar.")).Trim().ToUpper()
        if (-not $k) { return $null }
        if (-not (Test-FormatoChave $k)) { Write-Aviso 'Formato de chave inválido.' }
    } until (Test-FormatoChave $k)
    return $k
}

function Get-OsppPath {
    $candidatos = @(
        "$env:ProgramFiles\Microsoft Office\Office16\OSPP.VBS"
        "${env:ProgramFiles(x86)}\Microsoft Office\Office16\OSPP.VBS"
    )
    return $candidatos | Where-Object { $_ -and (Test-Path $_) } | Select-Object -First 1
}

function Show-StatusAtivacao {
    Write-Etapa 'Status da ativação'
    $win = Get-CimInstance SoftwareLicensingProduct -Filter "ApplicationID='55c92734-d682-4d71-983e-d6ec3f16059f' AND PartialProductKey IS NOT NULL" -ErrorAction SilentlyContinue |
           Select-Object -First 1
    $estados = @{ 0 = 'Não licenciado'; 1 = 'ATIVADO'; 2 = 'Período de carência'; 3 = 'Carência adicional'; 4 = 'Não original'; 5 = 'Notificação'; 6 = 'Carência estendida' }
    if ($win) {
        $txt = $estados[[int]$win.LicenseStatus]
        $cor = if ($win.LicenseStatus -eq 1) { 'Green' } else { 'Yellow' }
        Write-Host "   Windows: $txt  ($($win.Name), chave final $($win.PartialProductKey))" -ForegroundColor $cor
    } else {
        Write-Host '   Windows: nenhuma chave instalada' -ForegroundColor Yellow
    }

    $ospp = Get-OsppPath
    if ($ospp) {
        Write-Host '   Office:' -ForegroundColor White
        & cscript.exe //nologo $ospp /dstatus | Where-Object { $_ -match 'LICENSE NAME|LICENSE STATUS|Last 5' } |
            ForEach-Object { Write-Host "     $_" }
    } else {
        Write-Host '   Office: não encontrado (ou é Microsoft 365, que ativa entrando com a conta)' -ForegroundColor DarkGray
    }
}

function Invoke-AtivarWindowsChave([string]$Chave) {
    Write-Info 'Instalando a chave do Windows...'
    & cscript.exe //nologo $SLMGR /ipk $Chave | ForEach-Object { Write-Host "   $_" }
    Write-Info 'Ativando online...'
    & cscript.exe //nologo $SLMGR /ato | ForEach-Object { Write-Host "   $_" }
}

function Invoke-AtivarWindowsOEM {
    # Notebooks/desktops de marca geralmente têm a licença gravada no firmware (BIOS/UEFI)
    $chave = (Get-CimInstance SoftwareLicensingService -ErrorAction SilentlyContinue).OA3xOriginalProductKey
    if (-not $chave) {
        Write-Aviso 'Este computador não tem chave do Windows gravada na BIOS.'
        return
    }
    Write-Ok "Chave OEM encontrada: *****-*****-*****-*****-$($chave.Substring($chave.Length - 5))"
    Invoke-AtivarWindowsChave $chave
}

function Invoke-AtivarOfficeChave {
    $ospp = Get-OsppPath
    if (-not $ospp) {
        Write-Aviso 'Office não encontrado. Se for Microsoft 365, basta abrir o Word e entrar com a conta do cliente.'
        return
    }
    $chave = Read-Chave 'Chave do Office'
    if (-not $chave) { return }
    Write-Info 'Instalando a chave do Office...'
    & cscript.exe //nologo $ospp "/inpkey:$chave" | ForEach-Object { Write-Host "   $_" }
    Write-Info 'Ativando online...'
    & cscript.exe //nologo $ospp /act | ForEach-Object { Write-Host "   $_" }
}

function Invoke-MenuAtivacao {
    do {
        Write-Etapa 'Ativação'
        Write-Host '   [1]  Ver status do Windows e do Office'
        Write-Host '   [2]  Ativar Windows com a chave gravada na BIOS (OEM)'
        Write-Host '   [3]  Ativar Windows digitando uma chave'
        Write-Host '   [4]  Ativar Office digitando uma chave'
        Write-Host '   [5]  Abrir a tela de Ativação do Windows'
        Write-Host '   [0]  Voltar'
        Write-Host ''
        $op = (Read-Host '   Escolha uma opção').Trim()
        switch ($op) {
            '1' { Show-StatusAtivacao }
            '2' { Invoke-AtivarWindowsOEM }
            '3' { $k = Read-Chave 'Chave do Windows'; if ($k) { Invoke-AtivarWindowsChave $k } }
            '4' { Invoke-AtivarOfficeChave }
            '5' { Start-Process 'ms-settings:activation' }
        }
        if ($op -ne '0') { Write-Host ''; Read-Host '   Enter para continuar' | Out-Null }
    } until ($op -eq '0')
}

# ---------------------------------------------------------------------
#  7. Otimização avançada (desempenho e input lag)
# ---------------------------------------------------------------------
# Tudo que é alterado aqui tem o valor original salvo em um backup,
# então a opção "Reverter" volta o PC exatamente como estava.
$OTIM_BACKUP   = "$env:ProgramData\SetupPC\otimizacao_backup.json"
$GUID_ULTIMATE = 'e9a42b02-d5df-448d-aa00-03f14749eb61'   # Desempenho Máximo (oculto)
$GUID_ALTO     = '8c5e7fda-e8bf-4a96-9a85-a6e23a8c635c'   # Alto desempenho
$GUID_EQUIL    = '381b4222-f694-41f0-9685-ff5bb260df2e'   # Equilibrado
$script:BackupOtim = @{}

function Get-Guid([string]$Texto) {
    return [regex]::Match($Texto, '[0-9a-fA-F]{8}-([0-9a-fA-F]{4}-){3}[0-9a-fA-F]{12}').Value
}

function Read-BackupOtim {
    $h = @{}
    if (Test-Path $OTIM_BACKUP) {
        try {
            $obj = Get-Content $OTIM_BACKUP -Raw | ConvertFrom-Json
            foreach ($p in $obj.PSObject.Properties) { $h[$p.Name] = $p.Value }
        } catch { Write-Aviso 'Backup de otimização corrompido; ignorando.' }
    }
    return $h
}

function Save-BackupOtim {
    New-Item -ItemType Directory -Path (Split-Path $OTIM_BACKUP) -Force | Out-Null
    $script:BackupOtim | ConvertTo-Json -Depth 4 | Set-Content -Path $OTIM_BACKUP -Encoding UTF8
}

# Igual ao Set-Reg, mas guarda o valor original (só na primeira vez)
function Set-RegOtim([string]$Caminho, [string]$Nome, $Valor, [string]$Tipo = 'DWord') {
    $chave = "$Caminho|$Nome"
    if (-not $script:BackupOtim.ContainsKey($chave)) {
        $reg = @{ Caminho = $Caminho; Nome = $Nome; Existia = $false; Valor = $null; Tipo = $null }
        $item = Get-Item -Path $Caminho -ErrorAction SilentlyContinue
        if ($item -and ($item.GetValueNames() -contains $Nome)) {
            $reg.Existia = $true
            $reg.Valor   = $item.GetValue($Nome, $null, 'DoNotExpandEnvironmentNames')
            $reg.Tipo    = $item.GetValueKind($Nome).ToString()
        }
        $script:BackupOtim[$chave] = $reg
    }
    Set-Reg $Caminho $Nome $Valor $Tipo
}

function Test-Notebook {
    return [bool](Get-CimInstance Win32_Battery -ErrorAction SilentlyContinue)
}

function Set-PlanoDesempenho([bool]$Notebook) {
    Write-Info 'Plano de energia...'
    if ($Notebook -and $script:PlanoNotebook) {
        if ($script:PlanoNotebook -ne 'S') { Write-Info 'Plano de energia mantido.'; return }
    } elseif ($Notebook -and -not $Silencioso) {
        if (-not (Confirm-Usuario 'Este PC é um notebook. O plano Desempenho Máximo gasta mais bateria e esquenta mais. Ativar mesmo assim?')) {
            Write-Info 'Plano de energia mantido.'; return
        }
    } elseif ($Notebook) {
        Write-Info 'Notebook detectado: plano de energia mantido (modo silencioso).'; return
    }

    $anterior = Get-Guid ((powercfg /getactivescheme) -join ' ')
    $lista    = (powercfg /list) -join ' '

    # Reaproveita o plano se o script já criou antes
    $plano = $script:BackupOtim['__plano']
    if ($plano -and $lista -match $plano.Criado) {
        powercfg /setactive $plano.Criado | Out-Null
        Write-Ok 'Plano "SetupPC - Desempenho Máximo" ativado.'
        return
    }

    $guid = Get-Guid ((powercfg -duplicatescheme $GUID_ULTIMATE 2>&1) -join ' ')
    if (-not $guid) {
        # Alguns PCs (Modern Standby) não aceitam o Desempenho Máximo; usa o Alto desempenho
        $guid = Get-Guid ((powercfg -duplicatescheme $GUID_ALTO 2>&1) -join ' ')
    }
    if (-not $guid) { Write-Aviso 'Não foi possível criar o plano de energia neste PC.'; return }

    powercfg -changename $guid 'SetupPC - Desempenho Máximo' 'Criado pelo SetupPC (Danilo Américo)' | Out-Null
    powercfg /setactive $guid | Out-Null
    $script:BackupOtim['__plano'] = @{ Anterior = $anterior; Criado = $guid }
    Write-Ok 'Plano "SetupPC - Desempenho Máximo" criado e ativado.'
}

function Invoke-Otimizacoes {
    Write-Etapa 'Otimização avançada'
    $script:BackupOtim = Read-BackupOtim
    $notebook = Test-Notebook

    # 1. Plano de energia: CPU sem economia de energia e sem estacionar núcleos
    Set-PlanoDesempenho $notebook

    # 2. Desliga a "inicialização rápida" (evita bugs de driver/USB/áudio após desligar)
    Set-RegOtim 'HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager\Power' 'HiberbootEnabled' 0
    Write-Ok 'Inicialização rápida desativada.'

    # 3. Sem "power throttling" (Windows reduzindo CPU de programas em 2º plano) - só desktop
    if (-not $notebook) {
        Set-RegOtim 'HKLM:\SYSTEM\CurrentControlSet\Control\Power\PowerThrottling' 'PowerThrottlingOff' 1
        Write-Ok 'Power throttling desativado.'
    }

    # 4. Agendador de GPU acelerado por hardware (precisa de GPU/driver compatível; senão é ignorado)
    Set-RegOtim 'HKLM:\SYSTEM\CurrentControlSet\Control\GraphicsDrivers' 'HwSchMode' 2
    Write-Ok 'Agendamento de GPU acelerado por hardware ativado.'

    # 5. Otimizações para jogos em janela (modelo "flip": menos latência em janela/sem borda)
    Set-RegOtim 'HKCU:\Software\Microsoft\DirectX\UserGpuPreferences' 'DirectXUserGlobalSettings' 'SwapEffectUpgradeEnable=1;' 'String'
    Write-Ok 'Otimização para jogos em janela ativada.'

    # 6. Modo de Jogo ligado; gravação em segundo plano (Game DVR) desligada
    Set-RegOtim 'HKCU:\Software\Microsoft\GameBar' 'AutoGameModeEnabled' 1
    Set-RegOtim 'HKCU:\Software\Microsoft\GameBar' 'AllowAutoGameMode' 1
    Set-RegOtim 'HKCU:\System\GameConfigStore' 'GameDVR_Enabled' 0
    Set-RegOtim 'HKCU:\Software\Microsoft\Windows\CurrentVersion\GameDVR' 'AppCaptureEnabled' 0
    Set-RegOtim 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\GameDVR' 'AllowGameDVR' 0
    Write-Ok 'Modo de Jogo ligado e gravação em segundo plano desligada.'

    # 7. Prioridade para jogos no agendador multimídia (MMCSS)
    $mm = 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Multimedia\SystemProfile'
    Set-RegOtim $mm 'SystemResponsiveness' 10                 # reserva 10% de CPU p/ 2º plano (padrão 20%)
    Set-RegOtim $mm 'NetworkThrottlingIndex' 0xffffffff       # sem limitação de rede durante multimídia
    Set-RegOtim "$mm\Tasks\Games" 'GPU Priority' 8
    Set-RegOtim "$mm\Tasks\Games" 'Priority' 6
    Set-RegOtim "$mm\Tasks\Games" 'Scheduling Category' 'High' 'String'
    Set-RegOtim "$mm\Tasks\Games" 'SFIO Priority' 'High' 'String'
    Write-Ok 'Prioridade de CPU/GPU para jogos ajustada.'

    # 8. Mais tempo de CPU para o programa em primeiro plano (o jogo)
    Set-RegOtim 'HKLM:\SYSTEM\CurrentControlSet\Control\PriorityControl' 'Win32PrioritySeparation' 38   # 0x26
    Write-Ok 'Prioridade do programa em primeiro plano aumentada.'

    # 9. Mouse sem aceleração ("Aumentar precisão do ponteiro" desligado) = mira consistente
    Set-RegOtim 'HKCU:\Control Panel\Mouse' 'MouseSpeed' '0' 'String'
    Set-RegOtim 'HKCU:\Control Panel\Mouse' 'MouseThreshold1' '0' 'String'
    Set-RegOtim 'HKCU:\Control Panel\Mouse' 'MouseThreshold2' '0' 'String'
    # Teclado: menor atraso e repetição mais rápida
    Set-RegOtim 'HKCU:\Control Panel\Keyboard' 'KeyboardDelay' '0' 'String'
    Set-RegOtim 'HKCU:\Control Panel\Keyboard' 'KeyboardSpeed' '31' 'String'
    # Sem a janela de "Teclas de Aderência" ao apertar Shift 5x durante o jogo
    Set-RegOtim 'HKCU:\Control Panel\Accessibility\StickyKeys' 'Flags' '506' 'String'
    Set-RegOtim 'HKCU:\Control Panel\Accessibility\ToggleKeys' 'Flags' '58' 'String'
    Set-RegOtim 'HKCU:\Control Panel\Accessibility\Keyboard Response' 'Flags' '122' 'String'
    Write-Ok 'Mouse sem aceleração, teclado mais rápido e atalhos de acessibilidade desligados.'

    # 10. Interface mais leve (menos animações e transparência)
    Set-RegOtim 'HKCU:\Control Panel\Desktop' 'MenuShowDelay' '50' 'String'
    Set-RegOtim 'HKCU:\Control Panel\Desktop\WindowMetrics' 'MinAnimate' '0' 'String'
    Set-RegOtim 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced' 'TaskbarAnimations' 0
    Set-RegOtim 'HKCU:\Software\Microsoft\Windows\DWM' 'EnableAeroPeek' 0
    Set-RegOtim 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize' 'EnableTransparency' 0
    Write-Ok 'Animações e transparência reduzidas.'

    # 11. Rede: desliga o algoritmo de Nagle (pacotes pequenos saem na hora = menos ping em jogos)
    $ifs = Get-ChildItem 'HKLM:\SYSTEM\CurrentControlSet\Services\Tcpip\Parameters\Interfaces' -ErrorAction SilentlyContinue
    foreach ($if in $ifs) {
        $p = Get-ItemProperty $if.PSPath -ErrorAction SilentlyContinue
        if ($p.DhcpIPAddress -or $p.IPAddress) {
            Set-RegOtim $if.PSPath 'TcpAckFrequency' 1
            Set-RegOtim $if.PSPath 'TCPNoDelay' 1
        }
    }
    Write-Ok 'Algoritmo de Nagle desativado nas placas de rede.'

    Save-BackupOtim
    Add-Acao 'Otimização de desempenho aplicada (plano de energia, prioridade para jogos, menor input lag)'
    Write-Ok "Backup dos valores originais salvo em $OTIM_BACKUP"
    Write-Aviso 'Reinicie o computador para todas as otimizações valerem.'
    $script:PrecisaReiniciar = $true
}

function Disable-VBS {
    Write-Etapa 'Integridade de Memória (VBS/HVCI) e Plataforma de Máquina Virtual'
    Write-Host '   A Microsoft indica desligar esses dois recursos para ganhar desempenho em jogos.' -ForegroundColor Gray
    Write-Host '   Em troca, o PC fica menos protegido contra drivers/malware que atacam o kernel.'  -ForegroundColor Gray
    Write-Host '   Também deixam de funcionar: WSL2, Subsistema Android e Windows Sandbox.'        -ForegroundColor Gray
    if (-not $Silencioso) {
        if (-not (Confirm-Usuario "Desativar a Integridade de Memória (VBS) e a Plataforma de Máquina Virtual?`n`nGanha desempenho em jogos, mas o PC fica menos protegido. WSL2, Subsistema Android e Windows Sandbox deixam de funcionar.")) { return }
    }
    $script:BackupOtim = Read-BackupOtim
    Set-RegOtim 'HKLM:\SYSTEM\CurrentControlSet\Control\DeviceGuard\Scenarios\HypervisorEnforcedCodeIntegrity' 'Enabled' 0
    Set-RegOtim 'HKLM:\SYSTEM\CurrentControlSet\Control\DeviceGuard' 'EnableVirtualizationBasedSecurity' 0
    Write-Ok 'Integridade de Memória desativada.'

    $vmp = Get-WindowsOptionalFeature -Online -FeatureName VirtualMachinePlatform -ErrorAction SilentlyContinue
    if ($vmp -and $vmp.State -eq 'Enabled') {
        Disable-WindowsOptionalFeature -Online -FeatureName VirtualMachinePlatform -NoRestart -ErrorAction SilentlyContinue | Out-Null
        $script:BackupOtim['__vmp'] = @{ EstavaAtivo = $true }
        Write-Ok 'Plataforma de Máquina Virtual desativada.'
    }
    Save-BackupOtim
    Write-Aviso 'Reinicie o computador para aplicar.'
}

function Undo-Otimizacoes {
    Write-Etapa 'Revertendo a otimização avançada'
    $b = Read-BackupOtim
    if ($b.Count -eq 0) { Write-Aviso 'Nenhum backup encontrado: nada para reverter.'; return }

    foreach ($k in @($b.Keys)) {
        if ($k -like '__*') { continue }
        $r = $b[$k]
        if ($r.Existia) {
            Set-Reg $r.Caminho $r.Nome $r.Valor $r.Tipo
        } else {
            Remove-ItemProperty -Path $r.Caminho -Name $r.Nome -ErrorAction SilentlyContinue
        }
    }
    Write-Ok 'Registro restaurado aos valores originais.'

    if ($b['__plano']) {
        $volta = $b['__plano'].Anterior
        if (-not $volta) { $volta = $GUID_EQUIL }
        powercfg /setactive $volta | Out-Null
        powercfg /delete $b['__plano'].Criado 2>&1 | Out-Null
        Write-Ok 'Plano de energia anterior restaurado.'
    }
    if ($b['__vmp']) {
        Enable-WindowsOptionalFeature -Online -FeatureName VirtualMachinePlatform -NoRestart -ErrorAction SilentlyContinue | Out-Null
        Write-Ok 'Plataforma de Máquina Virtual reativada.'
    }
    Remove-Item $OTIM_BACKUP -Force -ErrorAction SilentlyContinue
    Write-Aviso 'Reinicie o computador para concluir.'
}

function Invoke-MenuOtimizacao {
    do {
        Write-Etapa 'Otimização avançada'
        Write-Host '   [1]  Aplicar otimizações de desempenho e input lag (recomendado)'
        Write-Host '   [2]  Desativar Integridade de Memória/VBS  (+FPS, -segurança)'
        Write-Host '   [3]  Reverter tudo (volta como estava antes)'
        Write-Host '   [0]  Voltar'
        Write-Host ''
        $op = (Read-Host '   Escolha uma opção').Trim()
        switch ($op) {
            '1' { Invoke-Otimizacoes }
            '2' { Disable-VBS }
            '3' { Undo-Otimizacoes }
        }
        if ($op -ne '0') { Write-Host ''; Read-Host '   Enter para continuar' | Out-Null }
    } until ($op -eq '0')
}

# ---------------------------------------------------------------------
#  8. Configurações do sistema (ícones, rede, hibernação)
# ---------------------------------------------------------------------
function Set-IconesDesktop([string]$Raiz) {
    $k = "$Raiz\Software\Microsoft\Windows\CurrentVersion\Explorer\HideDesktopIcons\NewStartPanel"
    $icones = @(
        '{20D04FE0-3AEA-1069-A2D8-08002B30309D}'   # Este Computador
        '{645FF040-5081-101B-9F08-00AA002F954E}'   # Lixeira
        '{5399E694-6CE5-4D6C-8FCE-1D8870FDCBA0}'   # Painel de Controle
        '{59031a47-3f72-44a7-89c5-5595fe6b30ee}'   # Pasta do usuário
    )
    foreach ($id in $icones) { Set-Reg $k $id 0 }
}

function Set-ConfigRede {
    Write-Info 'Configurando a rede...'
    # Redes conectadas como "Privada" (permite ver outros PCs e impressoras)
    $perfis = Get-NetConnectionProfile -ErrorAction SilentlyContinue | Where-Object { "$($_.NetworkCategory)" -eq 'Public' }
    foreach ($p in $perfis) {
        try {
            Set-NetConnectionProfile -InterfaceIndex $p.InterfaceIndex -NetworkCategory Private -ErrorAction Stop
            Write-Ok "Rede '$($p.Name)' definida como Privada."
        } catch { Write-Aviso "Não foi possível alterar a rede '$($p.Name)'." }
    }
    # Regras do firewall pelo ID interno (funciona em qualquer idioma do Windows):
    #   -32752 = Descoberta de Rede   |   -28502 = Compartilhamento de Arquivo e Impressora
    foreach ($grupo in '@FirewallAPI.dll,-32752', '@FirewallAPI.dll,-28502') {
        Get-NetFirewallRule -Group $grupo -ErrorAction SilentlyContinue |
            Where-Object { $_.Profile.ToString() -match 'Private' } |
            Enable-NetFirewallRule -ErrorAction SilentlyContinue
    }
    # Serviços que fazem o PC aparecer na rede
    foreach ($svc in 'FDResPub', 'fdPHost') {
        Set-Service $svc -StartupType Automatic -ErrorAction SilentlyContinue
        Start-Service $svc -ErrorAction SilentlyContinue
    }
    foreach ($svc in 'SSDPSrv', 'upnphost') { Start-Service $svc -ErrorAction SilentlyContinue }
    Write-Ok 'Descoberta de rede e compartilhamento de arquivos e impressoras ativados (apenas em rede privada).'
    Add-Acao 'Rede configurada como privada, com descoberta de rede e compartilhamento de impressoras'
}

function Disable-Hibernacao {
    if (Test-Notebook) {
        Write-Info 'Notebook detectado: hibernação mantida (ela salva o trabalho quando a bateria acaba).'
        return
    }
    $arq = Get-Item "$env:SystemDrive\hiberfil.sys" -Force -ErrorAction SilentlyContinue
    $tam = 0; if ($arq) { $tam = $arq.Length }
    & powercfg.exe /hibernate off 2>&1 | Out-Null
    if ($LASTEXITCODE -eq 0) {
        $txt = 'Hibernação desativada'
        if ($tam -gt 0) { $txt += " ($(Format-Tamanho $tam) liberados)" }
        Write-Ok "$txt."
        Add-Acao $txt
    } else {
        Write-Aviso 'Não foi possível desativar a hibernação.'
    }
}

function Invoke-ConfiguracoesSistema {
    if (-not ($IconesAreaTrabalho -or $ConfigurarRede -or $DesativarHibernacao)) { return }
    Write-Etapa 'Configurações do sistema'
    if ($IconesAreaTrabalho) {
        Set-IconesDesktop 'HKCU:'
        Write-Ok 'Ícones adicionados à Área de Trabalho.'
        Add-Acao 'Ícones Este Computador, Lixeira, Painel de Controle e pasta do usuário na Área de Trabalho'
    }
    if ($ConfigurarRede)      { Set-ConfigRede }
    if ($DesativarHibernacao) { Disable-Hibernacao }
}

# ---------------------------------------------------------------------
#  9. Limpeza
# ---------------------------------------------------------------------
function Get-EspacoLivre {
    return (Get-PSDrive -Name $env:SystemDrive.TrimEnd(':')).Free
}

function Clear-Pasta([string]$Caminho) {
    if (Test-Path -LiteralPath $Caminho) {
        Get-ChildItem -LiteralPath $Caminho -Force -ErrorAction SilentlyContinue |
            Remove-Item -Recurse -Force -ErrorAction SilentlyContinue
    }
}

function Invoke-Limpeza {
    Write-Etapa 'Limpeza do sistema'
    $antes = Get-EspacoLivre

    # 1) Limpeza de Disco do Windows com todas as categorias, MENOS:
    #    pasta Downloads, arquivos de "Restaurar o PC" (ESD) e versões do Histórico de Arquivos
    $vc = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches'
    $preservar = @('DownloadsFolder', 'Windows ESD installation files', 'User file versions')
    $cleanmgr = "$env:SystemRoot\System32\cleanmgr.exe"
    if ((Test-Path $cleanmgr) -and (Test-Path $vc)) {
        Write-Info 'Executando a Limpeza de Disco do Windows (pode levar alguns minutos)...'
        Get-ChildItem $vc | Where-Object { $preservar -notcontains $_.PSChildName } | ForEach-Object {
            Set-ItemProperty -Path $_.PSPath -Name 'StateFlags0099' -Value 2 -Type DWord -ErrorAction SilentlyContinue
        }
        $proc = Start-Process $cleanmgr -ArgumentList '/sagerun:99' -PassThru -WindowStyle Hidden
        if (-not $proc.WaitForExit(1200000)) {   # limite de 20 minutos
            Stop-Process -Id $proc.Id -Force -ErrorAction SilentlyContinue
            Write-Aviso 'A Limpeza de Disco demorou demais e foi interrompida.'
        }
        Get-ChildItem $vc | ForEach-Object { Remove-ItemProperty -Path $_.PSPath -Name 'StateFlags0099' -ErrorAction SilentlyContinue }
    }

    # 2) Temporários e relatórios de erro
    Write-Info 'Apagando arquivos temporários...'
    Clear-Pasta $env:TEMP
    Clear-Pasta "$env:SystemRoot\Temp"
    Clear-Pasta "$env:ProgramData\Microsoft\Windows\WER\ReportArchive"
    Clear-Pasta "$env:ProgramData\Microsoft\Windows\WER\ReportQueue"
    Clear-Pasta "$env:LOCALAPPDATA\CrashDumps"

    # 3) Downloads antigos do Windows Update
    Write-Info 'Limpando o cache do Windows Update...'
    Stop-Service wuauserv, bits -Force -ErrorAction SilentlyContinue
    Clear-Pasta "$env:SystemRoot\SoftwareDistribution\Download"
    Start-Service bits, wuauserv -ErrorAction SilentlyContinue

    # 4) Cache da Otimização de Entrega
    if (Get-Command Delete-DeliveryOptimizationCache -ErrorAction SilentlyContinue) {
        Delete-DeliveryOptimizationCache -Force -ErrorAction SilentlyContinue | Out-Null
    }

    # 5) Lixeira
    Clear-RecycleBin -Force -ErrorAction SilentlyContinue

    # 6) Componentes antigos do Windows (WinSxS). Não usa /ResetBase, então
    #    as atualizações instaladas continuam podendo ser desinstaladas.
    if ($LimpezaProfunda) {
        Write-Info 'Limpando componentes antigos do Windows (DISM)... pode levar alguns minutos'
        & dism.exe /Online /Cleanup-Image /StartComponentCleanup /Quiet /NoRestart 2>&1 | Out-Null
    }

    $liberado = (Get-EspacoLivre) - $antes
    if ($liberado -lt 0) { $liberado = 0 }
    Write-Ok "Limpeza concluída: $(Format-Tamanho $liberado) liberados."
    Add-Acao "Limpeza do sistema: $(Format-Tamanho $liberado) liberados"
}

# ---------------------------------------------------------------------
#  10. Relatórios (HTML)
# ---------------------------------------------------------------------
function Esc($Texto) { return [System.Net.WebUtility]::HtmlEncode([string]$Texto) }

function Get-InfoPC {
    Write-Info 'Coletando informações do computador...'
    $cs   = Get-CimInstance Win32_ComputerSystem  -ErrorAction SilentlyContinue
    $os   = Get-CimInstance Win32_OperatingSystem -ErrorAction SilentlyContinue
    $bios = Get-CimInstance Win32_BIOS            -ErrorAction SilentlyContinue
    $mb   = Get-CimInstance Win32_BaseBoard       -ErrorAction SilentlyContinue
    $cpu  = Get-CimInstance Win32_Processor       -ErrorAction SilentlyContinue | Select-Object -First 1
    $cv   = Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion' -ErrorAction SilentlyContinue
    $notebook = Test-Notebook
    $info = [ordered]@{}

    # --- Computador ---
    if ($notebook) { $tipo = 'Notebook' } else { $tipo = 'Desktop' }
    $dataBios = ''; if ($bios.ReleaseDate) { $dataBios = " (de $($bios.ReleaseDate.ToString('dd/MM/yyyy')))" }
    $info['Computador'] = [ordered]@{
        'Nome'            = $env:COMPUTERNAME
        'Tipo'            = $tipo
        'Fabricante'      = $cs.Manufacturer
        'Modelo'          = $cs.Model
        'Número de série' = $bios.SerialNumber
        'Placa-mãe'       = ("$($mb.Manufacturer) $($mb.Product)").Trim()
        'BIOS'            = "$($bios.SMBIOSBIOSVersion)$dataBios"
    }

    # --- Sistema ---
    $lic = Get-CimInstance SoftwareLicensingProduct -Filter "ApplicationID='55c92734-d682-4d71-983e-d6ec3f16059f' AND PartialProductKey IS NOT NULL" -ErrorAction SilentlyContinue |
           Select-Object -First 1
    if ($lic -and $lic.LicenseStatus -eq 1) { $ativ = 'Ativado' } elseif ($lic) { $ativ = 'Não ativado' } else { $ativ = 'Sem chave instalada' }
    $oem = (Get-CimInstance SoftwareLicensingService -ErrorAction SilentlyContinue).OA3xOriginalProductKey
    if ($oem -and -not $MostrarChaveOEMCompleta) { $oem = "*****-*****-*****-*****-$($oem.Substring($oem.Length - 5))" }
    if (-not $oem) { $oem = 'Não possui' }
    $instalado = ''; if ($os.InstallDate) { $instalado = $os.InstallDate.ToString('dd/MM/yyyy HH:mm') }
    $info['Sistema'] = [ordered]@{
        'Windows'          = $os.Caption
        'Versão'           = "$($cv.DisplayVersion) (build $($os.BuildNumber).$($cv.UBR))"
        'Arquitetura'      = $os.OSArchitecture
        'Instalado em'     = $instalado
        'Ativação'         = $ativ
        'Chave OEM (BIOS)' = $oem
    }

    # --- Processador ---
    $info['Processador'] = [ordered]@{
        'Modelo'          = ("$($cpu.Name)").Trim()
        'Núcleos'         = "$($cpu.NumberOfCores) núcleos / $($cpu.NumberOfLogicalProcessors) threads"
        'Frequência máx.' = '{0:N2} GHz' -f ($cpu.MaxClockSpeed / 1000)
    }

    # --- Memória ---
    $tiposRam = @{ 20 = 'DDR'; 21 = 'DDR2'; 24 = 'DDR3'; 26 = 'DDR4'; 34 = 'DDR5'; 35 = 'LPDDR5' }
    $pentes = @(Get-CimInstance Win32_PhysicalMemory -ErrorAction SilentlyContinue)
    $slots  = (Get-CimInstance Win32_PhysicalMemoryArray -ErrorAction SilentlyContinue | Measure-Object MemoryDevices -Sum).Sum
    $tipoRam = ''; if ($pentes.Count -gt 0) { $tipoRam = $tiposRam[[int]$pentes[0].SMBIOSMemoryType] }
    $info['Memória RAM'] = [ordered]@{
        'Total'        = ('{0:N0} GB {1}' -f ($cs.TotalPhysicalMemory / 1GB), $tipoRam).Trim()
        'Slots usados' = "$($pentes.Count) de $slots"
    }
    $info['Pentes de memória'] = @($pentes | ForEach-Object {
        $vel = $_.ConfiguredClockSpeed; if (-not $vel) { $vel = $_.Speed }
        [pscustomobject]@{
            'Slot'       = $_.DeviceLocator
            'Capacidade' = '{0:N0} GB' -f ($_.Capacity / 1GB)
            'Velocidade' = "$vel MHz"
            'Fabricante' = ("$($_.Manufacturer)").Trim()
        }
    })

    # --- Vídeo ---
    $info['Placa de vídeo'] = @(Get-CimInstance Win32_VideoController -ErrorAction SilentlyContinue | ForEach-Object {
        $res = ''
        if ($_.CurrentHorizontalResolution) { $res = "$($_.CurrentHorizontalResolution) x $($_.CurrentVerticalResolution) @ $($_.CurrentRefreshRate) Hz" }
        [pscustomobject]@{ 'Modelo' = $_.Name; 'Driver' = $_.DriverVersion; 'Resolução' = $res }
    })

    # --- Discos (com saúde SMART quando disponível) ---
    $saudeTxt = @{ 'Healthy' = 'Boa'; 'Warning' = 'Atenção'; 'Unhealthy' = 'Ruim' }
    $info['Discos'] = @(Get-PhysicalDisk -ErrorAction SilentlyContinue | Sort-Object DeviceId | ForEach-Object {
        $rel = $_ | Get-StorageReliabilityCounter -ErrorAction SilentlyContinue
        $midia = "$($_.MediaType)"; if ($midia -eq 'Unspecified') { $midia = '' }
        $temp  = ''; if ($rel.Temperature)     { $temp  = "$($rel.Temperature) °C" }
        $desg  = ''; if ($null -ne $rel.Wear)  { $desg  = "$($rel.Wear)%" }
        $horas = ''; if ($rel.PowerOnHours)    { $horas = '{0:N0} h' -f $rel.PowerOnHours }
        $saude = $saudeTxt["$($_.HealthStatus)"]; if (-not $saude) { $saude = "$($_.HealthStatus)" }
        [pscustomobject]@{
            'Disco'        = $_.FriendlyName
            'Tipo'         = ("$midia $($_.BusType)").Trim()
            'Tamanho'      = (Format-Tamanho $_.Size)
            'Saúde'        = $saude
            'Temperatura'  = $temp
            'Desgaste'     = $desg
            'Tempo ligado' = $horas
        }
    })
    $info['Partições'] = @(Get-Volume -ErrorAction SilentlyContinue |
        Where-Object { $_.DriveLetter -and "$($_.DriveType)" -eq 'Fixed' } | Sort-Object DriveLetter | ForEach-Object {
        $pct = 0; if ($_.Size) { $pct = [math]::Round($_.SizeRemaining / $_.Size * 100) }
        [pscustomobject]@{
            'Unidade' = "$($_.DriveLetter):"
            'Nome'    = $_.FileSystemLabel
            'Sistema' = $_.FileSystem
            'Tamanho' = (Format-Tamanho $_.Size)
            'Livre'   = "$(Format-Tamanho $_.SizeRemaining) ($pct%)"
        }
    })

    # --- Rede ---
    $statusRede = @{ 'Up' = 'Conectado'; 'Disconnected' = 'Desconectado'; 'Disabled' = 'Desativado' }
    $info['Rede'] = @(Get-NetAdapter -Physical -ErrorAction SilentlyContinue | Sort-Object Name | ForEach-Object {
        $st = $statusRede["$($_.Status)"]; if (-not $st) { $st = "$($_.Status)" }
        [pscustomobject]@{ 'Adaptador' = $_.InterfaceDescription; 'Status' = $st; 'Velocidade' = $_.LinkSpeed; 'MAC' = $_.MacAddress }
    })

    # --- Bateria (só notebook) ---
    if ($notebook) {
        $xmlPath = Join-Path $env:TEMP 'SetupPC_bateria.xml'
        & powercfg.exe /batteryreport /xml /output $xmlPath 2>&1 | Out-Null
        if (Test-Path $xmlPath) {
            try {
                [xml]$x = Get-Content $xmlPath -Raw
                $info['Bateria'] = @($x.BatteryReport.Batteries.Battery | ForEach-Object {
                    $proj = [double]$_.DesignCapacity; $atual = [double]$_.FullChargeCapacity
                    $pct = ''; if ($proj -gt 0) { $pct = '{0:N0}%' -f ($atual / $proj * 100) }
                    [pscustomobject]@{
                        'Bateria'             = ("$($_.Manufacturer) $($_.Id)").Trim()
                        'Capacidade original' = '{0:N0} mWh' -f $proj
                        'Capacidade atual'    = '{0:N0} mWh' -f $atual
                        'Saúde'               = $pct
                        'Ciclos'              = $_.CycleCount
                    }
                })
            } catch { Write-Aviso 'Não foi possível ler o relatório da bateria.' }
            Remove-Item $xmlPath -Force -ErrorAction SilentlyContinue
        }
    }

    # --- Segurança ---
    $tpmTxt = 'Não encontrado'
    $tpm = Get-Tpm -ErrorAction SilentlyContinue
    if ($tpm -and $tpm.TpmPresent) {
        $tpmTxt = 'Presente'
        $ver = (Get-CimInstance -Namespace 'root\cimv2\security\microsofttpm' -ClassName Win32_Tpm -ErrorAction SilentlyContinue).SpecVersion
        if ($ver) { $tpmTxt += " (versão $((($ver -split ',')[0]).Trim()))" }
    }
    try { if (Confirm-SecureBootUEFI -ErrorAction Stop) { $sboot = 'Ativado' } else { $sboot = 'Desativado' } }
    catch { $sboot = 'Não suportado (BIOS legado)' }
    $bl = 'Não disponível nesta edição'
    try {
        $v = Get-BitLockerVolume -MountPoint $env:SystemDrive -ErrorAction Stop
        if ("$($v.ProtectionStatus)" -eq 'On')   { $bl = "Ativado ($($v.EncryptionPercentage)% criptografado)" }
        elseif ($v.EncryptionPercentage -gt 0)    { $bl = 'Criptografado, proteção suspensa' }
        else                                      { $bl = 'Desativado' }
    } catch {}
    $av = (Get-CimInstance -Namespace 'root/SecurityCenter2' -ClassName AntiVirusProduct -ErrorAction SilentlyContinue |
           ForEach-Object { $_.displayName }) -join ', '
    $info['Dispositivos sem driver ou com erro'] = @(Get-DispositivosSemDriver)

    $info['Segurança'] = [ordered]@{
        'TPM'                       = $tpmTxt
        'Secure Boot'               = $sboot
        'BitLocker (disco sistema)' = $bl
        'Antivírus'                 = $av
    }
    return $info
}

# Cor da etiqueta conforme o conteúdo (verde / amarelo / vermelho)
function Get-ClasseValor([string]$Coluna, [string]$Valor) {
    if ($Coluna -notmatch 'Status|Saúde|Ativação|Secure Boot|TPM') { return '' }
    if ($Valor -match '^(\d+)%$') {
        $n = [int]$Matches[1]
        if ($n -ge 80) { return 'ok' } elseif ($n -ge 50) { return 'aviso' } else { return 'ruim' }
    }
    if ($Valor -match 'FALHOU|Ruim|Não ativado|Sem chave|Não encontrado') { return 'ruim' }
    if ($Valor -match 'Atenção|reiniciar|Desativado|Desconectado|Não suportado') { return 'aviso' }
    if ($Valor -match 'Instalado|Boa|Ativado|Conectado|Presente') { return 'ok' }
    return ''
}

function Format-Celula([string]$Coluna, [string]$Valor) {
    $cls = Get-ClasseValor $Coluna $Valor
    if ($cls -and $Valor) { return "<span class=`"tag $cls`">$(Esc $Valor)</span>" }
    return (Esc $Valor)
}

function ConvertTo-HtmlSecoes($Secoes) {
    $sb = New-Object System.Text.StringBuilder
    foreach ($titulo in $Secoes.Keys) {
        $dados = $Secoes[$titulo]
        [void]$sb.Append("<section><h2>$(Esc $titulo)</h2>")
        if ($dados -is [System.Collections.IDictionary]) {
            [void]$sb.Append('<table class="kv">')
            foreach ($k in $dados.Keys) {
                $v = [string]$dados[$k]
                if (-not $v.Trim()) { continue }
                [void]$sb.Append("<tr><th>$(Esc $k)</th><td>$(Format-Celula $k $v)</td></tr>")
            }
            [void]$sb.Append('</table>')
        } else {
            $itens = @($dados)
            if ($itens.Count -eq 0) {
                [void]$sb.Append('<p class="vazio">Nenhum item.</p>')
            } elseif ($itens[0] -is [string]) {
                [void]$sb.Append('<ul>')
                foreach ($i in $itens) { [void]$sb.Append("<li>$(Esc $i)</li>") }
                [void]$sb.Append('</ul>')
            } else {
                $cols = @($itens[0].PSObject.Properties.Name)
                [void]$sb.Append('<table class="grade"><tr>')
                foreach ($c in $cols) { [void]$sb.Append("<th>$(Esc $c)</th>") }
                [void]$sb.Append('</tr>')
                foreach ($i in $itens) {
                    [void]$sb.Append('<tr>')
                    foreach ($c in $cols) { [void]$sb.Append("<td>$(Format-Celula $c ([string]$i.$c))</td>") }
                    [void]$sb.Append('</tr>')
                }
                [void]$sb.Append('</table>')
            }
        }
        [void]$sb.Append('</section>')
    }
    return $sb.ToString()
}

function New-PaginaHtml([string]$Titulo, [string]$Subtitulo, [string]$Corpo) {
    $data = Get-Date -Format 'dd/MM/yyyy HH:mm'
    $contato = ''; if ($ContatoTecnico) { $contato = "<br>$(Esc $ContatoTecnico)" }
    return @"
<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>$(Esc $Titulo) - $(Esc $env:COMPUTERNAME)</title>
<style>
*{box-sizing:border-box}
body{margin:0;font-family:"Segoe UI",system-ui,sans-serif;background:#f3f5f8;color:#1d2433;font-size:14px}
header{background:#14306b;color:#fff;padding:28px 40px;display:flex;justify-content:space-between;align-items:flex-end;gap:24px;flex-wrap:wrap}
header h1{margin:0;font-size:24px;font-weight:600}
header p{margin:6px 0 0;opacity:.8}
.meta{text-align:right;font-size:13px;opacity:.9;line-height:1.6}
main{max-width:1000px;margin:24px auto;padding:0 20px;display:grid;gap:16px}
section{background:#fff;border:1px solid #e1e5ec;border-radius:10px;padding:18px 22px;overflow-x:auto}
h2{margin:0 0 12px;font-size:13px;color:#14306b;text-transform:uppercase;letter-spacing:.06em}
table{width:100%;border-collapse:collapse}
.kv th{width:34%;text-align:left;font-weight:500;color:#5a6478;padding:7px 12px 7px 0;vertical-align:top}
.kv td{padding:7px 0}
.kv tr+tr th,.kv tr+tr td{border-top:1px solid #eef1f5}
.grade th{text-align:left;font-weight:600;font-size:12px;color:#5a6478;border-bottom:2px solid #e1e5ec;padding:8px 12px 8px 0;white-space:nowrap}
.grade td{padding:8px 12px 8px 0;border-bottom:1px solid #eef1f5}
.grade tr:last-child td{border-bottom:0}
ul{margin:0;padding-left:20px;line-height:1.9}
.tag{display:inline-block;padding:2px 10px;border-radius:999px;font-size:12px;font-weight:600;white-space:nowrap}
.ok{background:#e3f6ea;color:#16703a}
.aviso{background:#fff3d6;color:#8a5a00}
.ruim{background:#fde4e4;color:#a12020}
.vazio{color:#8a93a5;margin:0}
footer{text-align:center;color:#8a93a5;font-size:12px;padding:8px 0 32px}
@media print{body{background:#fff}main{margin:0;max-width:none}section{break-inside:avoid;border-color:#ccc}header{-webkit-print-color-adjust:exact;print-color-adjust:exact}}
</style>
</head>
<body>
<header>
  <div><h1>$(Esc $Titulo)</h1><p>$(Esc $Subtitulo)</p></div>
  <div class="meta">$data<br>Técnico: $(Esc $NomeTecnico)$contato</div>
</header>
<main>
$Corpo
</main>
<footer>Gerado pelo SetupPC &middot; Script feito por $(Esc $NomeTecnico)</footer>
</body>
</html>
"@
}

function Save-Html([string]$Caminho, [string]$Html) {
    New-Item -ItemType Directory -Path (Split-Path $Caminho) -Force | Out-Null
    [IO.File]::WriteAllText($Caminho, $Html, (New-Object System.Text.UTF8Encoding($true)))
}

function New-RelatorioPC {
    Write-Etapa 'Relatório do computador'
    $info = Get-InfoPC
    $sub  = ("$($info['Computador']['Fabricante']) $($info['Computador']['Modelo'])").Trim()
    $html = New-PaginaHtml 'Relatório do Computador' $sub (ConvertTo-HtmlSecoes $info)
    $arq  = Join-Path $PastaRelatorios ('PC_{0}_{1}.html' -f $env:COMPUTERNAME, (Get-Date -Format 'yyyy-MM-dd_HH-mm'))
    Save-Html $arq $html
    Write-Ok "Relatório salvo em: $arq"
    if (-not $Silencioso) { Start-Process $arq }
}

function New-RelatorioEntrega([TimeSpan]$Duracao) {
    Write-Etapa 'Relatório de entrega'
    $secoes = [ordered]@{}

    $officeTxt = 'Não instalado'
    if ($script:OfficeEscolhido) { $officeTxt = $script:OfficeEscolhido.Nome }
    $secoes['Serviço realizado'] = [ordered]@{
        'Data'               = (Get-Date -Format 'dd/MM/yyyy HH:mm')
        'Técnico'            = $NomeTecnico
        'Contato'            = $ContatoTecnico
        'Tipo de instalação' = $script:ModoNome
        'Microsoft Office'   = $officeTxt
        'Duração'            = '{0} min' -f [math]::Max(1, [math]::Round($Duracao.TotalMinutes))
    }
    $secoes['Programas instalados'] = @($Resultados | ForEach-Object {
        [pscustomobject]@{ 'Programa' = $_.Programa; 'Status' = $_.Status }
    })
    $secoes['O que foi feito'] = @($script:Acoes)
    if ($script:ResultadoDrivers) { $secoes['Drivers'] = @($script:ResultadoDrivers) }

    $info = Get-InfoPC
    foreach ($k in $info.Keys) { $secoes[$k] = $info[$k] }

    $recs = New-Object System.Collections.Generic.List[string]
    $recs.Add('Reinicie o computador para concluir todas as configurações.')
    if ($script:OfficeEscolhido) {
        $recs.Add('Ative o Microsoft Office com a sua chave de licença ou entrando com a sua conta Microsoft no Word.')
    }
    if ((Get-ItemProperty "$WU_POL\AU" -ErrorAction SilentlyContinue).NoAutoUpdate -eq 1) {
        $recs.Add('O Windows Update está desativado: o computador não recebe atualizações de segurança automaticamente.')
    }
    if (@($script:DispositivosPendentes).Count -gt 0) {
        $recs.Add('Alguns dispositivos ficaram sem driver: veja a lista "Dispositivos sem driver ou com erro" neste relatório.')
    }
    $recs.Add('Mantenha cópias dos seus arquivos importantes em um HD externo ou na nuvem.')
    $secoes['Recomendações'] = @($recs)

    $html = New-PaginaHtml 'Relatório de Entrega' "Formatação e configuração do computador $env:COMPUTERNAME" (ConvertTo-HtmlSecoes $secoes)
    $arq  = Join-Path $PastaRelatorios ('Entrega_{0}_{1}.html' -f $env:COMPUTERNAME, (Get-Date -Format 'yyyy-MM-dd_HH-mm'))
    Save-Html $arq $html
    Write-Ok "Relatório salvo no pendrive: $arq"

    if ($CopiarEntregaParaAreaTrabalho) {
        $destino = Join-Path "$env:PUBLIC\Desktop" 'Relatório de Entrega.html'
        Copy-Item $arq $destino -Force -ErrorAction SilentlyContinue
        if (Test-Path $destino) { Write-Ok 'Cópia deixada na Área de Trabalho.' }
    }
    if (-not $Silencioso) { Start-Process $arq }
    return $arq
}

# ---------------------------------------------------------------------
#  10b. Drivers
# ---------------------------------------------------------------------
# Fabricante e modelo (o mesmo "Modelo do sistema" das Informações do Sistema)
function Get-Fabricante {
    $cs  = Get-CimInstance Win32_ComputerSystem -ErrorAction SilentlyContinue
    $fab = ("$($cs.Manufacturer)").Trim()
    $mod = ("$($cs.Model)").Trim()
    if ($fab -match 'LENOVO') {
        # Na Lenovo o nome comercial (ex.: ThinkPad T14) fica em outro campo
        $nome = (Get-CimInstance Win32_ComputerSystemProduct -ErrorAction SilentlyContinue).Version
        if ($nome) { $mod = "$nome ($mod)" }
    }
    $marca = 'Outra'
    if     ($fab -match 'Dell|Alienware')          { $marca = 'Dell' }
    elseif ($fab -match 'Lenovo')                  { $marca = 'Lenovo' }
    elseif ($fab -match '^HP\b|Hewlett')           { $marca = 'HP' }
    elseif ($fab -match 'To be filled|System manufacturer|Default string|^$') {
        # PC montado: mostra a placa-mãe
        $mb = Get-CimInstance Win32_BaseBoard -ErrorAction SilentlyContinue
        $fab = 'PC montado'
        $mod = ("Placa-mãe $($mb.Manufacturer) $($mb.Product)").Trim()
    }
    return @{ Marca = $marca; Fabricante = $fab; Modelo = $mod }
}

# Dispositivos com ponto de exclamação no Gerenciador de Dispositivos
function Get-DispositivosSemDriver {
    $problemas = @{ 1 = 'Mal configurado'; 10 = 'Não consegue iniciar'; 18 = 'Precisa reinstalar o driver'; 28 = 'Sem driver'; 31 = 'Driver com erro'; 39 = 'Driver corrompido'; 43 = 'Parou de responder' }
    return @(Get-CimInstance Win32_PnPEntity -Filter 'ConfigManagerErrorCode <> 0' -ErrorAction SilentlyContinue |
        Where-Object { $_.ConfigManagerErrorCode -notin 22, 45 } |   # 22 = desativado de propósito, 45 = desconectado
        ForEach-Object {
            $nome = $_.Name; if (-not $nome) { $nome = 'Dispositivo desconhecido' }
            $prob = $problemas[[int]$_.ConfigManagerErrorCode]; if (-not $prob) { $prob = "Código $($_.ConfigManagerErrorCode)" }
            [pscustomobject]@{ 'Dispositivo' = $nome; 'Problema' = $prob }
        })
}

function Show-DispositivosSemDriver {
    Write-Etapa 'Verificação de drivers'
    $f = Get-Fabricante
    Write-Info "Computador: $($f.Fabricante) $($f.Modelo)"
    switch ($f.Marca) {
        'Dell'   { Write-Info 'Ferramenta do fabricante disponível: Dell Command Update' }
        'Lenovo' { Write-Info 'Ferramenta do fabricante disponível: catálogo oficial da Lenovo (LSUClient)' }
        'HP'     { Write-Info 'Ferramenta do fabricante disponível: HP Image Assistant' }
        default  { Write-Info 'Sem ferramenta automática do fabricante: os drivers virão do Windows Update.' }
    }
    $sem = @(Get-DispositivosSemDriver)
    if ($sem.Count -eq 0) {
        Write-Ok 'Nenhum dispositivo sem driver ou com erro.'
    } else {
        Write-Aviso "$($sem.Count) dispositivo(s) com problema:"
        foreach ($d in $sem) { Write-Host "         - $($d.Dispositivo)  ($($d.Problema))" -ForegroundColor Yellow }
    }
    return $sem
}

# Instala um pacote do winget de forma silenciosa (usado pelas ferramentas de driver)
function Install-PacoteSilencioso([string]$Id, [string]$Choco) {
    $w = Get-WingetPath
    if (-not $w) { $w = Initialize-Winget }
    if ($w) {
        & $w install --id $Id --exact --silent --accept-package-agreements --accept-source-agreements --disable-interactivity --source winget 2>&1 | Out-Null
        if ($LASTEXITCODE -eq 0 -or $WG_JA_INSTALADO -contains $LASTEXITCODE) { return $true }
    }
    if ($Choco -and $UsarChocolatey) {
        $c = Initialize-Choco
        if ($c) {
            & $c install $Choco -y --no-progress --limit-output 2>&1 | Out-Null
            return ($LASTEXITCODE -in 0, 1641, 3010)
        }
    }
    return $false
}

# --- Dell: Dell Command Update ---
function Install-DriversDell {
    Write-Info 'Preparando o Dell Command Update...'
    $caminhos = "$env:ProgramFiles\Dell\CommandUpdate\dcu-cli.exe", "${env:ProgramFiles(x86)}\Dell\CommandUpdate\dcu-cli.exe"
    $dcu = $caminhos | Where-Object { Test-Path $_ } | Select-Object -First 1
    if (-not $dcu) {
        [void](Install-PacoteSilencioso 'Microsoft.DotNet.DesktopRuntime.8' 'dotnet-8.0-desktopruntime')   # exigido pelo DCU 5.x
        if (-not (Install-PacoteSilencioso 'Dell.CommandUpdate.Universal' 'dellcommandupdate-uwp')) { [void](Install-PacoteSilencioso 'Dell.CommandUpdate' 'dellcommandupdate') }
        $dcu = $caminhos | Where-Object { Test-Path $_ } | Select-Object -First 1
    }
    if (-not $dcu) { Write-Aviso 'Não foi possível instalar o Dell Command Update neste modelo.'; return 'Dell Command Update indisponível' }

    Write-Info 'Procurando e instalando drivers da Dell (pode levar vários minutos)...'
    $log = Join-Path $env:TEMP 'SetupPC_dcu.log'
    $p = Start-Process $dcu -ArgumentList "/applyUpdates -updateType=driver -silent -reboot=disable -outputLog=`"$log`"" -Wait -PassThru -WindowStyle Hidden
    switch ($p.ExitCode) {
        0       { Write-Ok 'Drivers da Dell instalados.'; return 'Drivers instalados' }
        1       { Write-Ok 'Drivers da Dell instalados (requer reinício).'; $script:PrecisaReiniciar = $true; return 'Drivers instalados (reiniciar)' }
        5       { Write-Ok 'Drivers da Dell instalados (requer reinício).'; $script:PrecisaReiniciar = $true; return 'Drivers instalados (reiniciar)' }
        500     { Write-Ok 'Dell: todos os drivers já estão atualizados.'; return 'Já atualizado' }
        default { Write-Aviso "Dell Command Update terminou com código $($p.ExitCode)."; return "Atenção (código $($p.ExitCode))" }
    }
}

# --- Lenovo: catálogo oficial via módulo LSUClient ---
function Install-DriversLenovo {
    Write-Info 'Preparando o catálogo de drivers da Lenovo (LSUClient)...'
    try {
        if (-not (Get-Module -ListAvailable -Name LSUClient)) {
            Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force -Scope AllUsers -ErrorAction Stop | Out-Null
            Install-Module -Name LSUClient -Force -Scope AllUsers -ErrorAction Stop
        }
        Import-Module LSUClient -ErrorAction Stop
        # Só drivers que instalam sozinhos; BIOS/firmware ficam de fora por segurança
        $ups = @(Get-LSUpdate -ErrorAction Stop | Where-Object {
            $_.Installer.Unattended -and "$($_.Type)" -notmatch 'BIOS|Firmware' -and "$($_.Category)" -notmatch 'BIOS|UEFI|Firmware' -and "$($_.Title)" -notmatch 'BIOS|UEFI'
        })
        if ($ups.Count -eq 0) { Write-Ok 'Lenovo: todos os drivers já estão atualizados.'; return 'Já atualizado' }
        Write-Info "Baixando $($ups.Count) atualização(ões) da Lenovo..."
        $ups | Save-LSUpdate -ErrorAction SilentlyContinue | Out-Null
        $i = 0
        foreach ($u in $ups) {
            $i++
            Write-Info "  [$i/$($ups.Count)] $($u.Title)"
            Install-LSUpdate -Package $u -ErrorAction SilentlyContinue | Out-Null
        }
        Write-Ok "Lenovo: $($ups.Count) driver(s) instalado(s)."
        return "$($ups.Count) driver(s) instalado(s)"
    } catch {
        Write-Aviso "Catálogo da Lenovo indisponível: $($_.Exception.Message)"
        return 'Catálogo Lenovo indisponível'
    }
}

# --- HP: HP Image Assistant ---
function Install-DriversHP {
    Write-Info 'Preparando o HP Image Assistant...'
    $locais = 'C:\SWSetup', "$env:ProgramFiles\HP", "${env:ProgramFiles(x86)}\HP", "$env:LOCALAPPDATA\Microsoft\WinGet\Packages", 'C:\HPIA'
    $acha = { Get-ChildItem -Path $locais -Filter 'HPImageAssistant.exe' -Recurse -Depth 4 -ErrorAction SilentlyContinue | Select-Object -First 1 }
    $hpia = & $acha
    if (-not $hpia) { [void](Install-PacoteSilencioso 'HP.ImageAssistant' 'hpimageassistant'); $hpia = & $acha }
    if (-not $hpia) { Write-Aviso 'Não foi possível preparar o HP Image Assistant.'; return 'HP Image Assistant indisponível' }

    Write-Info 'Procurando e instalando drivers da HP (pode levar vários minutos)...'
    $base = 'C:\SetupPC_HPIA'   # caminho sem espaços (o HPIA tem problemas com espaços)
    New-Item -ItemType Directory -Path "$base\Relatorio", "$base\Softpaqs" -Force | Out-Null
    $p = Start-Process $hpia.FullName -ArgumentList "/Operation:Analyze /Category:Drivers /Selection:All /Action:Install /Silent /AutoCleanup /ReportFolder:$base\Relatorio /SoftpaqDownloadFolder:$base\Softpaqs" -Wait -PassThru
    Remove-Item "$base\Softpaqs" -Recurse -Force -ErrorAction SilentlyContinue
    switch ($p.ExitCode) {
        0       { Write-Ok 'Drivers da HP instalados.'; return 'Drivers instalados' }
        3010    { Write-Ok 'Drivers da HP instalados (requer reinício).'; $script:PrecisaReiniciar = $true; return 'Drivers instalados (reiniciar)' }
        256     { Write-Ok 'HP: nenhum driver novo recomendado.'; return 'Já atualizado' }
        257     { Write-Ok 'HP: nenhum driver novo recomendado.'; return 'Já atualizado' }
        default { Write-Aviso "HP Image Assistant terminou com código $($p.ExitCode)."; return "Atenção (código $($p.ExitCode))" }
    }
}

# --- Qualquer marca: drivers pelo Windows Update ---
function Install-DriversWindowsUpdate {
    Write-Info 'Procurando drivers no Windows Update...'
    try {
        $sessao = New-Object -ComObject Microsoft.Update.Session
        $sessao.ClientApplicationID = 'SetupPC'
        $busca  = $sessao.CreateUpdateSearcher()
        $res    = $busca.Search("IsInstalled=0 and Type='Driver' and IsHidden=0")
        if ($res.Updates.Count -eq 0) { Write-Ok 'Windows Update: nenhum driver pendente.'; return 'Nenhum driver pendente' }

        $lista = New-Object -ComObject Microsoft.Update.UpdateColl
        foreach ($u in $res.Updates) {
            if (-not $u.EulaAccepted) { $u.AcceptEula() }
            Write-Info "  - $($u.Title)"
            [void]$lista.Add($u)
        }
        Write-Info "Baixando $($lista.Count) driver(s)..."
        $down = $sessao.CreateUpdateDownloader(); $down.Updates = $lista; [void]$down.Download()
        Write-Info 'Instalando...'
        $inst = $sessao.CreateUpdateInstaller(); $inst.Updates = $lista
        $r = $inst.Install()
        $ok = 0
        for ($i = 0; $i -lt $lista.Count; $i++) { if ($r.GetUpdateResult($i).ResultCode -in 2, 3) { $ok++ } }
        if ($r.RebootRequired) { $script:PrecisaReiniciar = $true }
        Write-Ok "Windows Update: $ok de $($lista.Count) driver(s) instalado(s)."
        return "$ok de $($lista.Count) driver(s) instalado(s)"
    } catch {
        Write-Aviso "Não foi possível usar o Windows Update: $($_.Exception.Message)"
        return 'Windows Update indisponível'
    }
}

# Fluxo completo: fabricante + Windows Update + verificação final
function Invoke-Drivers {
    Write-Etapa 'Drivers'
    $f = Get-Fabricante
    Write-Info "Computador: $($f.Fabricante) $($f.Modelo)"
    $antes = @(Get-DispositivosSemDriver)
    Write-Info "Dispositivos sem driver ou com erro antes: $($antes.Count)"

    $script:ResultadoDrivers = New-Object System.Collections.Generic.List[object]
    switch ($f.Marca) {
        'Dell'   { $script:ResultadoDrivers.Add([pscustomobject]@{ 'Origem' = 'Dell Command Update'; 'Status' = (Install-DriversDell) }) }
        'Lenovo' { $script:ResultadoDrivers.Add([pscustomobject]@{ 'Origem' = 'Catálogo Lenovo';     'Status' = (Install-DriversLenovo) }) }
        'HP'     { $script:ResultadoDrivers.Add([pscustomobject]@{ 'Origem' = 'HP Image Assistant';  'Status' = (Install-DriversHP) }) }
    }
    $script:ResultadoDrivers.Add([pscustomobject]@{ 'Origem' = 'Windows Update'; 'Status' = (Install-DriversWindowsUpdate) })

    $depois = @(Get-DispositivosSemDriver)
    $script:DispositivosPendentes = $depois
    if ($depois.Count -eq 0) {
        Write-Ok 'Todos os dispositivos estão com driver.'
        Add-Acao 'Drivers instalados e atualizados (todos os dispositivos com driver)'
    } else {
        Write-Aviso "$($depois.Count) dispositivo(s) ainda sem driver ou com erro. Tente o Driver Booster em Ferramentas."
        foreach ($d in $depois) { Write-Host "         - $($d.Dispositivo)  ($($d.Problema))" -ForegroundColor Yellow }
        Add-Acao "Drivers instalados e atualizados ($($depois.Count) dispositivo(s) ainda com pendência)"
    }
}

# --- Driver Booster (manual: você revisa e clica em Atualizar) ---
function Get-DriverBoosterExe {
    return Get-ChildItem -Path "${env:ProgramFiles(x86)}\IObit\Driver Booster", "$env:ProgramFiles\IObit\Driver Booster" -Filter 'DriverBooster.exe' -Recurse -ErrorAction SilentlyContinue |
           Sort-Object FullName -Descending | Select-Object -First 1
}

function Remove-DriverBooster {
    Write-Info 'Desinstalando o Driver Booster...'
    Get-Process -Name 'DriverBooster', 'DBDownloader', 'Scheduler' -ErrorAction SilentlyContinue |
        Where-Object { $_.Path -like '*IObit*' } | Stop-Process -Force -ErrorAction SilentlyContinue
    $w = Get-WingetPath
    if ($w) { & $w uninstall --id IObit.DriverBooster --exact --silent --accept-source-agreements --disable-interactivity 2>&1 | Out-Null }
    $unins = Get-ChildItem -Path "${env:ProgramFiles(x86)}\IObit\Driver Booster", "$env:ProgramFiles\IObit\Driver Booster" -Filter 'unins000.exe' -Recurse -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($unins) { Start-Process $unins.FullName -ArgumentList '/VERYSILENT /NORESTART' -Wait -ErrorAction SilentlyContinue }
    if (Get-DriverBoosterExe) { Write-Aviso 'O Driver Booster não foi removido totalmente; desinstale pelo Painel de Controle.' }
    else { Write-Ok 'Driver Booster desinstalado.' }
}

# Programas instalados (lidos do registro), para detectar "brindes" instalados junto
function Get-ProgramasInstalados {
    $chaves = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*',
              'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*',
              'HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*'
    return @(Get-ItemProperty -Path $chaves -ErrorAction SilentlyContinue |
        Where-Object { $_.DisplayName } |
        Select-Object PSChildName, DisplayName, Publisher, UninstallString, QuietUninstallString)
}

# Remove qualquer programa da IObit/iTop que apareceu depois de $Antes (menos o próprio Driver Booster)
function Remove-ExtrasIObit([object[]]$Antes) {
    $idsAntes = @($Antes | ForEach-Object { $_.PSChildName })
    $extras = @(Get-ProgramasInstalados | Where-Object {
        $idsAntes -notcontains $_.PSChildName -and
        ("$($_.Publisher) $($_.DisplayName)" -match 'IObit|iTop|Advanced SystemCare|Smart Defrag|Malware Fighter|Software Updater') -and
        $_.DisplayName -notmatch 'Driver Booster'
    })
    if ($extras.Count -eq 0) { Write-Ok 'Nenhum programa extra da IObit/iTop foi instalado junto.'; return }

    foreach ($x in $extras) {
        Write-Aviso "Programa extra detectado: $($x.DisplayName). Removendo..."
        $cmd = $x.QuietUninstallString
        if (-not $cmd) { $cmd = $x.UninstallString }
        if ($cmd -match '^\s*"([^"]+)"\s*(.*)$') { $exe = $Matches[1]; $args0 = $Matches[2] }
        elseif ($cmd -match '^\s*(\S+)\s*(.*)$') { $exe = $Matches[1]; $args0 = $Matches[2] }
        else { continue }
        if (-not $x.QuietUninstallString) { $args0 = "$args0 /VERYSILENT /SUPPRESSMSGBOXES /NORESTART".Trim() }
        Get-Process -ErrorAction SilentlyContinue | Where-Object { $_.Path -and $_.Path -like "$(Split-Path $exe)*" } | Stop-Process -Force -ErrorAction SilentlyContinue
        try { Start-Process -FilePath $exe -ArgumentList $args0 -Wait -ErrorAction Stop } catch { Write-Aviso "Não foi possível remover $($x.DisplayName); desinstale pelo Painel de Controle." ; continue }
        Write-Ok "Removido: $($x.DisplayName)"
        Add-Acao "Removido programa extra instalado pela IObit: $($x.DisplayName)"
    }
}

function Invoke-DriverBooster([bool]$Desinstalar) {
    Write-Etapa 'Driver Booster'
    $antes = Get-ProgramasInstalados   # "foto" dos programas antes de mexer
    $exe = Get-DriverBoosterExe
    if (-not $exe) {
        Write-Info 'Instalando o Driver Booster (silencioso)...'
        [void](Install-PacoteSilencioso 'IObit.DriverBooster' 'driverbooster')
        $exe = Get-DriverBoosterExe
    }
    if (-not $exe) { Write-Erro 'Não foi possível instalar o Driver Booster.'; return }

    # 1ª checagem: brindes instalados junto com o instalador
    Remove-ExtrasIObit $antes

    Write-Info 'Abrindo o Driver Booster. Ele faz a busca sozinho; confira a lista e clique em "Atualizar agora".'
    Start-Process $exe.FullName
    Show-AvisoUsuario "O Driver Booster foi aberto.`n`n1. Espere a busca terminar.`n2. Confira a lista e clique em ""Atualizar agora"".`n3. Se ele oferecer outros programas (Advanced SystemCare, iTop etc.), recuse.`n4. Quando os drivers terminarem de instalar, volte aqui e clique em OK."
    Add-Acao 'Drivers revisados com o Driver Booster'

    if ($Desinstalar) { Remove-DriverBooster }
    # 2ª checagem: algo oferecido dentro do programa durante o uso
    Remove-ExtrasIObit $antes

    $sem = @(Get-DispositivosSemDriver)
    if ($sem.Count -eq 0) { Write-Ok 'Todos os dispositivos estão com driver.' }
    else { Write-Aviso "$($sem.Count) dispositivo(s) ainda com pendência." }
}

function Invoke-MenuDrivers {
    do {
        Write-Etapa 'Drivers'
        Write-Host '   [1]  Verificar (lista dispositivos sem driver)'
        Write-Host '   [2]  Instalar drivers (fabricante + Windows Update)'
        Write-Host '   [3]  Abrir o Driver Booster (e desinstalar ao terminar)'
        Write-Host '   [4]  Abrir o Driver Booster (e manter instalado)'
        Write-Host '   [0]  Voltar'
        Write-Host ''
        $op = (Read-Host '   Escolha uma opção').Trim()
        switch ($op) {
            '1' { [void](Show-DispositivosSemDriver) }
            '2' { Invoke-Drivers }
            '3' { Invoke-DriverBooster $true }
            '4' { Invoke-DriverBooster $false }
        }
        if ($op -ne '0') { Write-Host ''; Read-Host '   Enter para continuar' | Out-Null }
    } until ($op -eq '0')
}

# ---------------------------------------------------------------------
#  11. Instalação (usada pelo modo texto e pela interface gráfica)
# ---------------------------------------------------------------------
function Invoke-Instalacao {
    $inicio = Get-Date
    $Resultados.Clear()
    $script:Acoes.Clear()
    $script:ArquivoEntrega = $null
    $script:ResultadoDrivers = $null
    $script:DispositivosPendentes = $null

    # Conta as etapas para a barra de progresso
    $total = 1   # configurações do sistema
    if (-not $SemPontoRestauracao) { $total++ }
    if (-not $SemProgramas) {
        $total += 1 + @(Get-ListaProgramas).Count + @($InstaladoresLocais).Count
        if ($script:OfficeEscolhido) { $total++ }
    }
    if ($InstalarDrivers)        { $total++ }
    if (-not $SemDebloat)        { $total++ }
    if ($Otimizar)               { $total++ }
    if ($DesativarWindowsUpdate) { $total++ }
    if ($GerarRelatorioEntrega)  { $total++ }
    $script:PassoAtual  = 0
    $script:TotalPassos = $total

    if (-not $SemPontoRestauracao) { Step-Progresso 'Criando ponto de restauração'; New-PontoRestauracao }
    if (-not $SemProgramas)        { Install-Programas }
    if ($InstalarDrivers)          { Step-Progresso 'Instalando drivers (fabricante + Windows Update)'; Invoke-Drivers }
    Step-Progresso 'Configurações do sistema'
    Invoke-ConfiguracoesSistema
    if (-not $SemDebloat)          { Step-Progresso 'Debloat: removendo apps e desativando telemetria'; Invoke-Debloat }
    if ($Otimizar)                 { Step-Progresso 'Aplicando a otimização avançada'; Invoke-Otimizacoes }
    if ($DesativarWindowsUpdate)   { Step-Progresso 'Desativando o Windows Update'; Disable-WindowsUpdate }

    Write-Etapa 'Resumo'
    foreach ($r in $Resultados) { Write-Host ('  {0,-45} {1}' -f $r.Programa, $r.Status) }
    $falhas = @($Resultados | Where-Object { $_.Status -like 'FALHOU*' -or $_.Status -like '*não encontrado' })
    if ($falhas.Count -gt 0) { Write-Aviso "$($falhas.Count) item(ns) com problema. Veja o log." }
    $duracao = (Get-Date) - $inicio
    Write-Ok ('Concluído em {0} min.' -f [math]::Max(1, [math]::Round($duracao.TotalMinutes)))
    Write-Info "Log salvo em: $ArquivoLog"

    if ($GerarRelatorioEntrega) {
        Step-Progresso 'Gerando o relatório de entrega'
        try   { $script:ArquivoEntrega = New-RelatorioEntrega $duracao }
        catch { Write-Erro "Falha ao gerar o relatório de entrega: $($_.Exception.Message)" }
    }
}

# ---------------------------------------------------------------------
#  12. Interface gráfica (WPF)
# ---------------------------------------------------------------------
$XamlPrincipal = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="SetupPC - Danilo Américo" Width="1040" Height="760" MinWidth="900" MinHeight="620"
        WindowStartupLocation="CenterScreen" Background="#F3F5F8" FontFamily="Segoe UI" FontSize="13">
  <Window.Resources>
    <Style x:Key="Card" TargetType="Border">
      <Setter Property="Background" Value="White"/>
      <Setter Property="BorderBrush" Value="#E1E5EC"/>
      <Setter Property="BorderThickness" Value="1"/>
      <Setter Property="CornerRadius" Value="10"/>
      <Setter Property="Padding" Value="16,12"/>
      <Setter Property="Margin" Value="0,0,0,12"/>
    </Style>
    <Style x:Key="Titulo" TargetType="TextBlock">
      <Setter Property="FontSize" Value="11.5"/>
      <Setter Property="FontWeight" Value="SemiBold"/>
      <Setter Property="Foreground" Value="#14306B"/>
      <Setter Property="Margin" Value="0,0,0,8"/>
    </Style>
    <Style x:Key="Dica" TargetType="TextBlock">
      <Setter Property="FontSize" Value="12"/>
      <Setter Property="Foreground" Value="#5A6478"/>
      <Setter Property="TextWrapping" Value="Wrap"/>
      <Setter Property="Margin" Value="0,6,0,0"/>
    </Style>
    <Style TargetType="CheckBox"><Setter Property="Margin" Value="0,3,12,3"/><Setter Property="Cursor" Value="Hand"/></Style>
    <Style TargetType="RadioButton"><Setter Property="Margin" Value="0,3"/><Setter Property="Cursor" Value="Hand"/></Style>
    <Style x:Key="Botao" TargetType="Button">
      <Setter Property="Background" Value="White"/>
      <Setter Property="Foreground" Value="#1D2433"/>
      <Setter Property="BorderBrush" Value="#C9D1DE"/>
      <Setter Property="Padding" Value="12,8"/>
      <Setter Property="Margin" Value="0,0,0,8"/>
      <Setter Property="Cursor" Value="Hand"/>
      <Setter Property="HorizontalContentAlignment" Value="Left"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="Button">
            <Border x:Name="b" Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}"
                    BorderThickness="1" CornerRadius="6" Padding="{TemplateBinding Padding}">
              <ContentPresenter HorizontalAlignment="{TemplateBinding HorizontalContentAlignment}" VerticalAlignment="Center"/>
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsMouseOver" Value="True"><Setter TargetName="b" Property="Background" Value="#EEF3FF"/><Setter TargetName="b" Property="BorderBrush" Value="#2563EB"/></Trigger>
              <Trigger Property="IsEnabled" Value="False"><Setter TargetName="b" Property="Opacity" Value="0.45"/></Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>
    <Style x:Key="BotaoPrincipal" TargetType="Button">
      <Setter Property="Foreground" Value="White"/>
      <Setter Property="FontSize" Value="15"/>
      <Setter Property="FontWeight" Value="SemiBold"/>
      <Setter Property="Cursor" Value="Hand"/>
      <Setter Property="Margin" Value="0,0,8,8"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="Button">
            <Border x:Name="b" Background="#2563EB" CornerRadius="8" Padding="22,12">
              <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsMouseOver" Value="True"><Setter TargetName="b" Property="Background" Value="#1D4ED8"/></Trigger>
              <Trigger Property="IsEnabled" Value="False"><Setter TargetName="b" Property="Background" Value="#94A3B8"/></Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>
  </Window.Resources>

  <DockPanel>
    <!-- Cabeçalho -->
    <Border DockPanel.Dock="Top" Background="#14306B" Padding="26,16">
      <Grid>
        <StackPanel>
          <TextBlock Text="SetupPC" Foreground="White" FontSize="24" FontWeight="SemiBold"/>
          <TextBlock Text="Pós-formatação automática  ·  Script feito por Danilo Américo" Foreground="#C7D3EE" Margin="0,2,0,0"/>
        </StackPanel>
        <TextBlock x:Name="txtPC" HorizontalAlignment="Right" VerticalAlignment="Center" Foreground="#C7D3EE" TextAlignment="Right" LineHeight="20"/>
      </Grid>
    </Border>

    <!-- Barra de progresso -->
    <Border DockPanel.Dock="Bottom" Background="White" BorderBrush="#E1E5EC" BorderThickness="0,1,0,0" Padding="26,12">
      <Grid>
        <Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="Auto"/></Grid.ColumnDefinitions>
        <StackPanel>
          <TextBlock x:Name="txtEtapa" Text="Pronto. Escolha as opções e clique em Iniciar instalação." Foreground="#1D2433"/>
          <ProgressBar x:Name="pbGeral" Height="10" Margin="0,7,0,0" Minimum="0" Maximum="100" Value="0" Foreground="#2563EB" Background="#E8EDF5" BorderThickness="0"/>
          <ProgressBar x:Name="pbAtividade" Height="3" Margin="0,4,0,0" Foreground="#93B4F5" Background="Transparent" BorderThickness="0" Visibility="Hidden"/>
        </StackPanel>
        <TextBlock x:Name="txtPct" Grid.Column="1" Text="" FontSize="20" FontWeight="SemiBold" Foreground="#14306B" VerticalAlignment="Center" Margin="18,0,0,0" MinWidth="60" TextAlignment="Right"/>
      </Grid>
    </Border>

    <TabControl x:Name="tabs" Margin="16,12,16,8" Background="Transparent" BorderThickness="0" Padding="0,8,0,0">

      <!-- ===================== INSTALAÇÃO ===================== -->
      <TabItem x:Name="tabInstalar" Header="   Instalação   " FontSize="14">
        <ScrollViewer VerticalScrollBarVisibility="Auto" FontSize="13">
          <Grid Margin="0,4,6,0">
            <Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="14"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>

            <StackPanel Grid.Column="0">
              <Border Style="{StaticResource Card}">
                <StackPanel>
                  <TextBlock Style="{StaticResource Titulo}" Text="TIPO DE INSTALAÇÃO"/>
                  <RadioButton x:Name="rbPadrao" GroupName="modo" IsChecked="True" Content="Padrão"/>
                  <RadioButton x:Name="rbGamer" GroupName="modo" Content="Gamer  (marca os jogos e a otimização avançada)"/>
                </StackPanel>
              </Border>
              <Border Style="{StaticResource Card}">
                <StackPanel>
                  <TextBlock Style="{StaticResource Titulo}" Text="MICROSOFT OFFICE"/>
                  <ComboBox x:Name="cbOffice" Padding="8,6"/>
                  <TextBlock Style="{StaticResource Dica}" Text="Em português do Brasil, baixado direto da Microsoft. Precisa de licença (chave ou conta Microsoft) para ativar."/>
                </StackPanel>
              </Border>
              <Border Style="{StaticResource Card}">
                <StackPanel>
                  <TextBlock Style="{StaticResource Titulo}" Text="PROGRAMAS"/>
                  <WrapPanel x:Name="spProgramas"/>
                  <TextBlock Style="{StaticResource Titulo}" Text="JOGOS" Margin="0,12,0,8"/>
                  <WrapPanel x:Name="spGamer"/>
                  <TextBlock Style="{StaticResource Titulo}" Text="OPCIONAIS" Margin="0,12,0,8"/>
                  <WrapPanel x:Name="spOpcionais"/>
                </StackPanel>
              </Border>
            </StackPanel>

            <StackPanel Grid.Column="2">
              <Border Style="{StaticResource Card}">
                <StackPanel>
                  <TextBlock Style="{StaticResource Titulo}" Text="DEBLOAT"/>
                  <CheckBox x:Name="chkDebloat" IsChecked="True" Content="Remover apps desnecessários, telemetria e anúncios"/>
                  <CheckBox x:Name="chkXbox" Margin="22,3,0,3" Content="Remover também os apps do Xbox"/>
                  <CheckBox x:Name="chkOneDrive" Margin="22,3,0,3" Content="Remover o OneDrive"/>
                </StackPanel>
              </Border>
              <Border Style="{StaticResource Card}">
                <StackPanel>
                  <TextBlock Style="{StaticResource Titulo}" Text="CONFIGURAÇÕES"/>
                  <CheckBox x:Name="chkRestauracao" IsChecked="True" Content="Criar ponto de restauração antes de começar"/>
                  <CheckBox x:Name="chkIcones" Content="Ícones na Área de Trabalho (Este Computador, Lixeira...)"/>
                  <CheckBox x:Name="chkRede" Content="Rede privada, descoberta de rede e impressoras"/>
                  <CheckBox x:Name="chkHibernacao" Content="Desativar a hibernação (desktops)"/>
                </StackPanel>
              </Border>
              <Border Style="{StaticResource Card}">
                <StackPanel>
                  <TextBlock Style="{StaticResource Titulo}" Text="DESEMPENHO"/>
                  <CheckBox x:Name="chkOtimizar" Content="Otimização avançada (plano de energia, input lag, jogos)"/>
                  <CheckBox x:Name="chkPlanoNotebook" Margin="22,3,0,3" Content="Usar Desempenho Máximo mesmo sendo notebook"/>
                  <CheckBox x:Name="chkWU" Content="Desativar o Windows Update (a Loja continua funcionando)"/>
                </StackPanel>
              </Border>
              <Border Style="{StaticResource Card}">
                <StackPanel>
                  <TextBlock Style="{StaticResource Titulo}" Text="DRIVERS E RELATÓRIO"/>
                  <CheckBox x:Name="chkDrivers" Content="Instalar drivers (fabricante + Windows Update)"/>
                  <CheckBox x:Name="chkRelatorio" Content="Gerar relatório de entrega para o cliente"/>
                </StackPanel>
              </Border>
              <Button x:Name="btnIniciar" Style="{StaticResource BotaoPrincipal}" Content="Iniciar instalação" Margin="0,2,0,12"/>
            </StackPanel>
          </Grid>
        </ScrollViewer>
      </TabItem>

      <!-- ===================== FERRAMENTAS ===================== -->
      <TabItem x:Name="tabFerramentas" Header="   Ferramentas   " FontSize="14">
        <ScrollViewer VerticalScrollBarVisibility="Auto" FontSize="13">
          <Grid Margin="0,4,6,0">
            <Grid.ColumnDefinitions>
              <ColumnDefinition Width="*"/><ColumnDefinition Width="14"/><ColumnDefinition Width="*"/>
            </Grid.ColumnDefinitions>
            <StackPanel Grid.Column="0">
              <Border Style="{StaticResource Card}">
                <StackPanel>
                  <TextBlock Style="{StaticResource Titulo}" Text="DRIVERS"/>
                  <TextBlock x:Name="txtModelo" Style="{StaticResource Dica}" Margin="0,0,0,10"/>
                  <Button x:Name="btnDriversVerificar" Style="{StaticResource Botao}" Content="Verificar (lista dispositivos sem driver)"/>
                  <Button x:Name="btnDriversInstalar" Style="{StaticResource Botao}" Content="Instalar drivers (fabricante + Windows Update)"/>
                  <Button x:Name="btnDriverBooster" Style="{StaticResource Botao}" Content="Abrir o Driver Booster (revisão manual)"/>
                  <CheckBox x:Name="chkDBDesinstalar" IsChecked="True" Content="Desinstalar o Driver Booster ao terminar"/>
                </StackPanel>
              </Border>
              <Border Style="{StaticResource Card}">
                <StackPanel>
                  <TextBlock Style="{StaticResource Titulo}" Text="DIAGNÓSTICO E LIMPEZA"/>
                  <Button x:Name="btnRelatorioPC" Style="{StaticResource Botao}" Content="Gerar relatório do PC (hardware, disco, bateria)"/>
                  <Button x:Name="btnLimpeza" Style="{StaticResource Botao}" Content="Limpeza de disco (temporários, cache, arquivos antigos)"/>
                  <Button x:Name="btnPasta" Style="{StaticResource Botao}" Content="Abrir pasta de relatórios e logs"/>
                </StackPanel>
              </Border>
              <Border Style="{StaticResource Card}">
                <StackPanel>
                  <TextBlock Style="{StaticResource Titulo}" Text="OTIMIZAÇÃO AVANÇADA"/>
                  <Button x:Name="btnOtimizar" Style="{StaticResource Botao}" Content="Aplicar otimizações de desempenho e input lag"/>
                  <Button x:Name="btnVBS" Style="{StaticResource Botao}" Content="Desativar Integridade de Memória / VBS  (+FPS, -segurança)"/>
                  <Button x:Name="btnReverter" Style="{StaticResource Botao}" Content="Reverter otimizações (volta como estava)"/>
                </StackPanel>
              </Border>
            </StackPanel>
            <StackPanel Grid.Column="2">
              <Border Style="{StaticResource Card}">
                <StackPanel>
                  <TextBlock Style="{StaticResource Titulo}" Text="WINDOWS UPDATE"/>
                  <Button x:Name="btnWUDesativar" Style="{StaticResource Botao}" Content="Desativar (mantém a Microsoft Store funcionando)"/>
                  <Button x:Name="btnWUReativar" Style="{StaticResource Botao}" Content="Reativar"/>
                  <Button x:Name="btnWUStatus" Style="{StaticResource Botao}" Content="Ver status"/>
                </StackPanel>
              </Border>
              <Border Style="{StaticResource Card}">
                <StackPanel>
                  <TextBlock Style="{StaticResource Titulo}" Text="ATIVAÇÃO (COM LICENÇA)"/>
                  <Button x:Name="btnAtivStatus" Style="{StaticResource Botao}" Content="Ver status do Windows e do Office"/>
                  <Button x:Name="btnAtivOEM" Style="{StaticResource Botao}" Content="Ativar Windows com a chave gravada na BIOS"/>
                  <Button x:Name="btnAtivWin" Style="{StaticResource Botao}" Content="Ativar Windows digitando uma chave..."/>
                  <Button x:Name="btnAtivOffice" Style="{StaticResource Botao}" Content="Ativar Office digitando uma chave..."/>
                  <Button x:Name="btnAtivTela" Style="{StaticResource Botao}" Content="Abrir a tela de Ativação do Windows"/>
                </StackPanel>
              </Border>
            </StackPanel>
          </Grid>
        </ScrollViewer>
      </TabItem>

      <!-- ===================== PROGRESSO ===================== -->
      <TabItem x:Name="tabProgresso" Header="   Progresso   " FontSize="14">
        <Grid Margin="0,4,0,0">
          <Grid.RowDefinitions><RowDefinition Height="*"/><RowDefinition Height="Auto"/></Grid.RowDefinitions>
          <Border Background="#0F172A" CornerRadius="10" Padding="4">
            <RichTextBox x:Name="rtbLog" IsReadOnly="True" BorderThickness="0" Background="#0F172A" Foreground="#CBD5E1"
                         FontFamily="Consolas" FontSize="12.5" Padding="10" VerticalScrollBarVisibility="Auto">
              <RichTextBox.Resources>
                <Style TargetType="Paragraph"><Setter Property="Margin" Value="0"/></Style>
              </RichTextBox.Resources>
            </RichTextBox>
          </Border>
          <WrapPanel x:Name="wpFinal" Grid.Row="1" Margin="0,12,0,0" Visibility="Collapsed">
            <Button x:Name="btnAbrirEntrega" Style="{StaticResource BotaoPrincipal}" Content="Abrir relatório de entrega"/>
            <Button x:Name="btnIrAtivacao" Style="{StaticResource Botao}" Content="Ir para Ativação" Margin="0,0,8,8" Padding="16,12" HorizontalContentAlignment="Center"/>
            <Button x:Name="btnReiniciar" Style="{StaticResource Botao}" Content="Reiniciar o computador" Margin="0,0,8,8" Padding="16,12" HorizontalContentAlignment="Center"/>
          </WrapPanel>
        </Grid>
      </TabItem>
    </TabControl>
  </DockPanel>
</Window>
'@

$XamlEntrada = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="SetupPC" Width="480" SizeToContent="Height" WindowStartupLocation="CenterOwner"
        ResizeMode="NoResize" FontFamily="Segoe UI" FontSize="13" Background="White">
  <StackPanel Margin="20">
    <TextBlock x:Name="lbl" TextWrapping="Wrap" Margin="0,0,0,12"/>
    <TextBox x:Name="txt" Padding="8,6" FontFamily="Consolas" FontSize="15" CharacterCasing="Upper"/>
    <StackPanel Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,16,0,0">
      <Button x:Name="ok" Content="OK" Width="96" Padding="0,6" Margin="0,0,8,0" IsDefault="True"/>
      <Button x:Name="cancelar" Content="Cancelar" Width="96" Padding="0,6" IsCancel="True"/>
    </StackPanel>
  </StackPanel>
</Window>
'@

# Código que roda na thread da janela (runspace separado). Ele só conversa com
# o script principal pelo $sync: manda comandos e lê log/progresso/perguntas.
$GuiScript = {
    try {
        Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase
        $janela = [Windows.Markup.XamlReader]::Parse($sync.Xaml)
        $c = @{}
        foreach ($m in [regex]::Matches($sync.Xaml, 'x:Name="(\w+)"')) {
            $n = $m.Groups[1].Value
            $obj = $janela.FindName($n)
            if ($obj) { $c[$n] = $obj }
        }
        $d = $sync.Dados
        $conv = New-Object System.Windows.Media.BrushConverter

        # ---------- helpers ----------
        function Add-LinhaLog($item) {
            $t = [string]$item.Texto
            if ($t -match '^\s*=+\s*$') { return }
            switch ($item.Cor) {
                'Green'    { $cor = '#4ADE80' }
                'Yellow'   { $cor = '#FACC15' }
                'Red'      { $cor = '#F87171' }
                'Cyan'     { $cor = '#7DD3FC' }
                'DarkCyan' { $cor = '#7DD3FC' }
                'White'    { $cor = '#F1F5F9' }
                'DarkGray' { $cor = '#64748B' }
                default    { $cor = '#CBD5E1' }
            }
            $par = New-Object System.Windows.Documents.Paragraph
            $run = New-Object System.Windows.Documents.Run($t)
            $run.Foreground = $conv.ConvertFromString($cor)
            if ($item.Cor -eq 'Cyan') { $run.FontWeight = 'Bold'; $par.Margin = '0,10,0,2' }
            $par.Inlines.Add($run)
            $c.rtbLog.Document.Blocks.Add($par)
        }

        function Show-Entrada([string]$Texto) {
            $w = [Windows.Markup.XamlReader]::Parse($sync.XamlEntrada)
            $w.Owner = $janela
            $w.FindName('lbl').Text = $Texto
            $tb = $w.FindName('txt')
            $w.FindName('ok').Add_Click({ $w.DialogResult = $true })
            $w.Add_ContentRendered({ [void]$tb.Focus() })
            if ($w.ShowDialog()) { return $tb.Text }
            return ''
        }

        function Send-Comando([string]$Acao, $Config) {
            if ($sync.Ocupado) {
                [void][System.Windows.MessageBox]::Show($janela, 'Aguarde a tarefa atual terminar.', 'SetupPC', 'OK', 'Information')
                return
            }
            $sync.Ocupado = $true      # já bloqueia os botões antes do script principal pegar o comando
            $sync.Concluido = $false
            $sync.Comandos.Enqueue(@{ Acao = $Acao; Config = $Config })
            $c.tabs.SelectedItem = $c.tabProgresso
        }

        function New-Check($Painel, $Prog, [bool]$Marcado) {
            $cb = New-Object System.Windows.Controls.CheckBox
            $cb.Content = $Prog.Nome
            $cb.Tag = $Prog.Id
            $cb.IsChecked = $Marcado
            $cb.Width = 205
            [void]$Painel.Children.Add($cb)
            return $cb
        }

        # ---------- preenche a tela ----------
        $c.txtPC.Text = "$($d.PC)`n$($d.Windows)"

        $item = New-Object System.Windows.Controls.ComboBoxItem
        $item.Content = 'Não instalar o Office'; $item.Tag = ''
        [void]$c.cbOffice.Items.Add($item)
        foreach ($o in $d.Office) {
            $item = New-Object System.Windows.Controls.ComboBoxItem
            $item.Content = $o.Nome; $item.Tag = $o.Chave
            [void]$c.cbOffice.Items.Add($item)
            if ($o.Chave -eq $d.OfficePadrao) { $c.cbOffice.SelectedItem = $item }
        }
        if ($c.cbOffice.SelectedIndex -lt 0) { $c.cbOffice.SelectedIndex = 0 }

        $checksGamer = @()
        foreach ($p in $d.Programas)  { [void](New-Check $c.spProgramas $p $true) }
        foreach ($p in $d.Gamer)      { $checksGamer += (New-Check $c.spGamer $p $false) }
        foreach ($p in $d.Opcionais)  { [void](New-Check $c.spOpcionais $p $false) }

        $pd = $d.Padroes
        $c.chkXbox.IsChecked       = $pd.RemoverXbox
        $c.chkOneDrive.IsChecked   = $pd.RemoverOneDrive
        $c.chkIcones.IsChecked     = $pd.Icones
        $c.chkRede.IsChecked       = $pd.Rede
        $c.chkHibernacao.IsChecked = $pd.Hibernacao
        $c.chkRelatorio.IsChecked  = $pd.Relatorio
        $c.chkDrivers.IsChecked    = $pd.Drivers
        $c.txtModelo.Text = "Detectado: $($d.Modelo)`n$($d.FerramentaDrivers)"
        if (-not $d.Notebook) { $c.chkPlanoNotebook.Visibility = 'Collapsed' }
        $c.chkPlanoNotebook.IsEnabled = $false

        # ---------- comportamento dos controles ----------
        $c.rbGamer.Add_Checked({
            foreach ($cb in $checksGamer) { $cb.IsChecked = $true }
            $c.chkOtimizar.IsChecked = $true
            $c.chkXbox.IsChecked = $false
            $c.chkXbox.IsEnabled = $false
        })
        $c.rbPadrao.Add_Checked({
            foreach ($cb in $checksGamer) { $cb.IsChecked = $false }
            $c.chkOtimizar.IsChecked = $false
            $c.chkXbox.IsEnabled = [bool]$c.chkDebloat.IsChecked
        })
        $atualizaDebloat = {
            $on = [bool]$c.chkDebloat.IsChecked
            $c.chkXbox.IsEnabled = $on -and -not [bool]$c.rbGamer.IsChecked
            $c.chkOneDrive.IsEnabled = $on
        }
        $c.chkDebloat.Add_Checked($atualizaDebloat)
        $c.chkDebloat.Add_Unchecked($atualizaDebloat)
        $c.chkOtimizar.Add_Checked({ $c.chkPlanoNotebook.IsEnabled = $true })
        $c.chkOtimizar.Add_Unchecked({ $c.chkPlanoNotebook.IsEnabled = $false })

        $c.btnIniciar.Add_Click({
            $ids = @()
            foreach ($painel in $c.spProgramas, $c.spGamer, $c.spOpcionais) {
                foreach ($cb in $painel.Children) { if ($cb.IsChecked) { $ids += [string]$cb.Tag } }
            }
            $off = $c.cbOffice.SelectedItem
            $resumo = "Programas: $($ids.Count) selecionado(s)`nOffice: $($off.Content)"
            if ($c.chkDebloat.IsChecked)  { $resumo += "`nDebloat: sim" }
            if ($c.chkOtimizar.IsChecked) { $resumo += "`nOtimização avançada: sim" }
            if ($c.chkDrivers.IsChecked)  { $resumo += "`nDrivers: sim" }
            if ($c.chkWU.IsChecked)       { $resumo += "`nDesativar Windows Update: sim" }
            $r = [System.Windows.MessageBox]::Show($janela, "Iniciar a instalação?`n`n$resumo", 'SetupPC', 'YesNo', 'Question')
            if ("$r" -ne 'Yes') { return }
            $cfg = @{
                Gamer         = [bool]$c.rbGamer.IsChecked
                Programas     = $ids
                Office        = [string]$off.Tag
                Debloat       = [bool]$c.chkDebloat.IsChecked
                RemoverXbox   = [bool]$c.chkXbox.IsChecked
                RemoverOneDrive = [bool]$c.chkOneDrive.IsChecked
                Restauracao   = [bool]$c.chkRestauracao.IsChecked
                Icones        = [bool]$c.chkIcones.IsChecked
                Rede          = [bool]$c.chkRede.IsChecked
                Hibernacao    = [bool]$c.chkHibernacao.IsChecked
                Otimizar      = [bool]$c.chkOtimizar.IsChecked
                PlanoNotebook = [bool]$c.chkPlanoNotebook.IsChecked
                DesativarWU   = [bool]$c.chkWU.IsChecked
                Relatorio     = [bool]$c.chkRelatorio.IsChecked
                Drivers       = [bool]$c.chkDrivers.IsChecked
            }
            $c.wpFinal.Visibility = 'Collapsed'
            $c.rtbLog.Document.Blocks.Clear()
            Send-Comando 'Instalar' $cfg
        })

        # Ferramentas
        $c.btnRelatorioPC.Add_Click({ Send-Comando 'RelatorioPC' })
        $c.btnLimpeza.Add_Click({
            $r = [System.Windows.MessageBox]::Show($janela, "Fazer a limpeza de disco agora?`n`nApaga temporários, cache do Windows Update, lixeira e arquivos antigos do Windows. A pasta Downloads não é tocada. Pode levar alguns minutos.", 'SetupPC', 'YesNo', 'Question')
            if ("$r" -eq 'Yes') { Send-Comando 'Limpeza' }
        })
        $c.btnPasta.Add_Click({ Start-Process explorer.exe $d.PastaScript })
        $c.btnDriversVerificar.Add_Click({ Send-Comando 'DriversVerificar' })
        $c.btnDriversInstalar.Add_Click({
            $r = [System.Windows.MessageBox]::Show($janela, "Procurar e instalar drivers agora?`n`nUsa a ferramenta oficial do fabricante (Dell, Lenovo ou HP) quando existir e depois o Windows Update. Pode levar vários minutos.", 'SetupPC', 'YesNo', 'Question')
            if ("$r" -eq 'Yes') { Send-Comando 'DriversInstalar' }
        })
        $c.btnDriverBooster.Add_Click({ Send-Comando 'DriverBooster' @{ Desinstalar = [bool]$c.chkDBDesinstalar.IsChecked } })
        $c.btnOtimizar.Add_Click({
            $r = [System.Windows.MessageBox]::Show($janela, 'Aplicar a otimização avançada? Os valores originais ficam salvos e podem ser revertidos.', 'SetupPC', 'YesNo', 'Question')
            if ("$r" -eq 'Yes') { Send-Comando 'Otimizar' }
        })
        $c.btnVBS.Add_Click({ Send-Comando 'VBS' })
        $c.btnReverter.Add_Click({
            $r = [System.Windows.MessageBox]::Show($janela, 'Reverter todas as otimizações e voltar como estava antes?', 'SetupPC', 'YesNo', 'Question')
            if ("$r" -eq 'Yes') { Send-Comando 'Reverter' }
        })
        $c.btnWUDesativar.Add_Click({
            $r = [System.Windows.MessageBox]::Show($janela, "Desativar o Windows Update?`n`nO PC deixa de receber correções de segurança automaticamente. A Microsoft Store continua funcionando.", 'SetupPC', 'YesNo', 'Warning')
            if ("$r" -eq 'Yes') { Send-Comando 'WUDesativar' }
        })
        $c.btnWUReativar.Add_Click({ Send-Comando 'WUReativar' })
        $c.btnWUStatus.Add_Click({ Send-Comando 'WUStatus' })
        $c.btnAtivStatus.Add_Click({ Send-Comando 'AtivStatus' })
        $c.btnAtivOEM.Add_Click({ Send-Comando 'AtivOEM' })
        $c.btnAtivWin.Add_Click({ Send-Comando 'AtivWin' })
        $c.btnAtivOffice.Add_Click({ Send-Comando 'AtivOffice' })
        $c.btnAtivTela.Add_Click({ Start-Process 'ms-settings:activation' })

        # Botões do final
        $c.btnAbrirEntrega.Add_Click({ if ($sync.ArquivoEntrega) { Start-Process $sync.ArquivoEntrega } })
        $c.btnIrAtivacao.Add_Click({ $c.tabs.SelectedItem = $c.tabFerramentas })
        $c.btnReiniciar.Add_Click({
            $r = [System.Windows.MessageBox]::Show($janela, 'Reiniciar o computador agora?', 'SetupPC', 'YesNo', 'Question')
            if ("$r" -eq 'Yes') { Send-Comando 'Reiniciar' }
        })

        $botoes = @($c.btnIniciar, $c.btnDriversVerificar, $c.btnDriversInstalar, $c.btnDriverBooster, $c.btnRelatorioPC, $c.btnLimpeza, $c.btnOtimizar, $c.btnVBS, $c.btnReverter,
                    $c.btnWUDesativar, $c.btnWUReativar, $c.btnWUStatus, $c.btnAtivStatus, $c.btnAtivOEM,
                    $c.btnAtivWin, $c.btnAtivOffice, $c.btnReiniciar)

        # ---------- atualização periódica (log, progresso, perguntas) ----------
        $script:perguntando = $false
        $timer = New-Object System.Windows.Threading.DispatcherTimer
        $timer.Interval = [TimeSpan]::FromMilliseconds(150)
        $timer.Add_Tick({
            $it = $null; $n = 0
            while ($n -lt 300 -and $sync.Log.TryDequeue([ref]$it)) { Add-LinhaLog $it; $n++ }
            if ($n -gt 0) { $c.rtbLog.ScrollToEnd() }

            $ocupado = [bool]$sync.Ocupado
            foreach ($b in $botoes) { $b.IsEnabled = -not $ocupado }
            if ($ocupado) {
                $c.pbAtividade.Visibility = 'Visible'
                $c.pbAtividade.IsIndeterminate = $true
            } else {
                $c.pbAtividade.IsIndeterminate = $false
                $c.pbAtividade.Visibility = 'Hidden'
            }
            if ($sync.Etapa) { $c.txtEtapa.Text = $sync.Etapa }
            if ($sync.TotalPassos -gt 0) {
                if ($sync.Concluido) { $pct = 100 }
                else { $pct = [math]::Min(99, [math]::Round(([math]::Max(0, $sync.PassoAtual - 1)) / $sync.TotalPassos * 100)) }
                $c.pbGeral.Value = $pct
                $c.txtPct.Text = "$pct%"
            }

            if ($sync.Pergunta -and -not $script:perguntando) {
                $script:perguntando = $true
                try {
                    $q = $sync.Pergunta
                    if ($q.Tipo -eq 'Aviso') {
                        [void][System.Windows.MessageBox]::Show($janela, $q.Texto, 'SetupPC', 'OK', 'Information')
                        $resp = $true
                    } elseif ($q.Tipo -eq 'SimNao') {
                        $r = [System.Windows.MessageBox]::Show($janela, $q.Texto, 'SetupPC', 'YesNo', 'Question')
                        $resp = ("$r" -eq 'Yes')
                    } else {
                        $resp = Show-Entrada $q.Texto
                    }
                    $sync.Resposta = $resp
                    $sync.Pergunta = $null
                    $sync.Respondida = $true
                } finally { $script:perguntando = $false }
            }

            if ($sync.MostrarFinal) {
                $sync.MostrarFinal = $false
                $c.wpFinal.Visibility = 'Visible'
                if ($sync.ArquivoEntrega) { $c.btnAbrirEntrega.Visibility = 'Visible' } else { $c.btnAbrirEntrega.Visibility = 'Collapsed' }
                [void][System.Windows.MessageBox]::Show($janela, "Instalação concluída!`n`nConfira o resultado na aba Progresso. Reinicie o computador para terminar de aplicar as configurações.", 'SetupPC', 'OK', 'Information')
            }
        })

        $janela.Add_Loaded({ $sync.Aberta = $true; $timer.Start() })
        $janela.Add_Closing({
            param($remetente, $e)
            if ($sync.Ocupado) {
                $r = [System.Windows.MessageBox]::Show($janela, "Uma tarefa ainda está em andamento e vai continuar em segundo plano até terminar.`n`nFechar a janela mesmo assim?", 'SetupPC', 'YesNo', 'Warning')
                if ("$r" -ne 'Yes') { $e.Cancel = $true }
            }
        })
        [void]$janela.ShowDialog()
    } catch {
        $sync.ErroGui = $_.Exception.Message
    } finally {
        $sync.Fechado = $true
    }
}

# Aplica as escolhas feitas na interface nas variáveis do script
function Set-ConfigDaGui($Cfg) {
    $script:Gamer               = [switch]$Cfg.Gamer
    $script:UsarSelecao         = $true
    $script:IdsSelecionados     = @($Cfg.Programas)
    $script:SemProgramas        = [switch]$false
    if ($Cfg.Office) { $script:OfficeEscolhido = $VersoesOffice[$Cfg.Office] } else { $script:OfficeEscolhido = $null }
    $script:SemDebloat          = [switch](-not $Cfg.Debloat)
    $script:RemoverXbox         = $Cfg.RemoverXbox -and -not $Cfg.Gamer
    $script:RemoverOneDrive     = [bool]$Cfg.RemoverOneDrive
    $script:SemPontoRestauracao = [switch](-not $Cfg.Restauracao)
    $script:IconesAreaTrabalho  = [bool]$Cfg.Icones
    $script:ConfigurarRede      = [bool]$Cfg.Rede
    $script:DesativarHibernacao = [bool]$Cfg.Hibernacao
    $script:Otimizar            = [switch]$Cfg.Otimizar
    if ($Cfg.PlanoNotebook) { $script:PlanoNotebook = 'S' } else { $script:PlanoNotebook = 'N' }
    $script:DesativarWindowsUpdate = [switch]$Cfg.DesativarWU
    $script:GerarRelatorioEntrega  = [bool]$Cfg.Relatorio
    $script:InstalarDrivers        = [bool]$Cfg.Drivers
    if ($Cfg.Gamer) { $modo = 'Gamer' } else { $modo = 'Padrão' }
    if (-not $Cfg.Office) { $modo += ' sem Office' }
    $script:ModoNome = $modo
}

# Executa um comando vindo da interface (roda no script principal)
function Invoke-ComandoGui($Cmd) {
    $s = $script:Sync
    $s.Ocupado = $true
    $s.Concluido = $false
    $s.PassoAtual = 1
    $s.TotalPassos = 1
    $nomes = @{
        RelatorioPC = 'Gerando o relatório do PC'; Limpeza = 'Limpeza de disco'; Otimizar = 'Aplicando a otimização avançada'
        VBS = 'Integridade de Memória (VBS)'; Reverter = 'Revertendo as otimizações'; WUDesativar = 'Desativando o Windows Update'
        WUReativar = 'Reativando o Windows Update'; WUStatus = 'Status do Windows Update'; AtivStatus = 'Status da ativação'
        AtivOEM = 'Ativando o Windows com a chave da BIOS'; AtivWin = 'Ativando o Windows'; AtivOffice = 'Ativando o Office'
        Instalar = 'Iniciando a instalação'; Reiniciar = 'Reiniciando'
        DriversVerificar = 'Verificando drivers'; DriversInstalar = 'Instalando drivers'; DriverBooster = 'Driver Booster'
    }
    $s.Etapa = $nomes[$Cmd.Acao]
    try {
        switch ($Cmd.Acao) {
            'Instalar'    { Set-ConfigDaGui $Cmd.Config; Invoke-Instalacao; $s.ArquivoEntrega = $script:ArquivoEntrega }
            'RelatorioPC' { New-RelatorioPC }
            'Limpeza'     { Invoke-Limpeza }
            'DriversVerificar' { [void](Show-DispositivosSemDriver) }
            'DriversInstalar'  { Invoke-Drivers }
            'DriverBooster'    { Invoke-DriverBooster ([bool]$Cmd.Config.Desinstalar) }
            'Otimizar'    { Invoke-Otimizacoes }
            'VBS'         { Disable-VBS }
            'Reverter'    { Undo-Otimizacoes }
            'WUDesativar' { Disable-WindowsUpdate }
            'WUReativar'  { Enable-WindowsUpdate }
            'WUStatus'    { Show-StatusWindowsUpdate }
            'AtivStatus'  { Show-StatusAtivacao }
            'AtivOEM'     { Invoke-AtivarWindowsOEM }
            'AtivWin'     { $k = Read-Chave 'Digite a chave do Windows'; if ($k) { Invoke-AtivarWindowsChave $k } }
            'AtivOffice'  { Invoke-AtivarOfficeChave }
            'Reiniciar'   { Stop-Transcript | Out-Null; Restart-Computer -Force }
        }
    } catch {
        Write-Erro "Erro: $($_.Exception.Message)"
    } finally {
        $s.Concluido = $true
        if ($Cmd.Acao -eq 'Instalar') { $s.Etapa = 'Instalação concluída.'; $s.MostrarFinal = $true }
        else { $s.Etapa = "$($s.Etapa): concluído." }
        $s.Ocupado = $false
    }
}

function Start-Gui {
    $s = [hashtable]::Synchronized(@{})
    $s.Log         = New-Object 'System.Collections.Concurrent.ConcurrentQueue[object]'
    $s.Comandos    = New-Object 'System.Collections.Concurrent.ConcurrentQueue[object]'
    $s.Xaml        = $XamlPrincipal
    $s.XamlEntrada = $XamlEntrada
    $s.Fechado = $false; $s.Aberta = $false; $s.Ocupado = $false
    $os = Get-CimInstance Win32_OperatingSystem -ErrorAction SilentlyContinue
    $fab = Get-Fabricante
    $ferramenta = @{ Dell = 'Drivers via Dell Command Update + Windows Update'; Lenovo = 'Drivers via catálogo Lenovo + Windows Update'; HP = 'Drivers via HP Image Assistant + Windows Update' }[$fab.Marca]
    if (-not $ferramenta) { $ferramenta = 'Sem ferramenta do fabricante: drivers via Windows Update' }
    $s.Dados = @{
        PC           = $env:COMPUTERNAME
        Windows      = "$($os.Caption)"
        Notebook     = (Test-Notebook)
        PastaScript  = $PastaScript
        Programas    = @($Programas          | ForEach-Object { @{ Nome = $_.Nome; Id = $_.Id } })
        Gamer        = @($ProgramasGamer     | ForEach-Object { @{ Nome = $_.Nome; Id = $_.Id } })
        Opcionais    = @($ProgramasOpcionais | ForEach-Object { @{ Nome = $_.Nome; Id = $_.Id } })
        Office       = @($VersoesOffice.Keys | ForEach-Object { @{ Chave = $_; Nome = $VersoesOffice[$_].Nome } })
        OfficePadrao = $OfficePadrao
        Modelo       = "$($fab.Fabricante) $($fab.Modelo)"
        FerramentaDrivers = $ferramenta
        Padroes      = @{
            RemoverXbox = $RemoverXbox; RemoverOneDrive = $RemoverOneDrive; Icones = $IconesAreaTrabalho
            Rede = $ConfigurarRede; Hibernacao = $DesativarHibernacao; Relatorio = $GerarRelatorioEntrega
            Drivers = $InstalarDrivers
        }
    }

    $rs = [runspacefactory]::CreateRunspace()
    $rs.ApartmentState = 'STA'
    $rs.ThreadOptions  = 'ReuseThread'
    $rs.Open()
    $rs.SessionStateProxy.SetVariable('sync', $s)
    $ps = [powershell]::Create()
    $ps.Runspace = $rs
    [void]$ps.AddScript($GuiScript.ToString())
    $handle = $ps.BeginInvoke()

    # Espera a janela abrir (até 20 s)
    for ($i = 0; $i -lt 200 -and -not $s.Aberta -and -not $s.Fechado; $i++) { Start-Sleep -Milliseconds 100 }
    if (-not $s.Aberta) {
        $script:ErroGui = $s.ErroGui
        try { $ps.Stop(); $rs.Close() } catch {}
        return $false
    }

    $script:Sync = $s
    Set-Console $false
    Write-Info "SetupPC aberto em $env:COMPUTERNAME. Log: $ArquivoLog"

    # Laço principal: executa os comandos clicados na janela
    while (-not $s.Fechado) {
        $cmd = $null
        if ($s.Comandos.TryDequeue([ref]$cmd)) { Invoke-ComandoGui $cmd }
        else { Start-Sleep -Milliseconds 120 }
    }
    $script:Sync = $null
    try { [void]$ps.EndInvoke($handle); $rs.Close() } catch {}
    return $true
}

# ---------------------------------------------------------------------
#  Execução
# ---------------------------------------------------------------------
$usarParametros = $Silencioso -or $Office -or $SemProgramas -or $SemDebloat -or $SemPontoRestauracao -or $Gamer -or $DesativarWindowsUpdate -or $Otimizar

# Padrão: interface gráfica. Se não abrir, cai para o modo texto.
if (-not $usarParametros -and -not $Console) {
    if (Start-Gui) { Stop-Transcript | Out-Null; exit }
    Set-Console $true
    Write-Aviso "Não foi possível abrir a interface gráfica ($script:ErroGui). Usando o modo texto."
    Start-Sleep -Seconds 2
}

Show-Banner

# Modo texto: menu principal
if (-not $usarParametros) {
    do {
        $op = Show-MenuPrincipal
        if ($op -eq '5') { Invoke-MenuOtimizacao;    Show-Banner }
        if ($op -eq '6') { Invoke-MenuWindowsUpdate; Show-Banner }
        if ($op -eq '7') { Invoke-MenuAtivacao;      Show-Banner }
        if ($op -eq '10') { Invoke-MenuDrivers;      Show-Banner }
        if ($op -in '8', '9') {
            if ($op -eq '8') { New-RelatorioPC } else { Invoke-Limpeza }
            Write-Host ''; Read-Host '   Enter para voltar ao menu' | Out-Null
            Show-Banner
        }
        if ($op -eq '0') { Stop-Transcript | Out-Null; exit }
    } until ($op -in '1', '2', '3', '4')
    if ($op -in '3', '4') { $Gamer = [switch]$true }
    if ($op -in '2', '4') { $Office = '0' }
    $script:ModoNome = @{ '1' = 'Padrão'; '2' = 'Padrão sem Office'; '3' = 'Gamer'; '4' = 'Gamer sem Office' }[$op]
    if ($Gamer) {
        $RemoverXbox = $false   # modo gamer nunca remove os apps Xbox
        if (Confirm-Usuario 'Aplicar também a otimização avançada para jogos no final?') {
            $Otimizar = [switch]$true
            if (Test-Notebook) {   # pergunta agora para não travar a instalação depois
                if (Confirm-Usuario 'É um notebook: usar o plano Desempenho Máximo (gasta mais bateria)?') { $script:PlanoNotebook = 'S' } else { $script:PlanoNotebook = 'N' }
            }
        }
    }
}

# Pergunta tudo logo no início, para o resto rodar sem precisar de ninguém
$script:OfficeEscolhido = Select-Office
Invoke-Instalacao

if (-not $Silencioso) {
    Write-Host ''
    if (Confirm-Usuario 'Abrir o menu de ativação agora?') { Invoke-MenuAtivacao }
}

Stop-Transcript | Out-Null

if (-not $Silencioso) {
    Write-Host ''
    if (Confirm-Usuario 'Reiniciar o computador agora?') { Restart-Computer -Force }
}
