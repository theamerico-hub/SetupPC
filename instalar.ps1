# SetupPC - inicializador rapido
# Uso (PowerShell):  irm https://raw.githubusercontent.com/theamerico-hub/SetupPC/main/instalar.ps1 | iex
#
# Baixa a versao mais recente do SetupPC para C:\SetupPC e abre a interface grafica como administrador.
# (Tudo dentro de & { } para nao fechar o PowerShell de quem executou.)

& {
    $ErrorActionPreference = 'Stop'
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    $ProgressPreference = 'SilentlyContinue'

    $url   = 'https://raw.githubusercontent.com/theamerico-hub/SetupPC/main/SetupPC.ps1'
    $pasta = Join-Path $env:SystemDrive 'SetupPC'
    $arq   = Join-Path $pasta 'SetupPC.ps1'

    try {
        Write-Host ''
        Write-Host '  SetupPC - feito por Danilo Americo' -ForegroundColor Cyan
        Write-Host '  Baixando a versao mais recente...' -ForegroundColor Gray
        New-Item -ItemType Directory -Path $pasta -Force | Out-Null
        Invoke-WebRequest -Uri $url -OutFile $arq -UseBasicParsing
        Unblock-File -Path $arq

        Write-Host '  Abrindo (confirme a permissao de administrador)...' -ForegroundColor Gray
        Start-Process powershell.exe -Verb RunAs -WindowStyle Hidden `
            -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$arq`""
        Write-Host '  Pronto! Logs e relatorios ficam em C:\SetupPC' -ForegroundColor Green
    } catch {
        Write-Host "  Nao foi possivel abrir o SetupPC: $($_.Exception.Message)" -ForegroundColor Red
    }
}
