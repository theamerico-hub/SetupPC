# SetupPC

Script de pós-formatação para Windows 10/11, feito por **Danilo Américo**.

Instala os programas padrão, o Microsoft Office (em português do Brasil) e faz o debloat do sistema, tudo por uma interface gráfica simples.

## Como usar

Abra o **PowerShell** e cole:

```powershell
irm https://raw.githubusercontent.com/theamerico-hub/SetupPC/main/instalar.ps1 | iex
```

Ou baixe a pasta e dê dois cliques em `Executar.bat` (ou `Executar (modo texto).bat` para o menu no console).

## O que ele faz

- **Programas:** Chrome, WinRAR, AnyDesk, Acrobat Reader + opcionais, via winget
- **Modo Gamer:** Steam, Epic Games, Discord e otimizações para jogos
- **Office:** escolha entre Microsoft 365, 2024, 2021, 2019, 2016 ou só Word/Excel/PowerPoint (instalação oficial da Microsoft; a ativação usa a licença do cliente)
- **Debloat:** remove apps desnecessários, telemetria, anúncios, Copilot e Widgets
- **Configurações:** ícones na Área de Trabalho, rede privada com compartilhamento, hibernação
- **Ferramentas:** relatório do PC, limpeza de disco, otimização avançada (reversível), Windows Update, ativação com chave
- **Relatório de entrega** em HTML para o cliente

Logs e relatórios ficam em `C:\SetupPC` (ou na pasta do pendrive).
