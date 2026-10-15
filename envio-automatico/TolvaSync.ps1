# =============================================================================
#  Envio automático da planilha GESTÃO TOLVA para o painel (GitHub)
#  Goiás Verde Alimentos — PCM
#
#  Roda escondido a cada 5 minutos (tarefa agendada do Windows). Se a planilha
#  foi salva desde o último envio, manda a versão nova para o GitHub; o GitHub
#  Actions reprocessa e o painel atualiza sozinho em 1 a 2 minutos.
#
#  Uso:  instalar.bat     -> escolhe a planilha, pede o token e agenda a tarefa
#        enviar-agora.bat -> envia na hora (mostra o resultado na tela)
#        desinstalar.bat  -> remove a tarefa e as configurações
# =============================================================================
param(
    [switch]$Instalar,
    [switch]$Desinstalar,
    [switch]$Agora
)

$ErrorActionPreference = 'Stop'

# ----- Configuração fixa ------------------------------------------------------
$Repo          = 'Geanfer10/Gestao-tolva'
$Branch        = 'main'
$Destino       = 'uploads/GESTAO_TOLVA.xlsm'      # nome fixo no GitHub
$IntervaloMin  = 5
$EsperaSegundos = 20                               # espera o Excel terminar de salvar
$NomeTarefa    = 'Tolva - envio automatico da planilha'

$Pasta     = Join-Path $env:LOCALAPPDATA 'TolvaSync'
$ArqConfig = Join-Path $Pasta 'config.json'
$ArqToken  = Join-Path $Pasta 'token.dat'
$ArqLog    = Join-Path $Pasta 'envios.log'
$ArqScript = Join-Path $Pasta 'TolvaSync.ps1'
$ArqVbs    = Join-Path $Pasta 'rodar-escondido.vbs'
$ApiBase   = "https://api.github.com/repos/$Repo"

# ----- Utilitários ------------------------------------------------------------
function Escrever-Log([string]$Mensagem, [switch]$SoNoArquivo) {
    $linha = '{0}  {1}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Mensagem
    if (($Agora -or $Instalar) -and -not $SoNoArquivo) { Write-Host $Mensagem }
    try {
        if (-not (Test-Path $Pasta)) { New-Item -ItemType Directory -Path $Pasta -Force | Out-Null }
        Add-Content -Path $ArqLog -Value $linha -Encoding UTF8
        # mantém só as últimas 500 linhas
        $todas = @(Get-Content -Path $ArqLog -Encoding UTF8)
        if ($todas.Count -gt 500) { $todas[-500..-1] | Set-Content -Path $ArqLog -Encoding UTF8 }
    } catch { }
}

function Ler-Config {
    if (-not (Test-Path $ArqConfig)) { return $null }
    return (Get-Content -Path $ArqConfig -Raw -Encoding UTF8 | ConvertFrom-Json)
}

function Salvar-Config($Config) {
    $Config | ConvertTo-Json | Set-Content -Path $ArqConfig -Encoding UTF8
}

function Preparar-Rede {
    # TLS 1.2 (obrigatório para o GitHub no Windows PowerShell 5.1) e proxy da empresa
    [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
    try { [Net.WebRequest]::DefaultWebProxy.Credentials = [Net.CredentialCache]::DefaultNetworkCredentials } catch { }
}

function Ler-Token {
    if (-not (Test-Path $ArqToken)) { throw 'Token do GitHub não configurado. Rode o instalar.bat.' }
    # O token fica criptografado com a conta do Windows (só este usuário, neste PC, consegue ler)
    # .Trim(): o arquivo termina com quebra de linha, que o ConvertTo-SecureString não aceita
    $seguro = (Get-Content -Path $ArqToken -Raw).Trim() | ConvertTo-SecureString
    $bstr = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($seguro)
    try { return [Runtime.InteropServices.Marshal]::PtrToStringBSTR($bstr) }
    finally { [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr) }
}

function Cabecalhos([string]$Token) {
    return @{
        'Authorization'        = "Bearer $Token"
        'Accept'               = 'application/vnd.github+json'
        'X-GitHub-Api-Version' = '2022-11-28'
        'User-Agent'           = 'TolvaSync'
    }
}

function Codigo-Http($Erro) {
    try { return [int]$Erro.Exception.Response.StatusCode } catch { return 0 }
}

function Ler-Bytes([string]$Caminho) {
    # Abre permitindo que o Excel continue com o arquivo aberto
    $fs = [IO.File]::Open($Caminho, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::ReadWrite)
    try {
        $ms = New-Object IO.MemoryStream
        $fs.CopyTo($ms)
        return ,$ms.ToArray()
    } finally { $fs.Dispose() }
}

function Hash-Sha256([byte[]]$Bytes) {
    $sha = [Security.Cryptography.SHA256]::Create()
    try { return ([BitConverter]::ToString($sha.ComputeHash($Bytes)) -replace '-', '').ToLower() }
    finally { $sha.Dispose() }
}

function Sha-Git([byte[]]$Bytes) {
    # Mesmo identificador que o GitHub usa para o arquivo (blob SHA-1)
    $cab = [Text.Encoding]::ASCII.GetBytes("blob $($Bytes.Length)`0")
    $tudo = New-Object byte[] ($cab.Length + $Bytes.Length)
    [Array]::Copy($cab, 0, $tudo, 0, $cab.Length)
    [Array]::Copy($Bytes, 0, $tudo, $cab.Length, $Bytes.Length)
    $sha1 = [Security.Cryptography.SHA1]::Create()
    try { return ([BitConverter]::ToString($sha1.ComputeHash($tudo)) -replace '-', '').ToLower() }
    finally { $sha1.Dispose() }
}

function Avisar-Erro([string]$Texto) {
    # Balão no canto da tela, no máximo uma vez por dia (para o token vencido não passar despercebido)
    $cfg = Ler-Config
    $hoje = Get-Date -Format 'yyyy-MM-dd'
    if ($cfg -and $cfg.PSObject.Properties['ultimo_aviso'] -and $cfg.ultimo_aviso -eq $hoje) { return }
    try {
        Add-Type -AssemblyName System.Windows.Forms
        Add-Type -AssemblyName System.Drawing
        $n = New-Object System.Windows.Forms.NotifyIcon
        $n.Icon = [System.Drawing.SystemIcons]::Warning
        $n.Visible = $true
        $n.ShowBalloonTip(15000, 'Painel da Tolva', $Texto, [System.Windows.Forms.ToolTipIcon]::Warning)
        Start-Sleep -Seconds 15
        $n.Dispose()
    } catch { }
    if ($cfg) {
        $cfg | Add-Member -NotePropertyName ultimo_aviso -NotePropertyValue $hoje -Force
        Salvar-Config $cfg
    }
}

# ----- Envio ------------------------------------------------------------------
function Enviar-Planilha([switch]$Forcar) {
    $cfg = Ler-Config
    if (-not $cfg) { throw 'Configuração não encontrada. Rode o instalar.bat.' }
    $arquivo = $cfg.arquivo

    if (-not (Test-Path -LiteralPath $arquivo)) {
        Escrever-Log "Planilha não encontrada em: $arquivo"
        try { Avisar-Erro 'Não encontrei a planilha da Tolva. Ela foi movida ou renomeada? Rode o instalar.bat de novo.' } catch { }
        return
    }

    $info = Get-Item -LiteralPath $arquivo
    $idade = (Get-Date) - $info.LastWriteTime
    if (-not $Forcar -and $idade.TotalSeconds -lt $EsperaSegundos) { return }   # ainda salvando; tenta na próxima

    $bytes = Ler-Bytes $arquivo
    # Todo .xlsm é um zip: começa com 'PK'. Se não começar, o arquivo está incompleto.
    if ($bytes.Length -lt 4 -or $bytes[0] -ne 0x50 -or $bytes[1] -ne 0x4B) {
        Escrever-Log 'Arquivo incompleto ou corrompido; não enviado (tento de novo na próxima rodada).'
        return
    }

    $hash = Hash-Sha256 $bytes
    $ultimo = $null
    if ($cfg.PSObject.Properties['ultimo_hash']) { $ultimo = $cfg.ultimo_hash }
    if (-not $Forcar -and $hash -eq $ultimo) { return }   # nada mudou desde o último envio

    Preparar-Rede
    $cab = Cabecalhos (Ler-Token)
    $urlArquivo = "$ApiBase/contents/$Destino"

    # Versão que está hoje no GitHub (necessária para substituir o arquivo)
    $shaRemoto = $null
    try {
        $remoto = Invoke-RestMethod -Method Get -Uri "$urlArquivo`?ref=$Branch" -Headers $cab
        $shaRemoto = $remoto.sha
    } catch {
        $cod = Codigo-Http $_
        if ($cod -ne 404) { throw }
    }

    if ($shaRemoto -and $shaRemoto -eq (Sha-Git $bytes)) {
        $cfg | Add-Member -NotePropertyName ultimo_hash -NotePropertyValue $hash -Force
        Salvar-Config $cfg
        Escrever-Log 'A planilha do GitHub já é igual à do computador; nada a enviar.'
        return
    }

    $corpo = @{
        message = "Atualiza planilha da Tolva (envio automático - $env:COMPUTERNAME)"
        content = [Convert]::ToBase64String($bytes)
        branch  = $Branch
    }
    if ($shaRemoto) { $corpo.sha = $shaRemoto }
    $json = $corpo | ConvertTo-Json -Compress

    Invoke-RestMethod -Method Put -Uri $urlArquivo -Headers $cab `
        -ContentType 'application/json; charset=utf-8' `
        -Body ([Text.Encoding]::UTF8.GetBytes($json)) | Out-Null

    $cfg | Add-Member -NotePropertyName ultimo_hash  -NotePropertyValue $hash -Force
    $cfg | Add-Member -NotePropertyName ultimo_envio -NotePropertyValue (Get-Date -Format 'yyyy-MM-dd HH:mm:ss') -Force
    Salvar-Config $cfg
    Escrever-Log ("Planilha enviada ({0:N0} KB). O painel atualiza em 1 a 2 minutos." -f ($bytes.Length / 1KB))
}

# ----- Instalação -------------------------------------------------------------
function Escolher-Planilha {
    Add-Type -AssemblyName System.Windows.Forms
    $dlg = New-Object System.Windows.Forms.OpenFileDialog
    $dlg.Title = 'Escolha a planilha GESTÃO TOLVA'
    $dlg.Filter = 'Planilha Excel (*.xlsm;*.xlsx)|*.xlsm;*.xlsx'
    if ($dlg.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) { throw 'Nenhuma planilha escolhida.' }
    return $dlg.FileName
}

function Instalar-Tarefa {
    if (-not (Test-Path $Pasta)) { New-Item -ItemType Directory -Path $Pasta -Force | Out-Null }

    # Cópia do script numa pasta fixa (a pasta baixada pode ser apagada depois)
    if ($PSCommandPath -ne $ArqScript) { Copy-Item -LiteralPath $PSCommandPath -Destination $ArqScript -Force }
    try { Unblock-File -LiteralPath $ArqScript } catch { }

    # Lançador VBS: roda o PowerShell sem abrir janela a cada 5 minutos
    $vbs = 'CreateObject("WScript.Shell").Run "powershell.exe -NoProfile -ExecutionPolicy Bypass -File ""' + $ArqScript + '""", 0, False'
    Set-Content -Path $ArqVbs -Value $vbs -Encoding ASCII

    $acao     = New-ScheduledTaskAction -Execute 'wscript.exe' -Argument ('"' + $ArqVbs + '"')
    $repetir  = New-ScheduledTaskTrigger -Once -At (Get-Date).AddMinutes(1) `
                    -RepetitionInterval (New-TimeSpan -Minutes $IntervaloMin) `
                    -RepetitionDuration (New-TimeSpan -Days 3650)
    $logon    = New-ScheduledTaskTrigger -AtLogOn -User "$env:USERDOMAIN\$env:USERNAME"
    $ajustes  = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries `
                    -StartWhenAvailable -MultipleInstances IgnoreNew `
                    -ExecutionTimeLimit (New-TimeSpan -Minutes 10)
    $usuario  = New-ScheduledTaskPrincipal -UserId "$env:USERDOMAIN\$env:USERNAME" -LogonType Interactive -RunLevel Limited

    Register-ScheduledTask -TaskName $NomeTarefa -Action $acao -Trigger @($repetir, $logon) `
        -Settings $ajustes -Principal $usuario -Force | Out-Null
}

function Instalar {
    Write-Host ''
    Write-Host '=== Envio automático da planilha da Tolva ===' -ForegroundColor Green
    Write-Host ''
    Write-Host '1) Escolha a planilha GESTÃO TOLVA na janela que vai abrir...'
    $arquivo = Escolher-Planilha
    Write-Host "   Planilha: $arquivo"
    Write-Host ''
    Write-Host '2) Cole o token do GitHub (o texto não aparece enquanto você cola) e aperte Enter:'
    $seguro = Read-Host -AsSecureString
    if (-not $seguro -or $seguro.Length -lt 20) { throw 'Token vazio ou curto demais.' }

    if (-not (Test-Path $Pasta)) { New-Item -ItemType Directory -Path $Pasta -Force | Out-Null }
    $seguro | ConvertFrom-SecureString | Set-Content -Path $ArqToken

    # Testa se o token enxerga o repositório
    Preparar-Rede
    $cabTeste = Cabecalhos (Ler-Token)
    try {
        Invoke-RestMethod -Method Get -Uri $ApiBase -Headers $cabTeste | Out-Null
    } catch {
        Remove-Item $ArqToken -Force -ErrorAction SilentlyContinue
        $cod = Codigo-Http $_
        if ($cod -eq 401) { throw 'O GitHub recusou o token (inválido ou vencido). Gere outro e rode o instalar.bat de novo.' }
        if ($cod -eq 404) { throw "O token não tem acesso ao repositório $Repo. Ao criar o token, selecione esse repositório." }
        throw "Não consegui falar com o GitHub: $($_.Exception.Message)"
    }

    # Preserva o último envio se já estava instalado e a planilha é a mesma
    $cfg = Ler-Config
    if (-not $cfg -or $cfg.arquivo -ne $arquivo) { $cfg = [pscustomobject]@{ arquivo = $arquivo } }
    Salvar-Config $cfg

    Write-Host ''
    Write-Host '3) Agendando a verificação a cada 5 minutos...'
    Instalar-Tarefa

    Write-Host ''
    Write-Host '4) Enviando a planilha agora para testar...'
    Enviar-Planilha -Forcar

    Write-Host ''
    Write-Host 'Pronto! A partir de agora é só salvar a planilha no Excel.' -ForegroundColor Green
    Write-Host 'O painel atualiza sozinho em até uns 7 minutos depois de salvar.'
    Write-Host "Histórico de envios: $ArqLog"
}

function Desinstalar {
    Unregister-ScheduledTask -TaskName $NomeTarefa -Confirm:$false -ErrorAction SilentlyContinue
    if (Test-Path $Pasta) { Remove-Item -Path $Pasta -Recurse -Force }
    Write-Host 'Envio automático removido deste computador.' -ForegroundColor Green
}

# ----- Execução ---------------------------------------------------------------
try {
    if ($Instalar)        { Instalar }
    elseif ($Desinstalar) { Desinstalar }
    elseif ($Agora)       { Enviar-Planilha -Forcar }
    else                  { Enviar-Planilha }
} catch {
    $msg = $_.Exception.Message
    $cod = Codigo-Http $_
    if ($cod -eq 401) { $msg = 'O GitHub recusou o token (vencido ou revogado). Gere um novo e rode o instalar.bat.' }
    elseif ($cod -eq 403 -or $cod -eq 404) { $msg = "O token não tem permissão de escrita no repositório $Repo." }
    Escrever-Log "ERRO: $msg" -SoNoArquivo
    if (-not ($Instalar -or $Agora)) { try { Avisar-Erro $msg } catch { } }
    if ($Instalar -or $Agora) { Write-Host "ERRO: $msg" -ForegroundColor Red; exit 1 }
}
