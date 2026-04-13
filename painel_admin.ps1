param(
    [string]$BaseUrl = "",
    [string]$AdminToken = ""
)

$ErrorActionPreference = "Stop"
$ConfigPath = Join-Path $PSScriptRoot "admin_panel_config.json"
$ExportDir = Join-Path $PSScriptRoot "admin_exports"
$script:LastHealthText = "Nao testado"
$script:LastHealthOk = $false

function Write-Title {
    Clear-Host
    Write-Host ""
    Write-Host "============================================================" -ForegroundColor Cyan
    Write-Host "  Guia Patrimonial - Painel Admin (Producao)" -ForegroundColor Cyan
    Write-Host "============================================================" -ForegroundColor Cyan
    Write-Host ""
}

function Load-Config {
    if (-not (Test-Path -LiteralPath $ConfigPath)) {
        return @{}
    }
    try {
        $json = Get-Content -LiteralPath $ConfigPath -Raw
        if ([string]::IsNullOrWhiteSpace($json)) { return @{} }
        return ($json | ConvertFrom-Json -AsHashtable)
    }
    catch {
        return @{}
    }
}

function Save-Config([string]$Url, [string]$Token) {
    $payload = @{
        base_url    = $Url
        admin_token = $Token
    } | ConvertTo-Json
    Set-Content -LiteralPath $ConfigPath -Value $payload -Encoding UTF8
}

function Read-NonEmpty([string]$Prompt, [string]$DefaultValue = "") {
    while ($true) {
        if ([string]::IsNullOrWhiteSpace($DefaultValue)) {
            $value = Read-Host $Prompt
        }
        else {
            $value = Read-Host "$Prompt [$DefaultValue]"
            if ([string]::IsNullOrWhiteSpace($value)) {
                $value = $DefaultValue
            }
        }
        if (-not [string]::IsNullOrWhiteSpace($value)) {
            return $value.Trim()
        }
        Write-Host "Valor obrigatorio." -ForegroundColor Yellow
    }
}

function Ensure-Settings {
    $cfg = Load-Config
    if ([string]::IsNullOrWhiteSpace($script:BaseUrl)) { $script:BaseUrl = [string]$cfg.base_url }
    if ([string]::IsNullOrWhiteSpace($script:AdminToken)) { $script:AdminToken = [string]$cfg.admin_token }

    if ([string]::IsNullOrWhiteSpace($script:BaseUrl)) {
        $script:BaseUrl = Read-NonEmpty "Informe a URL do servidor (ex: https://seu-app.onrender.com)"
    }
    if ([string]::IsNullOrWhiteSpace($script:AdminToken)) {
        $script:AdminToken = Read-NonEmpty "Informe o ADMIN_TOKEN"
    }

    $script:BaseUrl = $script:BaseUrl.TrimEnd("/")
    Save-Config -Url $script:BaseUrl -Token $script:AdminToken
}

function Get-MaskedToken([string]$Token) {
    if ([string]::IsNullOrWhiteSpace($Token)) { return "(vazio)" }
    if ($Token.Length -le 4) { return "*" * $Token.Length }
    return ("*" * ($Token.Length - 4)) + $Token.Substring($Token.Length - 4, 4)
}

function Confirm-Action([string]$Prompt) {
    $answer = Read-Host "$Prompt (s/N)"
    return ($answer -match "^(s|sim|y|yes)$")
}

function Invoke-AdminApi([string]$Method, [string]$Path) {
    $uri = "$($script:BaseUrl)$Path"
    $headers = @{ "X-Admin-Token" = $script:AdminToken }
    try {
        $response = Invoke-WebRequest -Method $Method -Uri $uri -Headers $headers -UseBasicParsing
    }
    catch {
        $details = $_.Exception.Message
        if ($_.Exception.Response -and $_.Exception.Response.GetResponseStream()) {
            try {
                $reader = New-Object System.IO.StreamReader($_.Exception.Response.GetResponseStream())
                $body = $reader.ReadToEnd()
                if (-not [string]::IsNullOrWhiteSpace($body)) {
                    $details = "$details | Body: $body"
                }
            } catch {}
        }
        throw $details
    }

    if ([string]::IsNullOrWhiteSpace($response.Content)) {
        return @{}
    }
    try {
        return ($response.Content | ConvertFrom-Json)
    }
    catch {
        return @{}
    }
}

function Update-HealthStatus {
    try {
        $resp = Invoke-WebRequest -Method GET -Uri "$($script:BaseUrl)/health" -UseBasicParsing
        $payload = $null
        if (-not [string]::IsNullOrWhiteSpace($resp.Content)) {
            $payload = $resp.Content | ConvertFrom-Json
        }
        $storage = if ($payload -and $payload.storage) { $payload.storage } else { "desconhecido" }
        $script:LastHealthText = "Online (storage: $storage)"
        $script:LastHealthOk = $true
    }
    catch {
        $script:LastHealthText = "Offline/Erro"
        $script:LastHealthOk = $false
    }
}

function Get-Devices {
    try {
        $resp = Invoke-AdminApi -Method "GET" -Path "/admin/devices"
        if ($null -eq $resp -or $null -eq $resp.devices) {
            return @()
        }
        return @($resp.devices)
    }
    catch {
        Write-Host ""
        Write-Host "Falha ao consultar dispositivos." -ForegroundColor Red
        Write-Host $_ -ForegroundColor Red
        Write-Host ""
        return @()
    }
}

function ConvertTo-DeviceTable($devices) {
    return @(
        for ($i = 0; $i -lt $devices.Count; $i++) {
            $d = $devices[$i]
            [PSCustomObject]@{
                idx          = $i + 1
                device_id    = $d.device_id
                machine_name = $d.machine_name
                user_name    = $d.user_name
                blocked      = [bool]$d.blocked
                last_seen    = $d.last_seen
            }
        }
    )
}

function Show-Devices {
    $devices = Get-Devices
    if ($devices.Count -eq 0) {
        Write-Host "Total de dispositivos: 0" -ForegroundColor Cyan
        Write-Host "Nenhum dispositivo registrado." -ForegroundColor Yellow
        return @()
    }

    $table = ConvertTo-DeviceTable -devices $devices
    $blockedCount = @($table | Where-Object { $_.blocked }).Count
    $activeCount = $table.Count - $blockedCount

    Write-Host ("Total de dispositivos: " + $table.Count) -ForegroundColor Cyan
    Write-Host ("Ativos: $activeCount | Bloqueados: $blockedCount") -ForegroundColor DarkCyan
    $table | Sort-Object -Property blocked, last_seen -Descending | Format-Table -AutoSize | Out-Host
    return $table
}

function Resolve-DeviceId($devicesTable) {
    $inputValue = Read-Host "Digite o idx da tabela ou o device_id"
    if ([string]::IsNullOrWhiteSpace($inputValue)) { return "" }

    $inputValue = $inputValue.Trim()
    if ($inputValue -match "^\d+$") {
        $idx = [int]$inputValue
        $row = $devicesTable | Where-Object { $_.idx -eq $idx } | Select-Object -First 1
        if ($null -ne $row) { return [string]$row.device_id }
        Write-Host "idx invalido." -ForegroundColor Yellow
        return ""
    }
    return $inputValue
}

function Block-Device {
    $table = Show-Devices
    if ($table.Count -eq 0) { return }

    $deviceId = Resolve-DeviceId -devicesTable $table
    if ([string]::IsNullOrWhiteSpace($deviceId)) { return }
    if (-not (Confirm-Action "Confirmar bloqueio de '$deviceId'?")) { return }

    try {
        $resp = Invoke-AdminApi -Method "POST" -Path "/admin/block/$deviceId"
        Write-Host ""
        Write-Host "Bloqueado com sucesso: $($resp.device_id)" -ForegroundColor Green
        Write-Host ""
    }
    catch {
        Write-Host ""
        Write-Host "Falha ao bloquear dispositivo." -ForegroundColor Red
        Write-Host $_ -ForegroundColor Red
        Write-Host ""
    }
}

function Unblock-Device {
    $table = Show-Devices
    if ($table.Count -eq 0) { return }

    $deviceId = Resolve-DeviceId -devicesTable $table
    if ([string]::IsNullOrWhiteSpace($deviceId)) { return }
    if (-not (Confirm-Action "Confirmar desbloqueio de '$deviceId'?")) { return }

    try {
        $resp = Invoke-AdminApi -Method "POST" -Path "/admin/unblock/$deviceId"
        Write-Host ""
        Write-Host "Desbloqueado com sucesso: $($resp.device_id)" -ForegroundColor Green
        Write-Host ""
    }
    catch {
        Write-Host ""
        Write-Host "Falha ao desbloquear dispositivo." -ForegroundColor Red
        Write-Host $_ -ForegroundColor Red
        Write-Host ""
    }
}

function Export-DevicesCsv {
    $devices = Get-Devices
    if ($devices.Count -eq 0) {
        Write-Host ""
        Write-Host "Sem dispositivos para exportar." -ForegroundColor Yellow
        Write-Host ""
        return
    }

    if (-not (Test-Path -LiteralPath $ExportDir)) {
        New-Item -ItemType Directory -Path $ExportDir | Out-Null
    }

    $table = ConvertTo-DeviceTable -devices $devices
    $fileName = "devices-{0}.csv" -f (Get-Date -Format "yyyyMMdd-HHmmss")
    $target = Join-Path $ExportDir $fileName
    $table | Export-Csv -LiteralPath $target -NoTypeInformation -Encoding UTF8
    Write-Host ""
    Write-Host "Exportado em: $target" -ForegroundColor Green
    Write-Host ""
}

function Edit-Settings {
    $script:BaseUrl = Read-NonEmpty "URL do servidor" $script:BaseUrl
    $script:AdminToken = Read-NonEmpty "ADMIN_TOKEN" $script:AdminToken
    $script:BaseUrl = $script:BaseUrl.TrimEnd("/")
    Save-Config -Url $script:BaseUrl -Token $script:AdminToken
    Update-HealthStatus
    Write-Host ""
    Write-Host "Configuracoes salvas em $ConfigPath" -ForegroundColor Green
    Write-Host ""
}

Ensure-Settings
Update-HealthStatus

while ($true) {
    Write-Title
    Write-Host "Servidor: $script:BaseUrl" -ForegroundColor DarkGray
    Write-Host "Token:    $(Get-MaskedToken -Token $script:AdminToken)" -ForegroundColor DarkGray
    if ($script:LastHealthOk) {
        Write-Host "Status:   $script:LastHealthText" -ForegroundColor Green
    } else {
        Write-Host "Status:   $script:LastHealthText" -ForegroundColor Red
    }
    Write-Host ""
    Write-Host "1) Listar dispositivos"
    Write-Host "2) Bloquear dispositivo"
    Write-Host "3) Desbloquear dispositivo"
    Write-Host "4) Alterar URL / TOKEN"
    Write-Host "5) Testar conexao"
    Write-Host "6) Exportar dispositivos (CSV)"
    Write-Host "0) Sair"
    Write-Host ""

    $option = Read-Host "Escolha uma opcao"

    switch ($option) {
        "1" { Write-Host ""; Show-Devices; Write-Host ""; Pause }
        "2" { Write-Host ""; Block-Device; Pause }
        "3" { Write-Host ""; Unblock-Device; Pause }
        "4" { Write-Host ""; Edit-Settings; Pause }
        "5" {
            Update-HealthStatus
            Write-Host ""
            if ($script:LastHealthOk) {
                Write-Host "Conexao validada: $script:LastHealthText" -ForegroundColor Green
            } else {
                Write-Host "Falha de conexao com o servidor." -ForegroundColor Red
            }
            Write-Host ""
            Pause
        }
        "6" { Write-Host ""; Export-DevicesCsv; Pause }
        "0" { break }
        default {
            Write-Host ""
            Write-Host "Opcao invalida." -ForegroundColor Yellow
            Start-Sleep -Milliseconds 900
        }
    }
}
