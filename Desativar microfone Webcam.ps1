# ==========================
# Definir o microfone Realtek como padrão
# ==========================

# Verifica se está sendo executado como administrador
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Start-Process powershell.exe "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs
    exit
}

# Nome do microfone Realtek (ajuste conforme seu sistema)
$MicrophoneName = "Realtek"

# Caminhos possíveis para o NirCMD
$NirCmdPaths = @(
    "C:\Windows\System32\nircmd.exe",
    "C:\Tools\nircmd.exe",
    "C:\Users\$env:USERNAME\Downloads\nircmd-x64\nircmd.exe"
)

$NirCmdPath = $NirCmdPaths | Where-Object { Test-Path $_ } | Select-Object -First 1

if (-not $NirCmdPath) {
    Write-Output "Erro: NirCMD não encontrado. Verifique a instalação!"
    exit 1
}

# Verifica se o microfone Realtek está presente (classe corrigida para AudioEndpoint)
$RealtekMic = Get-PnpDevice -Class AudioEndpoint -PresentOnly | 
              Where-Object { $_.FriendlyName -match $MicrophoneName }

if (-not $RealtekMic) {
    Write-Output "Erro: Microfone Realtek não encontrado! Dispositivos disponíveis:"
    Get-PnpDevice -Class AudioEndpoint -PresentOnly | Format-Table FriendlyName, InstanceId
    exit 1
}

Write-Output "Definindo '$($RealtekMic.FriendlyName)' como microfone padrão..."
& $NirCmdPath setdefaultsounddevice "$($RealtekMic.FriendlyName)" 1
& $NirCmdPath setdefaultsounddevice "$($RealtekMic.FriendlyName)" 2
Write-Output "Microfone padrão alterado."

# ==========================
# Desativar e remover o microfone da GENERAL WEBCAM
# ==========================

$WebcamMic = Get-PnpDevice -Class AudioEndpoint -PresentOnly | 
             Where-Object { $_.FriendlyName -match "GENERAL WEBCAM" }

if ($WebcamMic) {
    try {
        Write-Output "Desativando o microfone da GENERAL WEBCAM..."
        Disable-PnpDevice -InstanceId $WebcamMic.InstanceId -Confirm:$false -ErrorAction Stop
        
        Write-Output "Removendo o driver..."
        pnputil /remove-device $WebcamMic.InstanceId
        
        Write-Output "Dispositivo desativado e removido."
    } catch {
        Write-Output "Erro ao processar dispositivo: $_"
        exit 1
    }
} else {
    Write-Output "Microfone da GENERAL WEBCAM não encontrado."
}