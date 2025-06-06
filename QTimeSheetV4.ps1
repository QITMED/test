﻿# --- FUNCION PARA DETECTAR .NET DESKTOP RUNTIME 8.0 DESDE EL REGISTRO ---
function Test-DotNetDesktopRuntimeInstalled {
    $desktopRuntimeKey = 'HKLM:\SOFTWARE\WOW6432Node\dotnet\Setup\InstalledVersions\x64\sharedfx\Microsoft.WindowsDesktop.App'
    if (Test-Path $desktopRuntimeKey) {
        $installedVersions = (Get-ItemProperty -Path $desktopRuntimeKey).PSObject.Properties.Name
        return $installedVersions -match '^8\.0\.'
    }
    return $false
}

# --- FUNCION PARA DESCARGA FIABLE DESDE URL ---
function Download-File {
    param (
        [string]$url,
        [string]$outputPath
    )
    try {
        $wc = New-Object System.Net.WebClient
        $wc.Headers.Add("User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64)")
        $wc.DownloadFile($url, $outputPath)
        return $true
    } catch {
        Write-Output "ERROR al descargar $url - $_"
        return $false
    }
}

# --- VARIABLES ---
$runtimeUrl = "https://epmhyperv2azure.blob.core.windows.net/sistemas/windowsdesktop-runtime-8.0.16-win-x64.exe"
$zipUrl = "https://epmhyperv2azure.blob.core.windows.net/sistemas/QTimeSheet.zip"
$localInstaller = "$env:TEMP\windowsdesktop-runtime-8.0.16-win-x64.exe"
$localZip = "$env:TEMP\QTimeSheet.zip"
$destinationPath = "C:\QTimeSheet"
$netInstalled = $false

# --- DETECTAR USUARIO INTERACTIVO REAL ---
try {
    $currentUser = (Get-WmiObject -Class Win32_ComputerSystem | Select-Object -ExpandProperty UserName)
    if (-not $currentUser) {
        throw "No se pudo detectar el usuario interactivo actual."
    }
    Write-Output "Usuario interactivo detectado: $currentUser"
} catch {
    Write-Output "ERROR al detectar el usuario interactivo: $_"
    exit 1
}

# --- DESCARGA E INSTALACION DE .NET RUNTIME ---
if (Test-DotNetDesktopRuntimeInstalled) {
    Write-Output ".NET Desktop Runtime 8.0 ya está instalado. Continuando..."
    $netInstalled = $true
} elseif (Download-File -url $runtimeUrl -outputPath $localInstaller) {
    try {
        Write-Output "Instalando .NET Desktop Runtime..."
        Start-Process -FilePath $localInstaller -ArgumentList "/install /quiet /norestart" -Wait
        Remove-Item $localInstaller -Force
        Start-Sleep -Seconds 10
        if (Test-DotNetDesktopRuntimeInstalled) {
            Write-Output ".NET Desktop Runtime instalado correctamente."
            $netInstalled = $true
        } else {
            Write-Output "ERROR: Instalador ejecutado pero no se detecta el runtime."
        }
    } catch {
        Write-Output "ERROR al ejecutar el instalador: $_"
    }
} else {
    Write-Output "No se pudo descargar el instalador de .NET Desktop Runtime."
}

# --- DESCARGA Y EXTRACCION DE QTimeSheet ---
if (Download-File -url $zipUrl -outputPath $localZip) {
    try {
        if (Test-Path $destinationPath) {
            Remove-Item -Path $destinationPath -Recurse -Force
        }
        Expand-Archive -LiteralPath $localZip -DestinationPath "C:\" -Force
        Remove-Item $localZip -Force
        Write-Output "QTimeSheet.zip extraído correctamente."
    } catch {
        Write-Output "ERROR al descomprimir QTimeSheet.zip: $_"
    }
} else {
    Write-Output "No se pudo descargar QTimeSheet.zip"
}

# --- CREAR TAREAS PROGRAMADAS ---
if ($netInstalled -and (Test-Path "$destinationPath\QTimeSheet.exe")) {
    $exePath = "C:\QTimeSheet\QTimeSheet.exe"
    $taskName = "QTimeSheet_Combined"
    $taskNameRepeat = "QTimeSheet_Repeat"

    try {
        Unregister-ScheduledTask -TaskName $taskName -Confirm:$false -ErrorAction SilentlyContinue
        schtasks.exe /Delete /TN "$taskNameRepeat" /F | Out-Null

        $action = New-ScheduledTaskAction -Execute $exePath
        $triggerLogonImmediate = New-ScheduledTaskTrigger -AtLogOn
        $triggerLogonDelay = New-ScheduledTaskTrigger -AtLogOn
        $triggerLogonDelay.Delay = "PT15M"
        $settings = New-ScheduledTaskSettingsSet -StartWhenAvailable

        $task = New-ScheduledTask -Action $action `
            -Trigger @($triggerLogonImmediate, $triggerLogonDelay) `
            -Settings $settings `
            -Principal (New-ScheduledTaskPrincipal -UserId $currentUser -LogonType Interactive -RunLevel Limited)

        Register-ScheduledTask -TaskName $taskName -InputObject $task -Force

        schtasks.exe /Create /TN "$taskNameRepeat" /TR "$exePath" /SC HOURLY /ST 00:00 /RU $currentUser /RL LIMITED /F | Out-Null

        Write-Output "Tareas programadas creadas correctamente para $currentUser."
    } catch {
        Write-Output "ERROR al crear tareas programadas: $_"
    }
} else {
    Write-Output "No se crearon tareas porque no están disponibles los archivos requeridos."
}
