# ============================================================
#  PersonnelManagementApp - اصلاح خودکار موتور پایگاه داده Access
#  (Microsoft.ACE.OLEDB.12.0 / .16.0)
#
#  نحوه اجرا:
#    در PowerShell (به صورت Administrator اجرا کنید):
#      Set-ExecutionPolicy -Scope Process Bypass
#      .\setup-database-provider.ps1
# ============================================================

$ErrorActionPreference = 'Stop'

function Write-Step([string]$Message) {
    Write-Host "`n==> $Message" -ForegroundColor Cyan
}

function Test-ProviderRegistered([string]$Provider, [string]$View) {
    $root = "SOFTWARE\Classes\$Provider\CLSID"
    try {
        if ($View -eq '64') {
            $key = [Microsoft.Win32.RegistryKey]::OpenBaseKey(
                [Microsoft.Win32.RegistryHive]::LocalMachine,
                [Microsoft.Win32.RegistryView]::Registry64)
        } else {
            $key = [Microsoft.Win32.RegistryKey]::OpenBaseKey(
                [Microsoft.Win32.RegistryHive]::LocalMachine,
                [Microsoft.Win32.RegistryView]::Registry32)
        }
        $sub = $key.OpenSubKey($root)
        $result = ($sub -ne $null)
        if ($sub) { $sub.Dispose() }
        $key.Dispose()
        return $result
    } catch {
        return $false
    }
}

Write-Step "بررسی معماری ویندوز و حالت PowerShell..."
$is64BitPS = [Environment]::Is64BitProcess
$is64BitOS = [Environment]::Is64BitOperatingSystem
Write-Host "  PowerShell 64 بیتی: $is64BitPS  |  سیستم عامل 64 بیتی: $is64BitOS"

if (-not $is64BitOS) {
    Write-Host "  این سیستم 32 بیتی است. برنامه نیز در حالت 32 بیتی اجرا می‌شود." -ForegroundColor Yellow
    Write-Host "  تنها نسخه x86 AccessDatabaseEngine نصب شود." -ForegroundColor Yellow
}

# --- بررسی وضعیت فعلی ---
Write-Step "بررسی ثبت بودن provider های Access..."
$reg32_12 = Test-ProviderRegistered 'Microsoft.ACE.OLEDB.12.0' '32'
$reg32_16 = Test-ProviderRegistered 'Microsoft.ACE.OLEDB.16.0' '32'
$reg64_16 = Test-ProviderRegistered 'Microsoft.ACE.OLEDB.16.0' '64'

Write-Host "  ACE12 (32-bit): $reg32_12   ACE16 (32-bit): $reg32_16   ACE16 (64-bit): $reg64_16"

# --- انتخاب نصب‌کننده مناسب ---
if ($is64BitOS -and $is64BitPS) {
    $wanted = 'x64'
} else {
    $wanted = 'x86'
}
Write-Host "  نصب‌کننده مناسب: $wanted" -ForegroundColor Green

# --- راه ۱: winget (اگر در دسترس باشد) ---
$winget = Get-Command winget -ErrorAction SilentlyContinue
if ($winget) {
    Write-Step "تلاش برای نصب با winget (Microsoft.AccessDatabaseEngine.2016)..."
    & winget install --id Microsoft.AccessDatabaseEngine.2016 --exact --silent `
        --accept-package-agreements --accept-source-agreements --disable-interactivity
    if ($LASTEXITCODE -eq 0) {
        Write-Host "  نصب با winget موفق بود." -ForegroundColor Green
        Write-Host "  لطفاً برنامه را ببندید و دوباره اجرا کنید." -ForegroundColor Yellow
        Read-Host "برای خروج Enter را فشار دهید"
        exit 0
    }
    Write-Host "  نصب با winget ناموفق بود؛ تلاش با نصب‌کننده رسمی..." -ForegroundColor Yellow
} else {
    Write-Host "  winget موجود نیست؛ از نصب‌کننده رسمی استفاده می‌شود." -ForegroundColor Yellow
}

# --- راه ۲: دانلود و اجرای نصب‌کننده رسمی Microsoft ---
$url = "https://download.microsoft.com/download/3/5/c/35c84c36-661a-44e6-9324-8786b8dbe231/accessdatabaseengine_$wanted.exe"
$installer = Join-Path $env:TEMP "accessdatabaseengine_$wanted.exe"

Write-Step "دانلود نصب‌کننده رسمی Access Database Engine ($wanted)..."
try {
    Invoke-WebRequest -Uri $url -OutFile $installer -UseBasicParsing
} catch {
    Write-Host "  دانلود ناموفق: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "  لطفاً به‌صورت دستی از صفحه رسمی دانلود و نصب کنید:" -ForegroundColor Yellow
    Write-Host "    https://www.microsoft.com/en-us/download/details.aspx?id=54920"
    exit 1
}

Write-Step "اجرای نصب‌کننده (نیاز به تأیید UAC دارد)..."
try {
    $proc = Start-Process -FilePath $installer -ArgumentList '/quiet' -Verb RunAs -PassThru
    $proc.WaitForExit()
    if ($proc.ExitCode -eq 0 -or $proc.ExitCode -eq 3010) {
        Write-Host "  نصب با موفقیت انجام شد (کد $($proc.ExitCode))." -ForegroundColor Green
    } else {
        Write-Host "  نصب با کد خطای $($proc.ExitCode) پایان یافت." -ForegroundColor Red
    }
} catch {
    Write-Host "  اجرای نصب‌کننده ناموفق: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "  اگر UAC را لغو کردید، لطفاً به‌صورت دستی نصب را انجام دهید." -ForegroundColor Yellow
}

Write-Step "پایان. تأیید نهایی:"
$reg64_16b = Test-ProviderRegistered 'Microsoft.ACE.OLEDB.16.0' '64'
$reg32_16b = Test-ProviderRegistered 'Microsoft.ACE.OLEDB.16.0' '32'
Write-Host "  ACE16 (64-bit): $reg64_16b   ACE16 (32-bit): $reg32_16b"

Write-Host "`nبرای اعمال کامل تغییرات، برنامه را ببندید و دوباره اجرا کنید." -ForegroundColor Yellow
Read-Host "برای خروج Enter را فشار دهید"
