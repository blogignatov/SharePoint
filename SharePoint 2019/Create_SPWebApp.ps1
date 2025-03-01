Add-PSSnapin Microsoft.SharePoint.PowerShell -ErrorAction SilentlyContinue

# Переменные
$AppPoolAccount = "domen\ServiceAccountPool"  # Учётная запись пула приложений
$ApplicationPoolName = "SharePoint-NamaAPP.domen.ru" # Имя пула приложений
$Name = "SharePoint-NamaAPP.domen.ru"  # Описание веб-приложения
$HostHeader = "NamaAPP.domen.ru"   # Заголовок хоста
$WebAppUrl = "https://NamaAPP.domen.ru" # URL веб-приложения
$IISPath = "C:\inetpub\wwwroot\wss\VirtualDirectories\NamaAPP"  # Путь к директории IIS
$ContentDatabase = "WSS_content_NamaAPP" # Имя базы данных контента
$DatabaseServer = "SQLServer01"  # Сервер БД

# ✅ Создаём провайдер аутентификации с Kerberos
$AuthProvider = New-SPAuthenticationProvider -UseWindowsIntegratedAuthentication
$AuthProvider.DisableKerberos = $false  # Включаем Kerberos

# ✅ Создание веб-приложения с HTTPS и Claims-Based Authentication
New-SPWebApplication -ApplicationPool $ApplicationPoolName `
                     -ApplicationPoolAccount (Get-SPManagedAccount $AppPoolAccount) `
                     -Name $Name `
                     -DatabaseName $ContentDatabase `
                     -DatabaseServer $DatabaseServer `
                     -HostHeader $HostHeader `
                     -Path $IISPath `
                     -Port 443 `
                     -URL $WebAppUrl `
                     -SecureSocketsLayer `
                     -AuthenticationProvider $AuthProvider
