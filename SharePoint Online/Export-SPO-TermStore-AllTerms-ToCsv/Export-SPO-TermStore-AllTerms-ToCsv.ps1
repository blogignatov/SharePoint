
# ===================== НАСТРОЙКИ =====================
$Tenant = "nphgroup"                               # Задаём короткое имя тенанта: будет использовано в URL https://<tenant>.sharepoint.com
$OutCsv = "C:\Temp\AllTerms01.csv"                 # Полный путь, куда сохраним CSV (разделитель — точка с запятой)
# =====================================================

# TLS для PS 5.1
[Net.ServicePointManager]::SecurityProtocol =       # Включаем поддерживаемые версии TLS для сетевых вызовов из PowerShell 5.1
  [Net.SecurityProtocolType]::Tls12 -bor            # Добавляем TLS 1.2 (обязательно для SPO)
  [Net.SecurityProtocolType]::Tls11 -bor            # Добавляем TLS 1.1 (на случай обратной совместимости)
  [Net.SecurityProtocolType]::Tls                   # Добавляем “обычный” TLS (включает 1.0)

# Импорт legacy-модуля (ISE 5.1)
if (-not (Get-Module -ListAvailable -Name SharePointPnPPowerShellOnline)) {      # Проверяем, установлен ли модуль SharePointPnPPowerShellOnline
  Write-Error "Нет модуля 'SharePointPnPPowerShellOnline'. Установите:
  Install-Module SharePointPnPPowerShellOnline -Scope CurrentUser -Force -AllowClobber" # Сообщаем, как его установить
  return                                                                           # Прерываем выполнение, если модуля нет
}
Import-Module SharePointPnPPowerShellOnline -DisableNameChecking -ErrorAction Stop # Загружаем модуль в текущую сессию (отключаем предупреждения об именах)

# --- Подключаемся к КОРНЕВОМУ сайту (не к -admin) ---
$siteUrl = "https://$Tenant.sharepoint.com"          # Формируем URL корневого сайта для подключения
try {                                                # Пытаемся подключиться
  Connect-PnPOnline -Url $siteUrl -UseWebLogin -WarningAction Ignore -ErrorAction Stop # Открываем веб-окно входа (подходит для MFA), игнорируем варнинги
  Write-Host "Подключено к $siteUrl" -ForegroundColor Green                            # Пишем, что подключение успешно
} catch {
  Write-Error "Не удалось подключиться: $($_.Exception.Message)"; return               # При ошибке — сообщение и выход
}

# Получаем ClientContext
$ctx = Get-PnPContext                                # Берём CSOM-контекст текущего подключения для прямой работы с объектной моделью

# Помощник: безопасная загрузка объектов CSOM
function Load-CSOM {                                  # Объявляем вспомогательную функцию Load-CSOM
  param([Parameter(Mandatory=$true)]$Object)          # Принимает один обязательный параметр — объект CSOM для загрузки
  $ctx.Load($Object)                                  # Говорим контексту, какие свойства объекта надо вытащить с сервера
  Invoke-PnPQuery                                     # Выполняем запрос к серверу (SharePoint) и реально загружаем данные
}

# Коллекция результата
$Rows = New-Object System.Collections.Generic.List[object]  # Создаём список в памяти, куда будем складывать строки для CSV

# --- Берём TaxonomySession и Term Store ---
try {
  $taxSession = [Microsoft.SharePoint.Client.Taxonomy.TaxonomySession]::GetTaxonomySession($ctx) # Создаём таксономическую сессию на базе ClientContext
  Load-CSOM $taxSession                                   # Загружаем объект сессии с сервера

  $termStores = $taxSession.TermStores                    # Берём коллекцию Term Store’ов (обычно в SPO один)
  Load-CSOM $termStores                                   # Загружаем коллекцию со свойствами (Count и т.д.)
  if ($termStores.Count -eq 0) {                          # Если коллекция пустая —
    throw "Term Store не найден (нет доступа или служба таксономии отключена)."  # — бросаем исключение с понятным текстом
  }

  $termStore = $termStores[0]                             # Берём первый доступный Term Store
  Load-CSOM $termStore.Groups                             # Загружаем коллекцию групп Term Store (Term Groups)
} catch {
  Write-Error "Не удалось получить Term Store: $($_.Exception.Message)" # В случае ошибки выводим причину
  Disconnect-PnPOnline | Out-Null                         # Закрываем PnP-сессию аккуратно
  return                                                  # И выходим
}

# --- Обход групп → наборов → все термины набора (без рекурсии, через GetAllTerms) ---
foreach ($group in $termStore.Groups) {                   # Идём по каждой группе терминов в Term Store
  # Группа
  $ctx.Load($group) | Out-Null                            # Просим сервер отдать свойства текущей группы
  $ctx.Load($group.TermSets) | Out-Null                   # Просим сервер отдать коллекцию наборов терминов этой группы
  Invoke-PnPQuery                                         # Выполняем запрос (фактическая загрузка с сервера)

  foreach ($ts in $group.TermSets) {                      # Идём по каждому набору терминов (Term Set) в группе
    # Набор терминов
    $ctx.Load($ts) | Out-Null                             # Загружаем базовые свойства набора
    Invoke-PnPQuery                                       # Выполняем запрос к серверу

    # Все термины набора одним запросом
    $allTerms = $ts.GetAllTerms()                         # Получаем «плоскую» коллекцию всех терминов этого набора (включая вложенные)
    $ctx.Load($allTerms) | Out-Null                       # Говорим контексту загрузить эту коллекцию
    Invoke-PnPQuery                                       # Загружаем с сервера

    foreach ($term in $allTerms) {                        # Идём по каждому термину набора
      # На всякий — подгрузим базовые свойства термина (без тяжёлых Includes)
      $ctx.Load($term) | Out-Null                         # Запрашиваем свойства термина (Name, Id, PathOfTerm и пр.)
      Invoke-PnPQuery                                     # Загружаем их

      # Путь в наборе: PathOfTerm обычно "A;B;C" — нормализуем в "A/B/C"
      $path = $term.PathOfTerm                            # Берём строку пути термина внутри набора (уровни через ';' или '|')
      if ($null -ne $path) {                              # Если путь присутствует —
        $path = $path -replace '\|','/' -replace ';','/'  # — заменяем разделители на «/» для удобства чтения (A/B/C)
      } else {
        $path = $term.Name                                # Если PathOfTerm нет, используем просто имя термина
      }
      $level = ($path -split '/').Count - 1               # Уровень термина: считаем количество сегментов пути минус 1 (корень = 0)
      if ($level -lt 0) { $level = 0 }                    # На всякий: не даём уровню уйти в минус

      # Пытаемся прочитать необязательные свойства (если не подгружены — будет $null)
      $isTaggable = $null                                 # По умолчанию признак «доступен для тегирования» = $null
      $desc = $null                                       # По умолчанию описание = $null
      try { $isTaggable = $term.IsAvailableForTagging } catch {}  # Пробуем взять IsAvailableForTagging, игнорируя ошибку
      try { $desc = $term.Description } catch {}                 # Пробуем взять Description, игнорируя ошибку

      # Добавляем строку
      $Rows.Add([pscustomobject]@{                        # Создаём объект-строку с нужными полями…
        GroupName             = $group.Name               # …имя группы Term Store
        GroupId               = $group.Id.ToString()      # …GUID группы Term Store
        TermSetName           = $ts.Name                  # …имя набора терминов
        TermSetId             = $ts.Id.ToString()         # …GUID набора терминов
        TermName              = $term.Name                # …имя термина
        TermGuid              = $term.Id.ToString()       # …GUID термина
        TermPathWithinSet     = $path                     # …путь термина внутри набора (A/B/C)
        TermLevel             = $level                    # …уровень вложенности (0 — корневой термин)
        IsAvailableForTagging = $isTaggable               # …доступен ли термин для тегирования
        Description           = $desc                     # …описание термина (если есть)
      }) | Out-Null                                       # Добавляем строку в коллекцию $Rows и подавляем вывод в консоль
    }
  }
}

# --- Сохраняем CSV с разделителем ';' ---
$dir = Split-Path -Path $OutCsv -Parent                 # Вычисляем директорию, где должен лежать CSV
if ($dir -and -not (Test-Path $dir)) {                  # Если папки ещё нет —
  New-Item -Path $dir -ItemType Directory -Force | Out-Null  # — создаём её
}

try {
  $Rows | Export-Csv -Path $OutCsv -NoTypeInformation -Encoding UTF8 -Delimiter ';' # Пишем CSV: UTF-8, без строки типов, с разделителем «;»
  Write-Host "Готово → $OutCsv (строк: $($Rows.Count))" -ForegroundColor Green       # Сообщаем об успехе и количестве выгруженных строк
} catch {
  Write-Error "Не удалось записать CSV: $($_.Exception.Message)"                    # Если не удалось сохранить файл — выводим ошибку
}

Disconnect-PnPOnline | Out-Null                        # Закрываем PnP-подключение (аккуратно завершаем сессию)
