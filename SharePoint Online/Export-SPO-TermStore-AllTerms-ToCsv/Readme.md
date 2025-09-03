# Export-SPO-Taxonomy-ToCsv.ps1

[![PowerShell](https://img.shields.io/badge/PowerShell-5.1+-blue.svg)](#предварительные-условия)
[![SharePoint Online](https://img.shields.io/badge/SharePoint-Online-0366d6)](#как-это-работает)
[![Read-only](https://img.shields.io/badge/Mode-read--only-success)](#безопасность)
[![CSV](https://img.shields.io/badge/Output-CSV%20(;)-orange)](#что-получаем-на-выходе)

Скрипт экспортирует **всю таксономию SharePoint Online** (Term Store → Term Groups → Term Sets → Terms) в **CSV** с разделителем `;` и кодировкой **UTF-8**. Поддерживает **MFA** через веб-авторизацию и работает **только на чтение**.

---

## Оглавление
- [Что делает скрипт](#что-делает-скрипт)
- [Предварительные условия](#предварительные-условия)
- [Установка модуля (однократно)](#установка-модуля-однократно)
- [Быстрый старт](#быстрый-старт)
- [Параметры](#параметры)
- [Что получаем на выходе](#что-получаем-на-выходе)
- [Почему так (CSOM vs PnP)](#почему-так-csom-vs-pnp)
- [Безопасность](#безопасность)


---

## Что делает скрипт

- Подключается к корневому сайту **SharePoint Online** `https://<tenant>.sharepoint.com` через `-UseWebLogin` (удобно с MFA).
- Через **CSOM** (ClientContext) открывает **TaxonomySession**, получает первый **Term Store**.
- Обходит все **Term Groups** → все **Term Sets**.
- Для каждого **Term Set** вытягивает **все термины** одним вызовом `GetAllTerms()` (включая вложенные).
- Для каждого термина сохраняет:  
  `GroupName, GroupId, TermSetName, TermSetId, TermName, TermGuid, TermPathWithinSet (A/B/C), TermLevel, IsAvailableForTagging, Description`.
- Экспортирует в **CSV** (разделитель `;`, UTF-8).

> Скрипт ничего в вашем тенанте **не изменяет**.

---

## Предварительные условия

- **Windows PowerShell ISE 5.1** (64-бит).
- Модуль **SharePointPnPPowerShellOnline** (legacy для PS 5.1).
- Учётная запись с правом чтения таксономии (**Term Store Administrator** или эквивалентный доступ).

Проверка версии ISE (должно быть 64-бит):
```powershell
[Environment]::Is64BitProcess  # True

Установка модуля (однократно)
# Рекомендуемая подготовка TLS/репозитория
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
if (-not (Get-PackageProvider -Name NuGet -ListAvailable -ErrorAction SilentlyContinue)) {
  Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force
}
try { Set-PSRepository -Name PSGallery -InstallationPolicy Trusted } catch {}

# Установка legacy-модуля для ISE 5.1
Install-Module SharePointPnPPowerShellOnline -Scope CurrentUser -Force -AllowClobber

Быстрый старт

Открой Windows PowerShell ISE 5.1 (64-бит).

Скопируй содержимое Export-SPO-Taxonomy-ToCsv.ps1 в окно сценария.

Укажи параметры вверху скрипта:

$Tenant = "nphgroup"               # -> https://nphgroup.sharepoint.com
$OutCsv = "C:\Temp\AllTerms.csv"   # путь к CSV


Выдели весь скрипт и нажми F8.

Во всплывающем окне войди под нужной учёткой (MFA поддерживается).

На выходе получишь AllTerms.csv.

Параметры
Имя	Тип	Обязат.	Описание
Tenant	string	да	Короткое имя тенанта (contoso → https://contoso.sharepoint.com).
OutCsv	string	да	Полный путь к результирующему CSV (разделитель ;, кодировка UTF-8).
Что получаем на выходе

CSV c полями:

GroupName — название группы Term Store

GroupId — GUID группы

TermSetName — название Term Set

TermSetId — GUID Term Set

TermName — имя термина

TermGuid — GUID термина

TermPathWithinSet — путь внутри набора в виде A/B/C

TermLevel — уровень вложенности (0 — корень)

IsAvailableForTagging — доступен ли термин для тегирования

Description — описание (если задано)

Пример (фрагмент):

GroupName;GroupId;TermSetName;TermSetId;TermName;TermGuid;TermPathWithinSet;TermLevel;IsAvailableForTagging;Description
Corporate Taxonomy;6b1...;Departments;8c2...;Finance;3f4...;Finance;0;True;Finance team
Corporate Taxonomy;6b1...;Departments;8c2...;Accounts Payable;7d9...;Finance/Accounts Payable;1;True;

Почему так (CSOM vs PnP)

В legacy-модуле SharePointPnPPowerShellOnline некоторые удобные PnP-команды/параметры недоступны или ведут себя иначе.
Поэтому обход реализован через CSOM (ClientContext + TaxonomySession + TermSet.GetAllTerms()), что:

стабильно работает в ISE 5.1;

получает всю иерархию без сложной рекурсии;

исключает ошибки Get-PnPProperty на коллекциях.

Подсказки и частые проблемы

Окно входа не появляется (IE-компоненты отключены): временно используйте учётку без MFA:

$cred = Get-Credential
Connect-PnPOnline -Url "https://<tenant>.sharepoint.com" -Credentials $cred


Модуль не найден — установите SharePointPnPPowerShellOnline (см. раздел выше).

CSV «ломается» в Excel — импортируйте как UTF-8 и с разделителем ;.

Нет доступа к Term Store — проверьте роли (напр., Term Store Administrator).

Безопасность

Скрипт выполняет только чтение (read-only) и не изменяет данные в SharePoint Online.


если хочешь — могу сгенерировать этот README.md файлом (и, например, приложить версию с именем `README.md`), скажи путь — сохраню и дам ссылку.
::contentReference[oaicite:0]{index=0}
