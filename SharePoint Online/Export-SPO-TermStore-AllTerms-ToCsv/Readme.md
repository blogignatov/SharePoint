Кратко о скрипте:

Подключается к корневому сайту SharePoint Online (https://<tenant>.sharepoint.com) через окно входа (UseWebLogin, удобно с MFA).

Через CSOM (ClientContext) открывает TaxonomySession, берёт первый Term Store, затем проходит все Term Groups → все Term Sets.

Для каждого Term Set берёт все термины одним вызовом GetAllTerms() (включая вложенные).

Формирует для каждого термина строку с полями:
GroupName, GroupId, TermSetName, TermSetId, TermName, TermGuid, TermPathWithinSet (в виде A/B/C), TermLevel (0–n), IsAvailableForTagging, Description.

Сохраняет результат в CSV по указанному пути, с разделителем ; и кодировкой UTF-8.

Работает только на чтение — ничего в SPO не изменяет.

Предназначен для Windows PowerShell ISE 5.1 с установленным модулем SharePointPnPPowerShellOnline.
Запуск: вставить, указать $Tenant и $OutCsv, выделить всё и нажать F8.