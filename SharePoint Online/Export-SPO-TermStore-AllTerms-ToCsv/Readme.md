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
- [Подсказки и частые проблемы](#подсказки-и-частые-проблемы)
- [Безопасность](#безопасность)
- [Лицензия](#лицензия)

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
