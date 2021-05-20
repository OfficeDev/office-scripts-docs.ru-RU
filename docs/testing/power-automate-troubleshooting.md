---
title: Устранение неполадок Office, работающих в Power Automate
description: Советы, информация о платформе и известные проблемы с интеграцией между Office и Power Automate.
ms.date: 05/17/2021
localization_priority: Normal
ms.openlocfilehash: e26378051c764d97b4e8d748abc85fbe095c7b03
ms.sourcegitcommit: 4687693f02fc90a57ba30c461f35046e02e6f5fb
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 05/19/2021
ms.locfileid: "52545574"
---
# <a name="troubleshoot-office-scripts-running-in-power-automate"></a>Устранение неполадок Office, работающих в Power Automate

Power Automate позволяет выть автоматизацию Office скрипта на новый уровень. Однако, Power Automate выполняет скрипты от вашего имени в Excel сеансах, есть несколько важных вещей, чтобы отметить.

> [!TIP]
> Если вы только начинаете использовать Office с Power Automate, пожалуйста, начните [с Run Office Scripts с Power Automate,](../develop/power-automate-integration.md) чтобы узнать о платформах.

## <a name="avoid-relative-references"></a>Избегайте относительных ссылок

Power Automate выполняет свой скрипт в выбранной Excel от вашего имени. Рабочая книга может быть закрыта, когда это произойдет. Любой API, который зависит от текущего состояния пользователя, например, может вести `Workbook.getActiveWorksheet` себя по-разному в Power Automate. Это связано с тем, что API основаны на относительном положении представления или курсора пользователя и что ссылка не существует в Power Automate потоке.

Некоторые относительные ссылки API бросают ошибки в Power Automate. Другие имеют поведение по умолчанию, которое подразумевает состояние пользователя. При проектировании скриптов обязательно используйте абсолютные ссылки для листов и диапазонов. Это делает Power Automate потоком, даже если листы переставлены.

### <a name="script-methods-that-fail-when-run-power-automate-flows"></a>Методы скрипта, которые терпят неудачу при Power Automate потоков

Следующие методы будут бросать ошибки и сбой при призвании из сценария в Power Automate потоке.

| Класс | Метод |
|--|--|
| [Chart](/javascript/api/office-scripts/excelscript/excelscript.chart) | `activate` |
| [Range](/javascript/api/office-scripts/excelscript/excelscript.range) | `select` |
| [Workbook](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getActiveCell` |
| [Workbook](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getActiveChart` |
| [Workbook](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getActiveSlicer` |
| [Workbook](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getSelectedRange` |
| [Workbook](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getSelectedRanges` |

### <a name="script-methods-with-a-default-behavior-in-power-automate-flows"></a>Методы скрипта с поведением по умолчанию Power Automate потоков

Следующие методы используют поведение по умолчанию вместо текущего состояния пользователя.

| Класс | Метод | Power Automate поведение |
|--|--|--|
| [Workbook](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getActiveWorksheet` | Возвращает либо первый лист в рабочей книге, либо лист, активированный `Worksheet.activate` методом. |
| [Worksheet](/javascript/api/office-scripts/excelscript/excelscript.worksheet) | `activate` | Отмечает лист как активный лист для целей `Workbook.getActiveWorksheet` . |

## <a name="select-workbooks-with-the-file-browser-control"></a>Выберите рабочие книги с управлением файлом браузера

При создании **шага сценария Run** Power Automate потока необходимо выбрать, какая рабочая книга является частью потока. Используйте файл браузера, чтобы выбрать вашу трудовую книжку, вместо того, чтобы вручную ввода названия рабочей книги.

:::image type="content" source="../images/power-automate-file-browser.png" alt-text="Действие Power Automate Run, показывающее опцию браузера файла Show Picker":::

Для получения дополнительной Power Automate ограничения и обсуждения потенциальных обходных путей для динамического выбора трудовых книжек, см. эту [тему в microsoft Power Automate Community](https://powerusers.microsoft.com/t5/Power-Automate-Ideas/Allow-for-dynamic-quot-file-quot-value-for-excel-quot-get-a-row/idi-p/103091#).

## <a name="time-zone-differences"></a>Различия в часовом поясе

Excel файлы не имеют неотъемлемого местоположения или часового пояса. Каждый раз, когда пользователь открывает трудовую книжку, его сеанс использует местный часовой пояс пользователя для расчета даты. Power Automate всегда использует UTC.

Если в скрипте используются даты или время, могут возникнуть поведенческие различия при локальном тестировании скрипта по сравнению с тем, когда он Power Automate. Power Automate позволяет конвертировать, форматировать и корректировать время. Можно [найти инструкции по использованию](https://flow.microsoft.com/blog/working-with-dates-and-times/) этих функций в Power Automate и [ `main` параметрах: Передать данные](../develop/power-automate-integration.md#main-parameters-pass-data-to-a-script) в сценарий, чтобы узнать, как предоставить информацию о времени для скрипта.

## <a name="see-also"></a>См. также

- [Устранение неполадок Office скриптов](troubleshooting.md)
- [Вы запустите Office скрипты с Power Automate](../develop/power-automate-integration.md)
- [Excel Онлайн (Бизнес) разъем справочная документация](/connectors/excelonlinebusiness/)
