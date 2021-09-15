---
title: Устранение Office скриптов, запущенных в Power Automate
description: Советы, сведения о платформе и известные проблемы с интеграцией между Office и Power Automate.
ms.date: 05/18/2021
ms.localizationpriority: medium
ms.openlocfilehash: aa0602720233afddd88ccfb8ee86d3934892a05f
ms.sourcegitcommit: d3ed4bdeeba805d97c930394e172e8306a0cf484
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/15/2021
ms.locfileid: "59326851"
---
# <a name="troubleshoot-office-scripts-running-in-power-automate"></a>Устранение Office скриптов, запущенных в Power Automate

Power Automate позволяет выровизировать Office скрипта на следующий уровень. Однако, Power Automate сценарии от вашего имени в независимых сеансах Excel, необходимо отметить несколько важных моментов.

> [!TIP]
> Если вы только начинаете использовать Office скрипты с Power Automate, начните с run [Office Scripts с](../develop/power-automate-integration.md) Power Automate, чтобы узнать о платформах.

## <a name="avoid-relative-references"></a>Избегайте относительных ссылок

Power Automate сценарий выполняется в выбранной Excel от вашего имени. Книга может быть закрыта, если это произойдет. Любой API, который зависит от текущего состояния пользователя, например, может вести себя по-другому в `Workbook.getActiveWorksheet` Power Automate. Это происходит потому, что API основаны на относительном расположении представления или курсора пользователя, и эта ссылка не существует в потоке Power Automate.

Некоторые относительные API ссылки бросают ошибки в Power Automate. Другие имеют поведение по умолчанию, которое подразумевает состояние пользователя. При разработке сценариев обязательно используйте абсолютные ссылки для таблиц и диапазонов. Это делает поток Power Automate согласованным, даже если таблицы переостановки.

### <a name="script-methods-that-fail-when-run-in-power-automate-flows"></a>Методы скрипта, которые не работают при Power Automate потоках

Следующие методы вбрасывать ошибку и сбой при призыве из сценария в потоке Power Automate.

| Класс | Метод |
|--|--|
| [Chart](/javascript/api/office-scripts/excelscript/excelscript.chart) | `activate` |
| [Range](/javascript/api/office-scripts/excelscript/excelscript.range) | `select` |
| [Workbook](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getActiveCell` |
| [Workbook](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getActiveChart` |
| [Workbook](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getActiveSlicer` |
| [Workbook](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getSelectedRange` |
| [Workbook](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getSelectedRanges` |

### <a name="script-methods-with-a-default-behavior-in-power-automate-flows"></a>Методы скрипта с поведением по умолчанию в Power Automate потоках

В следующих методах используется поведение по умолчанию вместо текущего состояния любого пользователя.

| Класс | Метод | Power Automate поведения |
|--|--|--|
| [Workbook](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getActiveWorksheet` | Возвращает либо первую таблицу в книге, либо таблицу, активированную `Worksheet.activate` методом. |
| [Worksheet](/javascript/api/office-scripts/excelscript/excelscript.worksheet) | `activate` | Отмечает таблицу как активную таблицу для целей `Workbook.getActiveWorksheet` . |

## <a name="data-refresh-not-supported-in-power-automate"></a>Обновление данных не поддерживается в Power Automate

Office Скрипты не могут обновлять данные при Power Automate. Такие методы, `PivotTable.refresh` как не делают ничего, когда они вызваны в потоке. Кроме того, Power Automate не запускает обновление данных для формул, которые используют ссылки на книги.

### <a name="script-methods-that-do-nothing-when-run-in-power-automate-flows"></a>Методы скрипта, которые ничего не делают при Power Automate потоках

Следующие методы ничего не делают в скрипте при Power Automate. Они по-прежнему успешно возвращаются и не выбрасывают ошибок.

| Класс | Метод |
|--|--|
| [PivotTable](/javascript/api/office-scripts/excelscript/excelscript.pivottable) | `refresh` |
| [Workbook](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `refreshAllDataConnections` |
| [Workbook](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `refreshAllPivotTables` |
| [Worksheet](/javascript/api/office-scripts/excelscript/excelscript.worksheet) | `refreshAllPivotTables` |

## <a name="select-workbooks-with-the-file-browser-control"></a>Выбор книг с управлением браузером файлов

При создании шага **сценария run** Power Automate потока необходимо выбрать, какая книга является частью потока. Используйте браузер файлов, чтобы выбрать книгу, а не вручную вводить имя книги.

:::image type="content" source="../images/power-automate-file-browser.png" alt-text="Действие Power Automate запуска скрипта, показывающая параметр браузера файлов Show Picker.":::

Дополнительные контексты Power Automate ограничения и обсуждения потенциальных обходных пути для динамического выбора книг см. в этом потоке в [Microsoft Power Automate Community](https://powerusers.microsoft.com/t5/Power-Automate-Ideas/Allow-for-dynamic-quot-file-quot-value-for-excel-quot-get-a-row/idi-p/103091#).

## <a name="time-zone-differences"></a>Различия часовой зоны

Excel файлы не имеют неотъемлемого расположения или часовой пояс. Каждый раз, когда пользователь открывает книгу, его сеанс использует локальный часовой пояс пользователя для расчетов дат. Power Automate всегда использует UTC.

Если в сценарии используются даты или время, при локальной проверке скрипта могут возникнуть различия в поведении по сравнению с тем, когда он Power Automate. Power Automate позволяет преобразовывать, форматировать и настраивать время. Сведения [](https://flow.microsoft.com/blog/working-with-dates-and-times/) о том, как использовать эти функции в Power Automate и [ `main` Параметры,](../develop/power-automate-integration.md#main-parameters-pass-data-to-a-script) см. в инструкции по работе с датами и временем внутри потоков.

## <a name="see-also"></a>См. также

- [Устранение Office скриптов](troubleshooting.md)
- [Запустите Office скрипты с Power Automate](../develop/power-automate-integration.md)
- [Excel Справочная документация по соединители online (Business)](/connectors/excelonlinebusiness/)
