---
title: Устранение Office скриптов, запущенных в Power Automate
description: Советы, сведения о платформе и известные проблемы с интеграцией между Office скриптами и Power Automate.
ms.date: 11/01/2021
ms.localizationpriority: medium
ms.openlocfilehash: 2c256c2ddc64fcfc510f24e27662234f44b65ac0
ms.sourcegitcommit: 7023b9e23499806901a5ecf8ebc460b76887cca6
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/31/2022
ms.locfileid: "64586033"
---
# <a name="troubleshoot-office-scripts-running-in-power-automate"></a>Устранение Office скриптов, запущенных в Power Automate

Power Automate позволяет выровизировать Office скрипта на следующий уровень. Однако, поскольку Power Automate выполняет скрипты от вашего имени в независимых Excel сеансах, необходимо отметить несколько важных моментов.

> [!TIP]
> Если вы только начинаете использовать Office скрипты с Power Automate, начните с run [Office Scripts](../develop/power-automate-integration.md) с Power Automate, чтобы узнать о платформах.

## <a name="avoid-relative-references"></a>Избегайте относительных ссылок

Power Automate выполняет сценарий в выбранной Excel от вашего имени. Книга может быть закрыта, если это произойдет. Любой API, который зависит от текущего состояния пользователя, `Workbook.getActiveWorksheet`например, может вести себя по-Power Automate. Это происходит потому, что API основаны на относительном расположении представления или курсора пользователя, и эта ссылка не существует в потоке Power Automate.

Некоторые относительные API ссылки бросают ошибки в Power Automate. Другие имеют поведение по умолчанию, которое подразумевает состояние пользователя. При разработке сценариев обязательно используйте абсолютные ссылки для таблиц и диапазонов. Это делает поток Power Automate согласованным, даже если таблицы перенастановки.

### <a name="script-methods-that-fail-when-run-in-power-automate-flows"></a>Методы скрипта, которые не работают при Power Automate потоках

Следующие методы вбрасывать ошибку и сбой при призыве из скрипта в потоке Power Automate.

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
| [Workbook](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getActiveWorksheet` | Возвращает либо первую таблицу в книге, либо таблицу, активированную методом `Worksheet.activate` . |
| [Worksheet](/javascript/api/office-scripts/excelscript/excelscript.worksheet) | `activate` | Отмечает таблицу как активную таблицу для целей `Workbook.getActiveWorksheet`. |

## <a name="data-refresh-not-supported-in-power-automate"></a>Обновление данных не поддерживается в Power Automate

Office скрипты не могут обновлять данные при запуске Power Automate. Такие методы, как `PivotTable.refresh` не делают ничего, когда они вызваны в потоке. Кроме того, Power Automate не запускает обновление данных для формул, которые используют ссылки на книги.

### <a name="script-methods-that-do-nothing-when-run-in-power-automate-flows"></a>Методы скрипта, которые ничего не делают при Power Automate потоках

Следующие методы ничего не делают в скрипте при Power Automate. Они по-прежнему успешно возвращаются и не выбрасывают ошибок.

| Класс | Метод |
|--|--|
| [PivotTable](/javascript/api/office-scripts/excelscript/excelscript.pivottable) | `refresh` |
| [Workbook](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `refreshAllDataConnections` |
| [Workbook](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `refreshAllPivotTables` |
| [Worksheet](/javascript/api/office-scripts/excelscript/excelscript.worksheet) | `refreshAllPivotTables` |

## <a name="select-workbooks-with-the-file-browser-control"></a>Выбор книг с управлением браузером файлов

При создании шага **сценария run** из потока Power Automate необходимо выбрать, какая книга является частью потока. Используйте браузер файлов, чтобы выбрать книгу, а не вручную вводить имя книги.

:::image type="content" source="../images/power-automate-file-browser.png" alt-text="Действие Power Automate запуска скрипта, показывающая параметр браузера файлов Show Picker.":::

Дополнительные возможности для Power Automate и обсуждения потенциальных обходных пути динамического выбора книг см. в этом потоке в [microsoft Power Automate Community](https://powerusers.microsoft.com/t5/Power-Automate-Ideas/Allow-for-dynamic-quot-file-quot-value-for-excel-quot-get-a-row/idi-p/103091#).

## <a name="pass-entire-arrays-as-script-parameters"></a>Передать целые массивы в качестве параметров скрипта

Power Automate позволяет пользователям передавать массивы соединитетелям в качестве переменной или как отдельные элементы в массиве. По умолчанию необходимо передать отдельные элементы, которые строят массив в потоке. Для скриптов или других соединители, которые принимают целые массивы в качестве аргументов, необходимо выбрать кнопку **Переключатель** для ввода всего массива, чтобы передать массив как один полный объект. Эта кнопка находится в верхнем правом углу каждого поля ввода параметров массива.

:::image type="content" source="../images/combine-worksheets-flow-3.png" alt-text="Кнопка для перехода на ввод всего массива в поле ввода поля управления.":::

## <a name="time-zone-differences"></a>Различия часовой зоны

Excel файлы не имеют неотъемлемого расположения или часовой пояс. Каждый раз, когда пользователь открывает книгу, его сеанс использует локальный часовой пояс пользователя для расчетов дат. Power Automate всегда использует UTC.

Если в сценарии используются даты или время, при локальной проверке скрипта могут возникнуть различия в поведении по сравнению с тем, когда он Power Automate. Power Automate позволяет преобразовывать, форматировать и настраивать время. Сведения [](https://flow.microsoft.com/blog/working-with-dates-and-times/) о том, как использовать эти функции в Power Automate [`main` и Параметры:](../develop/power-automate-integration.md#main-parameters-pass-data-to-a-script) передайте данные в скрипт, чтобы узнать, как предоставить эти сведения о времени для сценария, см. в этой ссылке.

## <a name="script-parameter-fields-or-returned-output-not-appearing-in-power-automate"></a>Поля параметров скрипта или возвращаемая производительность, не отображаются в Power Automate

Существует две причины, по которой параметры или возвращенные данные скрипта не отражаются точно в Power Automate потоке.

- Подпись скрипта (параметры или значение возврата) изменилась с момента добавленного **соединиттеля Excel Бизнеса (Online**).
- Подпись скрипта использует неподтверченные типы. Проверьте типы в списках по параметрам и [](../develop/power-automate-integration.md#main-parameters-pass-data-to-a-script) возвращает разделы [](../develop/power-automate-integration.md#return-data-from-a-script) [Run Office с помощью Power Automate](../develop/power-automate-integration.md) статьи.

Подпись скрипта хранится с соединитетелем **Excel (Online)** при его создания. Удалите старый соединитатель и создайте новый, чтобы получить последние параметры и значения возврата для действия **скрипта Run** .

## <a name="see-also"></a>См. также

- [Устранение Office скриптов](troubleshooting.md)
- [Запустите Office скрипты с Power Automate](../develop/power-automate-integration.md)
- [Excel справочная документация по соединители Online (Business)](/connectors/excelonlinebusiness/)
