---
title: Сведения об устранении неполадок для Power Automate с помощью скриптов Office
description: Советы, сведения о платформе и известные проблемы с интеграцией сценариев Office и Power Automate.
ms.date: 01/14/2021
localization_priority: Normal
ms.openlocfilehash: b0f5b2f542216789f0d96f309cb7d799d201ba0f
ms.sourcegitcommit: e7e019ba36c2f49451ec08c71a1679eb6dba4268
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 01/22/2021
ms.locfileid: "49933268"
---
# <a name="troubleshooting-information-for-power-automate-with-office-scripts"></a>Сведения об устранении неполадок для Power Automate с помощью скриптов Office

Power Automate позволяет перенаусть автоматизацию сценариев Office на следующий уровень. Однако так как Power Automate запускает сценарии от вашего имени в независимых сеансах Excel, следует отметить несколько важных моментов.

> [!TIP]
> If you're just starting to use Office Scripts with Power Automate, please start with [Run Office Scripts with Power Automate](../develop/power-automate-integration.md) to learn about the platforms.

## <a name="avoid-using-relative-references"></a>Избегайте использования относительных ссылок

Power Automate запускает сценарий в выбранной книге Excel от вашего имени. В этом случае книга может быть закрыта. Любой API, который зависит от текущего состояния пользователя, например, может вести себя иначе `Workbook.getActiveWorksheet` в Power Automate. Это происходит потому, что API-интерфейсы основаны на относительном положении представления или курсора пользователя и эта ссылка не существует в потоке Power Automate.

Некоторые относительные эталонные API-api высылают ошибки в Power Automate. Другие имеют поведение по умолчанию, которое подразумевает состояние пользователя. При разработке сценариев обязательно используйте абсолютные ссылки на таблицы и диапазоны. Это делает поток Power Automate согласованным, даже если переустановка таблиц.

### <a name="script-methods-that-fail-when-run-power-automate-flows"></a>Методы сценариев, которые не работают при запуске потоков Power Automate

Следующие методы выдают ошибку и сбой при ее подоздавке из скрипта в потоке Power Automate.

| Класс | Method |
|--|--|
| [Chart](/javascript/api/office-scripts/excelscript/excelscript.chart) | `activate` |
| [Range](/javascript/api/office-scripts/excelscript/excelscript.range) | `select` |
| [Workbook](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getActiveCell` |
| [Workbook](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getActiveChart` |
| [Workbook](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getActiveSlicer` |
| [Workbook](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getSelectedRange` |
| [Workbook](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getSelectedRanges` |

### <a name="script-methods-with-a-default-behavior-in-power-automate-flows"></a>Методы сценариев с поведением по умолчанию в потоках Power Automate

В следующих методах используется поведение по умолчанию вместо текущего состояния любого пользователя.

| Класс | Method | Поведение Power Automate |
|--|--|--|
| [Workbook](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getActiveWorksheet` | Возвращает либо первый рабочий таблицу в книге, либо таблицу, активированную методом в данный `Worksheet.activate` момент. |
| [Worksheet](/javascript/api/office-scripts/excelscript/excelscript.worksheet) | `activate` | Пометка таблицы в качестве активного для целей `Workbook.getActiveWorksheet` . |

## <a name="select-workbooks-with-the-file-browser-control"></a>Выбор книг с помощью управления браузером файлов

При создании **скрипта запуска** потока Power Automate необходимо выбрать, какая книга является частью потока. Используйте браузер файлов, чтобы выбрать книгу, а не вводить ее имя вручную.

![Параметр браузера файлов при создании действия "Выполнить сценарий" в Power Automate](../images/power-automate-file-browser.png)

Дополнительные контекст ограничения Power Automate и обсуждение возможных обходных обходных пути для динамического выбора книг см. в этом потоке в сообществе [Microsoft Power Automate Community.](https://powerusers.microsoft.com/t5/Power-Automate-Ideas/Allow-for-dynamic-quot-file-quot-value-for-excel-quot-get-a-row/idi-p/103091#)

## <a name="time-zone-differences"></a>Различия в часовом поясе

Файлы Excel не имеют точного расположения или часовой пояс. Каждый раз, когда пользователь открывает книгу, его сеанс использует локальный часовой пояс пользователя для вычисления даты. Power Automate всегда использует UTC.

Если в сценарии используются даты или время, то при локальном и тестовом сценарии могут быть различия в поведении, а не при его запуске с помощью Power Automate. Power Automate позволяет преобразовывать, форматировать и настраивать время. Инструкции по использованию этих функций в Power Automate и [ `main` Parameters:](../develop/power-automate-integration.md#main-parameters-passing-data-to-a-script) Передача данных в сценарий см. в инструкциях по работе с датами и временем в потоках. [](https://flow.microsoft.com/blog/working-with-dates-and-times/)

## <a name="see-also"></a>См. также

- [Устранение неполадок в сценариях Office](troubleshooting.md)
- [Запуск сценариев Office с помощью Power Automate](../develop/power-automate-integration.md)
- [Справочная документация по соединители Excel Online (бизнес)](/connectors/excelonlinebusiness/)
