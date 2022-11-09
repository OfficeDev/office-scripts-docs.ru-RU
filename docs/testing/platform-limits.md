---
title: Ограничения платформы и требования к ним с помощью сценариев Office
description: Ограничения ресурсов и поддержка браузеров для сценариев Office при использовании с Excel в Интернете.
ms.date: 11/07/2022
ms.localizationpriority: medium
ms.openlocfilehash: 764d1eddaf303a941a098ec1d3f3056d63e8693f
ms.sourcegitcommit: 7cadf2b637bf62874e43b6e595286101816662aa
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 11/09/2022
ms.locfileid: "68891248"
---
# <a name="platform-limits-and-requirements-with-office-scripts"></a>Ограничения платформы и требования к ним с помощью сценариев Office

Существуют некоторые ограничения платформы, которые следует учитывать при разработке сценариев Office. В этой статье подробно описана поддержка браузера и ограничения данных для сценариев Office для Excel в Интернете.

## <a name="browser-support"></a>Поддержка браузеров

Сценарии Office работают в любом браузере, [поддерживающем Office для Интернета](https://support.microsoft.com/office/ad1303e0-a318-47aa-b409-d3a5eb44e452). Однако некоторые функции JavaScript не поддерживаются в Internet Explorer 11 (IE 11). Все функции, представленные в [ES6 или более поздних версиях](https://www.w3schools.com/Js/js_es6.asp) , не будут работать с IE 11. Если пользователи в вашей организации по-прежнему используют этот браузер, обязательно протестируйте скрипты в этой среде при предоставлении общего доступа к ним.

[!INCLUDE [Teams support note](../includes/teams-support-note.md)]

### <a name="third-party-cookies"></a>Сторонние файлы cookie

В браузере должны быть включены сторонние файлы cookie, чтобы отобразить вкладку **Автоматизация** в Excel в Интернете. Проверьте параметры браузера, если вкладка не отображается. Если вы используете закрытый сеанс браузера, вам может потребоваться каждый раз повторно включать этот параметр.

> [!NOTE]
> В некоторых браузерах этот параметр называется "все файлы cookie", а не "сторонние файлы cookie".

#### <a name="instructions-for-adjusting-cookie-settings-in-popular-browsers"></a>Инструкции по настройке параметров файлов cookie в популярных браузерах

- [Chrome](https://support.google.com/chrome/answer/95647)
- [Microsoft Edge](https://support.microsoft.com/microsoft-edge/597f04f2-c0ce-f08c-7c2b-541086362bd2)
- [Firefox](https://support.mozilla.org/kb/disable-third-party-cookies)
- [Safari](https://support.apple.com/guide/safari/manage-cookies-and-website-data-sfri11471/mac)

## <a name="data-limits"></a>Ограничения данных

Существуют ограничения на то, сколько данных Excel может быть передано одновременно и сколько отдельных транзакций Power Automate может быть выполнено.

### <a name="excel"></a>Excel

Excel для Интернета имеет следующие ограничения при вызове книги с помощью скрипта:

- Объем запросов и ответов ограничен **5 МБ**.
- Диапазон ограничен **пятью миллионами ячеек**.

Если при работе с большими наборами данных возникают ошибки, попробуйте использовать несколько меньших диапазонов вместо больших диапазонов. Пример см. в примере [записи большого набора данных](../resources/samples/write-large-dataset.md) . Вы также можете использовать ТАКИЕ API, как [Range.getSpecialCells](/javascript/api/office-scripts/excelscript/excelscript.range#excelscript-excelscript-range-getspecialcells-member(1)) , для целевых ячеек вместо больших диапазонов.

Ограничения Excel, не относящиеся к сценариям Office, см. в статье [Спецификации и ограничения Excel](https://support.microsoft.com/office/excel-specifications-and-limits-1672b34d-7043-467e-8e27-269d656771c3).

### <a name="power-automate"></a>Power Automate

При использовании сценариев Office с Power Automate каждый пользователь может выполнять только **1600 вызовов к действию Выполнение скрипта в день**. Это ограничение сбрасывается в 12:00 UTC.

Платформа Power Automate также имеет ограничения на использование, которые можно найти в следующих статьях.

- [Ограничения и настройка в Power Automate](/power-automate/limits-and-config)
- [Известные проблемы и ограничения соединителя Excel Online (бизнес)](/connectors/excelonlinebusiness/#known-issues-and-limitations)

> [!NOTE]
> Если у вас есть длительный сценарий, имейте в виду [время ожидания 120 секунд для синхронных операций Power Automate](/power-automate/limits-and-config#timeout). Вам потребуется [либо оптимизировать скрипт](../develop/web-client-performance.md) , либо разделить автоматизацию Excel на несколько сценариев.

## <a name="see-also"></a>См. также

- [Спецификации и ограничения Excel](https://support.microsoft.com/office/excel-specifications-and-limits-1672b34d-7043-467e-8e27-269d656771c3)
- [Устранение неполадок со сценариями Office](troubleshooting.md)
- [Отмена эффектов сценариев Office](undo.md)
- [Повышение производительности сценариев Office](../develop/web-client-performance.md)
