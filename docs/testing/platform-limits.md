---
title: Ограничения и требования платформы с Office скриптами
description: Ограничения ресурсов и поддержка браузера для Office скриптов при Excel в Интернете
ms.date: 05/17/2021
ms.localizationpriority: medium
ms.openlocfilehash: 2140ebf249af76447f64efae7fd2008e781bf815
ms.sourcegitcommit: d3ed4bdeeba805d97c930394e172e8306a0cf484
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/15/2021
ms.locfileid: "59327877"
---
# <a name="platform-limits-and-requirements-with-office-scripts"></a>Ограничения и требования платформы с Office скриптами

Существуют некоторые ограничения платформы, о которых следует знать при разработке Office скриптов. В этой статье подробно извесятся о поддержке браузера и ограничениях данных для Office скриптов для Excel в Интернете.

## <a name="browser-support"></a>Поддержка браузеров

Office Скрипты работают в любом [браузере, который поддерживает Office для Интернета.](https://support.microsoft.com/office/ad1303e0-a318-47aa-b409-d3a5eb44e452) Однако некоторые функции JavaScript не поддерживаются в Internet Explorer 11 (IE 11). Любые функции, [введенные в ES6 или](https://www.w3schools.com/Js/js_es6.asp) более поздней, не будут работать с IE 11. Если люди в организации по-прежнему используют этот браузер, обязательно проверьте свои скрипты в этой среде при их совместном использовании.

[!INCLUDE [Teams support note](../includes/teams-support-note.md)]

### <a name="third-party-cookies"></a>Сторонние файлы cookie

Вашему браузеру нужны сторонние файлы cookie, включенные для показа вкладки **Automate** в Excel в Интернете. Проверьте параметры браузера, если вкладка не отображается. При использовании закрытого сеанса браузера может потребоваться каждый раз повторно включить этот параметр.

> [!NOTE]
> Некоторые браузеры ссылаются на этот параметр как на "все файлы cookie", а не на "сторонние файлы cookie".

#### <a name="instructions-for-adjusting-cookie-settings-in-popular-browsers"></a>Инструкции по настройке параметров cookie в популярных браузерах

- [Chrome](https://support.google.com/chrome/answer/95647)
- [Microsoft Edge](https://support.microsoft.com/microsoft-edge/597f04f2-c0ce-f08c-7c2b-541086362bd2)
- [Firefox](https://support.mozilla.org/kb/disable-third-party-cookies)
- [Safari](https://support.apple.com/guide/safari/manage-cookies-and-website-data-sfri11471/mac)

## <a name="data-limits"></a>Ограничения данных

Существуют ограничения по объему Excel данных, которые могут быть переданы одновременно, и Power Automate отдельных транзакций.

### <a name="excel"></a>Excel

Excel для Интернета имеет следующие ограничения при вызове в книгу с помощью скрипта:

- Запросы и ответы ограничены **5МБ.**
- Диапазон ограничен пятью **миллионами ячеек.**

Если вы сталкиваетесь с ошибками при работе с большими наборами данных, попробуйте использовать несколько меньших диапазонов вместо больших диапазонов. Пример см. в [примере Write a large dataset](../resources/samples/write-large-dataset.md) sample. Вы также можете использовать API, такие как [Range.getSpecialCells,](/javascript/api/office-scripts/excelscript/excelscript.range#getSpecialCells_cellType__cellValueType_) чтобы нацелить определенные ячейки вместо больших диапазонов.

### <a name="power-automate"></a>Power Automate

При использовании Office скриптов с Power Automate каждый пользователь может использовать **400** вызовов к действию Run Script в день. Это ограничение сбрасывается в 12:00 утра по UTC.

Платформа Power Automate также имеет ограничения использования, которые можно найти в следующих статьях:

- [Ограничения и конфигурация в Power Automate](/power-automate/limits-and-config)
- [Известные проблемы и ограничения для соединиттеля Excel Online (Бизнес)](/connectors/excelonlinebusiness/#known-issues-and-limitations)

## <a name="see-also"></a>См. также

- [Устранение Office скриптов](troubleshooting.md)
- [Отмена эффектов сценариев Office](undo.md)
- [Повышение производительности Office скриптов](../develop/web-client-performance.md)
- [Основы сценариев для Office скриптов в Excel в Интернете](../develop/scripting-fundamentals.md)
