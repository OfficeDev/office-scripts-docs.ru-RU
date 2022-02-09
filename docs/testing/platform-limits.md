---
title: Ограничения и требования платформы с Office скриптами
description: Ограничения ресурсов и поддержка браузера для Office скриптов при Excel в Интернете.
ms.date: 01/31/2022
ms.localizationpriority: medium
ms.openlocfilehash: 385248e5c62ed3dbf2827105b3097ef27e5187a7
ms.sourcegitcommit: b84d4c8dd31335e4e39b0da6ad25fd528cb9d8f3
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 02/09/2022
ms.locfileid: "62462504"
---
# <a name="platform-limits-and-requirements-with-office-scripts"></a>Ограничения и требования платформы с Office скриптами

Существуют некоторые ограничения платформы, о которых следует знать при разработке Office скриптов. В этой статье подробно извесятся о поддержке браузера и ограничениях данных для Office скриптов для Excel в Интернете.

## <a name="browser-support"></a>Поддержка браузеров

Office скрипты работают в любом браузере[, который поддерживает Office для Интернета](https://support.microsoft.com/office/ad1303e0-a318-47aa-b409-d3a5eb44e452). Однако некоторые функции JavaScript не поддерживаются в Internet Explorer 11 (IE 11). Любые функции, [введенные в ES6 или](https://www.w3schools.com/Js/js_es6.asp) более поздней, не будут работать с IE 11. Если люди в организации по-прежнему используют этот браузер, обязательно проверьте свои скрипты в этой среде при их совместном использовании.

[!INCLUDE [Teams support note](../includes/teams-support-note.md)]

### <a name="third-party-cookies"></a>Сторонние файлы cookie

Вашему браузеру необходимы сторонние файлы cookie, включенные для показа вкладки **Automate** в Excel в Интернете. Проверьте параметры браузера, если вкладка не отображается. При использовании закрытого сеанса браузера может потребоваться каждый раз повторно включить этот параметр.

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

- Количество запросов и ответов ограничено **5 МБ**.
- Диапазон ограничен пятью **миллионами ячеек**.

Если вы сталкиваетесь с ошибками при работе с большими наборами данных, попробуйте использовать несколько меньших диапазонов вместо больших диапазонов. Пример см. в [примере Write a large dataset](../resources/samples/write-large-dataset.md) sample. Вы также можете использовать API, такие как [Range.getSpecialCells](/javascript/api/office-scripts/excelscript/excelscript.range#excelscript-excelscript-range-getspecialcells-member(1)) , чтобы нацелить определенные ячейки вместо больших диапазонов.

### <a name="power-automate"></a>Power Automate

При использовании Office скриптов с Power Automate каждый пользователь может использовать **1600 вызовов для действия Run Script в день**. Это ограничение сбрасывается в 12:00 утра по UTC.

Платформа Power Automate также имеет ограничения использования, которые можно найти в следующих статьях.

- [Ограничения и конфигурация в Power Automate](/power-automate/limits-and-config)
- [Известные проблемы и ограничения для соединиттеля Excel Online (Бизнес)](/connectors/excelonlinebusiness/#known-issues-and-limitations)

> [!NOTE]
> Если у вас есть длительный сценарий, следует помнить о [120-секундном](/power-automate/limits-and-config#timeout) таймауте для синхронных Power Automate операций. Необходимо либо оптимизировать сценарий[](../develop/web-client-performance.md), либо разделить Excel на несколько скриптов.

## <a name="see-also"></a>См. также

- [Устранение Office скриптов](troubleshooting.md)
- [Отмена эффектов сценариев Office](undo.md)
- [Повышение производительности Office скриптов](../develop/web-client-performance.md)
- [Основы сценариев для Office скриптов в Excel в Интернете](../develop/scripting-fundamentals.md)
