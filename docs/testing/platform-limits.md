---
title: Ограничения и требования платформы с Office скриптами
description: Ограничения ресурсов и поддержка браузера для Office скриптов при использовании с Excel в Интернете
ms.date: 05/17/2021
localization_priority: Normal
ms.openlocfilehash: 7e81aaf2f96faeb67c815814fe3b7f1795651318
ms.sourcegitcommit: 4687693f02fc90a57ba30c461f35046e02e6f5fb
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 05/19/2021
ms.locfileid: "52545583"
---
# <a name="platform-limits-and-requirements-with-office-scripts"></a>Ограничения и требования платформы с Office скриптами

Есть некоторые ограничения платформы, о которых вы должны знать при разработке Office скриптов. В этой статье подробно подробно поддержки браузера и ограничения данных для Office скриптов для Excel в Интернете.

## <a name="browser-support"></a>Поддержка браузеров

Office Скрипты работают в любом [браузере, который Office для Интернета.](https://support.microsoft.com/office/ad1303e0-a318-47aa-b409-d3a5eb44e452) Однако некоторые функции JavaScript не поддерживаются в Internet Explorer 11 (IE 11). Любые функции, [введенные в ES6 или позже,](https://www.w3schools.com/Js/js_es6.asp) не будут работать с IE 11. Если люди в вашей организации по-прежнему используют этот браузер, обязательно проверьте ваши скрипты в этой среде при их совместном использовании.

[!INCLUDE [Teams support note](../includes/teams-support-note.md)]

### <a name="third-party-cookies"></a>Сторонние файлы cookie

Вашему браузеру нужны сторонние файлы cookie, которые могут показывать **вкладку Automate** в Excel в Интернете. Проверьте настройки браузера, если вкладка не отображается. Если вы используете сеанс частного браузера, возможно, потребуется каждый раз повторно включать эту настройку.

> [!NOTE]
> Некоторые браузеры называют эту настройку «всеми файлами cookie», а не «сторонними файлами cookie».

#### <a name="instructions-for-adjusting-cookie-settings-in-popular-browsers"></a>Инструкции по настройке настроек файлов cookie в популярных браузерах

- [Chrome](https://support.google.com/chrome/answer/95647)
- [Microsoft Edge](https://support.microsoft.com/microsoft-edge/temporarily-allow-cookies-and-site-data-in-microsoft-edge-597f04f2-c0ce-f08c-7c2b-541086362bd2)
- [Firefox](https://support.mozilla.org/kb/disable-third-party-cookies)
- [Safari](https://support.apple.com/guide/safari/manage-cookies-and-website-data-sfri11471/mac)

## <a name="data-limits"></a>Ограничения данных

Существуют ограничения на то, Excel данные могут быть переданы сразу и сколько отдельных Power Automate транзакций могут быть проведены.

### <a name="excel"></a>Excel

Excel для Интернета имеет следующие ограничения при звонках в трудовую книжку через скрипт:

- Запросы и ответы ограничены **5MB**.
- Диапазон ограничен пятью **миллионами ячеек.**

Если вы столкнулись с ошибками при работе с большими наборами данных, попробуйте использовать несколько меньших диапазонов вместо больших диапазонов. Например, [см.](../resources/samples/write-large-dataset.md) Вы также можете использовать API, такие как [Range.getSpecialCells, для таргетинга](/javascript/api/office-scripts/excelscript/excelscript.range#getspecialcells-celltype--cellvaluetype-) на определенные ячейки вместо больших диапазонов.

### <a name="power-automate"></a>Power Automate

При использовании Office скриптов Power Automate, каждый пользователь ограничен **400 вызовов на run Script действий в день**. Этот лимит сбрасывается в 12:00 UTC.

Платформа Power Automate также имеет ограничения использования, которые можно найти в следующих статьях:

- [Ограничения и конфигурация в Power Automate](/power-automate/limits-and-config)
- [Известные проблемы и ограничения для Excel Online (Бизнес) разъем](/connectors/excelonlinebusiness/#known-issues-and-limitations)

## <a name="see-also"></a>См. также

- [Устранение неполадок Office скриптов](troubleshooting.md)
- [Отмена эффектов сценариев Office](undo.md)
- [Улучшение производительности ваших Office скриптов](../develop/web-client-performance.md)
- [Основы сценариев для Office сценариев в Excel в Интернете](../develop/scripting-fundamentals.md)
