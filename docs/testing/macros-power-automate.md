---
title: Использование макрофайлов в Power Automate потоках
description: Узнайте, как использовать макрофайлы или xlsm-файлы в Power Automate потоках.
ms.date: 04/28/2021
localization_priority: Normal
ms.openlocfilehash: b232a1d31a7ff6e28016c5e28fd8a83c8d3f1859
ms.sourcegitcommit: f7a7aebfb687f2a35dbed07ed62ff352a114525a
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 05/06/2021
ms.locfileid: "52232657"
---
# <a name="how-to-use-macro-files-in-power-automate-flows"></a>Использование макрофайлов в Power Automate потоках

[Power Automate](https://flow.microsoft.com/) потоки предоставляют [Excel](https://flow.microsoft.com/connectors/shared_excelonlinebusiness/excel-online-business/) соединители для подключения Excel файлов с остальными организационными данными и приложениями, такими как Teams, Outlook и SharePoint.

Однако макрофайлы не могут быть выбраны в отсеве файла (см. пример на следующем скриншоте).

:::image type="content" source="../images/no-xlsm.png" alt-text="Действие Power Automate запуска скрипта, в котором не было выбрано макрофайла. Показано, что ошибка &quot;Файл&quot; требуется":::

Один из способов решения этой проблемы — включив действие "Get File Metadata" (OneDrive или SharePoint) и используйте свойство ID в действии "Сценарий запуска", как показано на следующем скриншоте.

:::image type="content" source="../images/xlsm-in-pa.png" alt-text="Действие Power Automate запуска, показывающая выбранный макрофайл и отсутствие ошибки сценария Run":::

> [!NOTE]
> Некоторые XLSM (особенно те, которые ActiveX/Form) могут не работать в сетевом соединителенном Excel. Убедитесь, что перед развертыванием решения необходимо протестировать.

## <a name="other-resources"></a>Другие ресурсы

[Просмотрите видео Sudhi Ramamurthy](https://youtu.be/o-H9BbywJQQ)на YouTube о том, как использовать файл xlsm в действии Run Script.
