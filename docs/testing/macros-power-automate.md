---
title: Использование макрофайлов в потоках Power Automate
description: Узнайте, как использовать макрофайлы или xlsm-файлы в потоках Power Automate.
ms.date: 03/18/2021
localization_priority: Normal
ms.openlocfilehash: a7929fc485ae2118d30a4f2783538d0e04deca2a
ms.sourcegitcommit: 45ffe3dbd2c834b78592ad35928cf8096f5e80bc
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 04/14/2021
ms.locfileid: "51755016"
---
# <a name="how-to-use-macro-files-in-power-automate-flows"></a>Использование макрофайлов в потоках Power Automate

[Потоки Power Automate](https://flow.microsoft.com/) предоставляют [соединители Excel,](https://flow.microsoft.com/connectors/shared_excelonlinebusiness/excel-online-business/) которые помогают подключать файлы Excel к остальным организационным данным и приложениям, таким как Teams, Outlook и SharePoint.

Однако макрофайлы не могут быть выбраны в отсеве файла (см. пример на следующем скриншоте).

:::image type="content" source="../images/no-xlsm.png" alt-text="Действие скрипта Power Automate Run, в котором не было выбрано макрофайла. Показана ошибка &quot;Файл&quot;.":::

Один из способов решения этой проблемы — включите действие "Get File Metadata" (OneDrive или SharePoint) и используйте свойство ID в действии "Сценарий запуска", как показано на следующем скриншоте.

:::image type="content" source="../images/xlsm-in-pa.png" alt-text="Действие скрипта Power Automate Run, показывающая выбранный макрофайл и отсутствие ошибки скрипта Run.":::

> [!NOTE]
> Некоторые XLSM (особенно те, которые ActiveX/Form) могут не работать в сетевом соединитель Excel. Убедитесь, что перед развертыванием решения необходимо протестировать.

[![Просмотр видео об использовании XLSM в действии Run Script](../images/xlsm-vid.png)](https://youtu.be/o-H9BbywJQQ "Видео об использовании XLSM в действии Run Script")
