---
title: Использование макрофайлов в потоках Power Automate
description: Узнайте, как использовать макрофайлы или xlsm-файлы в потоках Power Automate.
ms.date: 03/18/2021
localization_priority: Normal
ms.openlocfilehash: ec1fe00eb9ddc382ae4bc02187de7a36c97288b1
ms.sourcegitcommit: 5d24e77df70aa2c1c982275d53213c2a9323ff86
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 04/02/2021
ms.locfileid: "51571476"
---
# <a name="how-to-use-macro-files-in-power-automate-flows"></a>Использование макрофайлов в потоках Power Automate

[Потоки Power Automate](https://flow.microsoft.com/) предоставляют [соединители Excel,](https://flow.microsoft.com/connectors/shared_excelonlinebusiness/excel-online-business/) которые помогают подключать файлы Excel к остальным организационным данным и приложениям, таким как Teams, Outlook и SharePoint.

Однако макрофайлы не могут быть выбраны в отсеве файла (см. пример на следующем скриншоте).

![Нет xlsm в действии Сценарий запуска](../images/no-xlsm.png)

Один из способов решения этой проблемы — включите действие "Get File Metadata" (OneDrive или SharePoint) и используйте свойство ID в действии "Сценарий запуска", как показано на следующем скриншоте.

![xlsm в действии Run Script](../images/xlsm-in-pa.png)

> [!NOTE]
> Некоторые XLSM (особенно те, которые ActiveX/Form) могут не работать в сетевом соединитель Excel. Убедитесь, что перед развертыванием решения необходимо протестировать.

[![Просмотр видео об использовании XLSM в действии Run Script](../images/xlsm-vid.png)](https://youtu.be/o-H9BbywJQQ "Видео об использовании XLSM в действии Run Script")
