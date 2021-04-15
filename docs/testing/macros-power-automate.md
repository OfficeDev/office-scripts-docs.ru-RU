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
# <a name="how-to-use-macro-files-in-power-automate-flows"></a><span data-ttu-id="c62c7-103">Использование макрофайлов в потоках Power Automate</span><span class="sxs-lookup"><span data-stu-id="c62c7-103">How to use macro files in Power Automate flows</span></span>

<span data-ttu-id="c62c7-104">[Потоки Power Automate](https://flow.microsoft.com/) предоставляют [соединители Excel,](https://flow.microsoft.com/connectors/shared_excelonlinebusiness/excel-online-business/) которые помогают подключать файлы Excel к остальным организационным данным и приложениям, таким как Teams, Outlook и SharePoint.</span><span class="sxs-lookup"><span data-stu-id="c62c7-104">[Power Automate flows](https://flow.microsoft.com/) provide [Excel connectors](https://flow.microsoft.com/connectors/shared_excelonlinebusiness/excel-online-business/) to help connect Excel files with the rest of your organizational data and apps such as Teams, Outlook, and SharePoint.</span></span>

<span data-ttu-id="c62c7-105">Однако макрофайлы не могут быть выбраны в отсеве файла (см. пример на следующем скриншоте).</span><span class="sxs-lookup"><span data-stu-id="c62c7-105">However, macro files can't be selected in the file dropdown (see an example in the following screenshot).</span></span>

:::image type="content" source="../images/no-xlsm.png" alt-text="Действие скрипта Power Automate Run, в котором не было выбрано макрофайла. Показана ошибка &quot;Файл&quot;.":::

<span data-ttu-id="c62c7-107">Один из способов решения этой проблемы — включите действие "Get File Metadata" (OneDrive или SharePoint) и используйте свойство ID в действии "Сценарий запуска", как показано на следующем скриншоте.</span><span class="sxs-lookup"><span data-stu-id="c62c7-107">One way to get around this issue is by including the "Get File Metadata" action (OneDrive or SharePoint) and use the ID property in the "Run Script" action as shown in the following screenshot.</span></span>

:::image type="content" source="../images/xlsm-in-pa.png" alt-text="Действие скрипта Power Automate Run, показывающая выбранный макрофайл и отсутствие ошибки скрипта Run.":::

> [!NOTE]
> <span data-ttu-id="c62c7-109">Некоторые XLSM (особенно те, которые ActiveX/Form) могут не работать в сетевом соединитель Excel.</span><span class="sxs-lookup"><span data-stu-id="c62c7-109">Some XLSM (especially the ones with ActiveX/Form controls) may not work in the Excel online connector.</span></span> <span data-ttu-id="c62c7-110">Убедитесь, что перед развертыванием решения необходимо протестировать.</span><span class="sxs-lookup"><span data-stu-id="c62c7-110">Be sure to test before deploying your solution.</span></span>

<span data-ttu-id="c62c7-111">[![Просмотр видео об использовании XLSM в действии Run Script](../images/xlsm-vid.png)](https://youtu.be/o-H9BbywJQQ "Видео об использовании XLSM в действии Run Script")</span><span class="sxs-lookup"><span data-stu-id="c62c7-111">[![Watch video about using XLSM in Run Script action](../images/xlsm-vid.png)](https://youtu.be/o-H9BbywJQQ "Video about using XLSM in Run Script action")</span></span>
