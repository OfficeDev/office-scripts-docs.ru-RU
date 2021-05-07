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
# <a name="how-to-use-macro-files-in-power-automate-flows"></a><span data-ttu-id="bd52c-103">Использование макрофайлов в Power Automate потоках</span><span class="sxs-lookup"><span data-stu-id="bd52c-103">How to use macro files in Power Automate flows</span></span>

<span data-ttu-id="bd52c-104">[Power Automate](https://flow.microsoft.com/) потоки предоставляют [Excel](https://flow.microsoft.com/connectors/shared_excelonlinebusiness/excel-online-business/) соединители для подключения Excel файлов с остальными организационными данными и приложениями, такими как Teams, Outlook и SharePoint.</span><span class="sxs-lookup"><span data-stu-id="bd52c-104">[Power Automate flows](https://flow.microsoft.com/) provide [Excel connectors](https://flow.microsoft.com/connectors/shared_excelonlinebusiness/excel-online-business/) to help connect Excel files with the rest of your organizational data and apps such as Teams, Outlook, and SharePoint.</span></span>

<span data-ttu-id="bd52c-105">Однако макрофайлы не могут быть выбраны в отсеве файла (см. пример на следующем скриншоте).</span><span class="sxs-lookup"><span data-stu-id="bd52c-105">However, macro files can't be selected in the file dropdown (see an example in the following screenshot).</span></span>

:::image type="content" source="../images/no-xlsm.png" alt-text="Действие Power Automate запуска скрипта, в котором не было выбрано макрофайла. Показано, что ошибка &quot;Файл&quot; требуется":::

<span data-ttu-id="bd52c-107">Один из способов решения этой проблемы — включив действие "Get File Metadata" (OneDrive или SharePoint) и используйте свойство ID в действии "Сценарий запуска", как показано на следующем скриншоте.</span><span class="sxs-lookup"><span data-stu-id="bd52c-107">One way to get around this issue is by including the "Get File Metadata" action (OneDrive or SharePoint) and use the ID property in the "Run Script" action as shown in the following screenshot.</span></span>

:::image type="content" source="../images/xlsm-in-pa.png" alt-text="Действие Power Automate запуска, показывающая выбранный макрофайл и отсутствие ошибки сценария Run":::

> [!NOTE]
> <span data-ttu-id="bd52c-109">Некоторые XLSM (особенно те, которые ActiveX/Form) могут не работать в сетевом соединителенном Excel.</span><span class="sxs-lookup"><span data-stu-id="bd52c-109">Some XLSM (especially the ones with ActiveX/Form controls) may not work in the Excel online connector.</span></span> <span data-ttu-id="bd52c-110">Убедитесь, что перед развертыванием решения необходимо протестировать.</span><span class="sxs-lookup"><span data-stu-id="bd52c-110">Be sure to test before deploying your solution.</span></span>

## <a name="other-resources"></a><span data-ttu-id="bd52c-111">Другие ресурсы</span><span class="sxs-lookup"><span data-stu-id="bd52c-111">Other resources</span></span>

<span data-ttu-id="bd52c-112">[Просмотрите видео Sudhi Ramamurthy](https://youtu.be/o-H9BbywJQQ)на YouTube о том, как использовать файл xlsm в действии Run Script.</span><span class="sxs-lookup"><span data-stu-id="bd52c-112">[Watch Sudhi Ramamurthy's YouTube video on how use an .xlsm file in a Run Script action](https://youtu.be/o-H9BbywJQQ).</span></span>
