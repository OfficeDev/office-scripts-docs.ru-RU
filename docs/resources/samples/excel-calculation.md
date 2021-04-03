---
title: Управление режимом вычисления в Excel
description: Узнайте, как использовать скрипты Office для управления режимом вычислений в Excel в Интернете.
ms.date: 03/18/2021
localization_priority: Normal
ms.openlocfilehash: 0239437c7b52dca1fd8d1a4fc66bab7965cbd91a
ms.sourcegitcommit: 5d24e77df70aa2c1c982275d53213c2a9323ff86
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 04/02/2021
ms.locfileid: "51571529"
---
# <a name="manage-calculation-mode-in-excel"></a><span data-ttu-id="39352-103">Управление режимом вычисления в Excel</span><span class="sxs-lookup"><span data-stu-id="39352-103">Manage calculation mode in Excel</span></span>

<span data-ttu-id="39352-104">В этом примере показано, как использовать режим [вычисления](/javascript/api/office-scripts/excelscript/excelscript.calculationmode) и вычислять методы в Excel в Интернете с помощью Office Scripts.</span><span class="sxs-lookup"><span data-stu-id="39352-104">This sample shows how to use the [calculation mode](/javascript/api/office-scripts/excelscript/excelscript.calculationmode) and calculate methods in Excel on the web using Office Scripts.</span></span> <span data-ttu-id="39352-105">Сценарий можно попробовать в любом файле Excel.</span><span class="sxs-lookup"><span data-stu-id="39352-105">You can try the script on any Excel file.</span></span>

## <a name="scenario"></a><span data-ttu-id="39352-106">Сценарий</span><span class="sxs-lookup"><span data-stu-id="39352-106">Scenario</span></span>

<span data-ttu-id="39352-107">В Excel в Интернете режим вычисления файла можно управлять программным образом с помощью API.</span><span class="sxs-lookup"><span data-stu-id="39352-107">In Excel on the web, a file's calculation mode can be controlled programmatically using APIs.</span></span> <span data-ttu-id="39352-108">Следующие действия возможны с помощью скриптов Office.</span><span class="sxs-lookup"><span data-stu-id="39352-108">The following actions are possible using Office Scripts.</span></span>

1. <span data-ttu-id="39352-109">Получите режим вычисления.</span><span class="sxs-lookup"><span data-stu-id="39352-109">Get the calculation mode.</span></span>
1. <span data-ttu-id="39352-110">Установите режим вычисления.</span><span class="sxs-lookup"><span data-stu-id="39352-110">Set the calculation mode.</span></span>
1. <span data-ttu-id="39352-111">Вычислять формулы Excel для файлов, задамых в ручном режиме (также именуемого перерасчетом).</span><span class="sxs-lookup"><span data-stu-id="39352-111">Calculate Excel formulas for files that are set to the manual mode (also referred to as recalculate).</span></span>

## <a name="sample-code-control-calculation-mode"></a><span data-ttu-id="39352-112">Пример кода: режим вычисления управления</span><span class="sxs-lookup"><span data-stu-id="39352-112">Sample code: Control calculation mode</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
    // Set calculation mode.
    workbook.getApplication().setCalculationMode(ExcelScript.CalculationMode.manual);
    // Get calculation mode.
    const calcMode = workbook.getApplication().getCalculationMode();    
    console.log(calcMode);
    // Calculate (for manual mode files).
    workbook.getApplication().calculate(ExcelScript.CalculationType.full);
}
```

## <a name="training-video-manage-calculation-mode"></a><span data-ttu-id="39352-113">Обучающее видео: управление режимом вычисления</span><span class="sxs-lookup"><span data-stu-id="39352-113">Training video: Manage calculation mode</span></span>

<span data-ttu-id="39352-114">[![Просмотр пошагового видео об управлении режимом вычислений в Excel в Интернете](../../images/calc-mode-vid.jpg)](https://youtu.be/iw6O8QH01CI "Пошаговая видеокадры об управлении режимом вычислений в Excel в Интернете")</span><span class="sxs-lookup"><span data-stu-id="39352-114">[![Watch step-by-step video on how to manage calculation mode in Excel on the web](../../images/calc-mode-vid.jpg)](https://youtu.be/iw6O8QH01CI "Step-by-step video on how to manage calculation mode in Excel on the web")</span></span>
