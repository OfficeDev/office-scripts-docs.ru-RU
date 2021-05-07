---
title: Управление режимом вычисления в Excel
description: Узнайте, как использовать Office скрипты для управления режимом вычисления в Excel в Интернете.
ms.date: 04/28/2021
localization_priority: Normal
ms.openlocfilehash: 34a14874197ffda8487df5e450e3dcab980f7ed5
ms.sourcegitcommit: f7a7aebfb687f2a35dbed07ed62ff352a114525a
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 05/06/2021
ms.locfileid: "52232454"
---
# <a name="manage-calculation-mode-in-excel"></a><span data-ttu-id="63d8a-103">Управление режимом вычисления в Excel</span><span class="sxs-lookup"><span data-stu-id="63d8a-103">Manage calculation mode in Excel</span></span>

<span data-ttu-id="63d8a-104">В этом примере показано, как использовать режим [вычисления](/javascript/api/office-scripts/excelscript/excelscript.calculationmode) и вычислять методы Excel в Интернете с помощью Office скриптов.</span><span class="sxs-lookup"><span data-stu-id="63d8a-104">This sample shows how to use the [calculation mode](/javascript/api/office-scripts/excelscript/excelscript.calculationmode) and calculate methods in Excel on the web using Office Scripts.</span></span> <span data-ttu-id="63d8a-105">Сценарий можно попробовать в любом Excel файле.</span><span class="sxs-lookup"><span data-stu-id="63d8a-105">You can try the script on any Excel file.</span></span>

## <a name="scenario"></a><span data-ttu-id="63d8a-106">Сценарий</span><span class="sxs-lookup"><span data-stu-id="63d8a-106">Scenario</span></span>

<span data-ttu-id="63d8a-107">В Excel в Интернете режим вычисления файла можно управлять программным образом с помощью API.</span><span class="sxs-lookup"><span data-stu-id="63d8a-107">In Excel on the web, a file's calculation mode can be controlled programmatically using APIs.</span></span> <span data-ttu-id="63d8a-108">Следующие действия возможны с помощью Office скриптов.</span><span class="sxs-lookup"><span data-stu-id="63d8a-108">The following actions are possible using Office Scripts.</span></span>

1. <span data-ttu-id="63d8a-109">Получите режим вычисления.</span><span class="sxs-lookup"><span data-stu-id="63d8a-109">Get the calculation mode.</span></span>
1. <span data-ttu-id="63d8a-110">Установите режим вычисления.</span><span class="sxs-lookup"><span data-stu-id="63d8a-110">Set the calculation mode.</span></span>
1. <span data-ttu-id="63d8a-111">Вычислять Excel для файлов, задамых в ручном режиме (также именуемого перерасчетом).</span><span class="sxs-lookup"><span data-stu-id="63d8a-111">Calculate Excel formulas for files that are set to the manual mode (also referred to as recalculate).</span></span>

## <a name="sample-code-control-calculation-mode"></a><span data-ttu-id="63d8a-112">Пример кода: режим вычисления управления</span><span class="sxs-lookup"><span data-stu-id="63d8a-112">Sample code: Control calculation mode</span></span>

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

## <a name="training-video-manage-calculation-mode"></a><span data-ttu-id="63d8a-113">Обучающее видео: управление режимом вычисления</span><span class="sxs-lookup"><span data-stu-id="63d8a-113">Training video: Manage calculation mode</span></span>

<span data-ttu-id="63d8a-114">[Смотреть Sudhi Ramamurthy ходить через этот пример на YouTube](https://youtu.be/iw6O8QH01CI).</span><span class="sxs-lookup"><span data-stu-id="63d8a-114">[Watch Sudhi Ramamurthy walk through this sample on YouTube](https://youtu.be/iw6O8QH01CI).</span></span>
