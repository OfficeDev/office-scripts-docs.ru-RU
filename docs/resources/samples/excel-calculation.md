---
title: Управление режимом вычисления в Excel
description: Узнайте, как использовать Office скрипты для управления режимом вычисления в Excel в Интернете.
ms.date: 05/06/2021
localization_priority: Normal
ms.openlocfilehash: a60fddc91b3a8f124a44722d0d75e6e9f239351d
ms.sourcegitcommit: 763d341857bcb209b2f2c278a82fdb63d0e18f0a
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 05/08/2021
ms.locfileid: "52285915"
---
# <a name="manage-calculation-mode-in-excel"></a><span data-ttu-id="23438-103">Управление режимом вычисления в Excel</span><span class="sxs-lookup"><span data-stu-id="23438-103">Manage calculation mode in Excel</span></span>

<span data-ttu-id="23438-104">В этом примере показано, как использовать режим [вычисления](/javascript/api/office-scripts/excelscript/excelscript.calculationmode) и вычислять методы Excel в Интернете с помощью Office скриптов.</span><span class="sxs-lookup"><span data-stu-id="23438-104">This sample shows how to use the [calculation mode](/javascript/api/office-scripts/excelscript/excelscript.calculationmode) and calculate methods in Excel on the web using Office Scripts.</span></span> <span data-ttu-id="23438-105">Сценарий можно попробовать в любом Excel файле.</span><span class="sxs-lookup"><span data-stu-id="23438-105">You can try the script on any Excel file.</span></span>

## <a name="scenario"></a><span data-ttu-id="23438-106">Сценарий</span><span class="sxs-lookup"><span data-stu-id="23438-106">Scenario</span></span>

<span data-ttu-id="23438-107">Перерасчет книг с большим количеством формул может занять некоторое время.</span><span class="sxs-lookup"><span data-stu-id="23438-107">Workbooks with large numbers of formulas can take a while to recalculate.</span></span> <span data-ttu-id="23438-108">Вместо того, чтобы Excel управлять при вычислениях, вы можете управлять ими в рамках сценария.</span><span class="sxs-lookup"><span data-stu-id="23438-108">Rather than letting Excel control when calculations happen, you can manage them as part of your script.</span></span> <span data-ttu-id="23438-109">Это поможет с производительностью в определенных сценариях.</span><span class="sxs-lookup"><span data-stu-id="23438-109">This will help with performance in certain scenarios.</span></span>

<span data-ttu-id="23438-110">Пример сценария задает режим вычисления вручную.</span><span class="sxs-lookup"><span data-stu-id="23438-110">The sample script sets the calculation mode to manual.</span></span> <span data-ttu-id="23438-111">Это означает, что книга будет пересчитывать формулы только в том случае, если сценарий подсказывает ему (или вручную вычисляется с помощью [пользовательского интерфейса).](https://support.microsoft.com/office/change-formula-recalculation-iteration-or-precision-in-excel-73fc7dac-91cf-4d36-86e8-67124f6bcce4)</span><span class="sxs-lookup"><span data-stu-id="23438-111">This means that the workbook will only recalculate formulas when the script tells it to (or you [manually calculate through the UI](https://support.microsoft.com/office/change-formula-recalculation-iteration-or-precision-in-excel-73fc7dac-91cf-4d36-86e8-67124f6bcce4)).</span></span> <span data-ttu-id="23438-112">Затем сценарий отображает текущий режим вычислений и полностью пересчитывает всю книгу.</span><span class="sxs-lookup"><span data-stu-id="23438-112">The script then displays the current calculation mode and fully recalculates the entire workbook.</span></span>

## <a name="sample-code-control-calculation-mode"></a><span data-ttu-id="23438-113">Пример кода: режим вычисления управления</span><span class="sxs-lookup"><span data-stu-id="23438-113">Sample code: Control calculation mode</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
    // Set the calculation mode to manual.
    workbook.getApplication().setCalculationMode(ExcelScript.CalculationMode.manual);
    // Get and log the calculation mode.
    const calcMode = workbook.getApplication().getCalculationMode();    
    console.log(calcMode);
    // Manually calculate the file.
    workbook.getApplication().calculate(ExcelScript.CalculationType.full);
}
```

## <a name="training-video-manage-calculation-mode"></a><span data-ttu-id="23438-114">Обучающее видео: управление режимом вычисления</span><span class="sxs-lookup"><span data-stu-id="23438-114">Training video: Manage calculation mode</span></span>

<span data-ttu-id="23438-115">[Смотреть Sudhi Ramamurthy ходить через этот пример на YouTube](https://youtu.be/iw6O8QH01CI).</span><span class="sxs-lookup"><span data-stu-id="23438-115">[Watch Sudhi Ramamurthy walk through this sample on YouTube](https://youtu.be/iw6O8QH01CI).</span></span>
