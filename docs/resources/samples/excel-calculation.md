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
# <a name="manage-calculation-mode-in-excel"></a>Управление режимом вычисления в Excel

В этом примере показано, как использовать режим [вычисления](/javascript/api/office-scripts/excelscript/excelscript.calculationmode) и вычислять методы Excel в Интернете с помощью Office скриптов. Сценарий можно попробовать в любом Excel файле.

## <a name="scenario"></a>Сценарий

Перерасчет книг с большим количеством формул может занять некоторое время. Вместо того, чтобы Excel управлять при вычислениях, вы можете управлять ими в рамках сценария. Это поможет с производительностью в определенных сценариях.

Пример сценария задает режим вычисления вручную. Это означает, что книга будет пересчитывать формулы только в том случае, если сценарий подсказывает ему (или вручную вычисляется с помощью [пользовательского интерфейса).](https://support.microsoft.com/office/change-formula-recalculation-iteration-or-precision-in-excel-73fc7dac-91cf-4d36-86e8-67124f6bcce4) Затем сценарий отображает текущий режим вычислений и полностью пересчитывает всю книгу.

## <a name="sample-code-control-calculation-mode"></a>Пример кода: режим вычисления управления

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

## <a name="training-video-manage-calculation-mode"></a>Обучающее видео: управление режимом вычисления

[Смотреть Sudhi Ramamurthy ходить через этот пример на YouTube](https://youtu.be/iw6O8QH01CI).
