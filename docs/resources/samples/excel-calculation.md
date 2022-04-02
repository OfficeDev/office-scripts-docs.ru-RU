---
title: Управление режимом вычисления в Excel
description: Узнайте, как использовать Office скрипты для управления режимом вычисления в Excel в Интернете.
ms.date: 05/06/2021
ms.localizationpriority: medium
ms.openlocfilehash: fec88c904d95bfdab1514d44921f7fb1c6e9dd35
ms.sourcegitcommit: 7023b9e23499806901a5ecf8ebc460b76887cca6
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/31/2022
ms.locfileid: "64585515"
---
# <a name="manage-calculation-mode-in-excel"></a>Управление режимом вычисления в Excel

В этом примере показано, как использовать режим [вычисления](/javascript/api/office-scripts/excelscript/excelscript.calculationmode) и вычислять методы в Excel в Интернете с Office скриптами. Сценарий можно попробовать в любом Excel файле.

## <a name="scenario"></a>Сценарий

Перерасчет книг с большим количеством формул может занять некоторое время. Вместо того, чтобы Excel управлять при вычислениях, вы можете управлять ими в рамках сценария. Это поможет с производительностью в определенных сценариях.

Пример сценария задает режим вычисления вручную. Это означает, что книга будет пересчитывать формулы только тогда, когда сценарий сообщает ей (или вы вручную вычислять [с помощью пользовательского интерфейса](https://support.microsoft.com/office/73fc7dac-91cf-4d36-86e8-67124f6bcce4)). Затем сценарий отображает текущий режим вычислений и полностью пересчитывает всю книгу.

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

[Посмотрите, как суди Рамамурти (Sudhi Ramamurthy) пройдите этот пример на YouTube](https://youtu.be/iw6O8QH01CI).
