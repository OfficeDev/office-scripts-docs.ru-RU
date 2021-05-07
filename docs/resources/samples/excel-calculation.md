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
# <a name="manage-calculation-mode-in-excel"></a>Управление режимом вычисления в Excel

В этом примере показано, как использовать режим [вычисления](/javascript/api/office-scripts/excelscript/excelscript.calculationmode) и вычислять методы Excel в Интернете с помощью Office скриптов. Сценарий можно попробовать в любом Excel файле.

## <a name="scenario"></a>Сценарий

В Excel в Интернете режим вычисления файла можно управлять программным образом с помощью API. Следующие действия возможны с помощью Office скриптов.

1. Получите режим вычисления.
1. Установите режим вычисления.
1. Вычислять Excel для файлов, задамых в ручном режиме (также именуемого перерасчетом).

## <a name="sample-code-control-calculation-mode"></a>Пример кода: режим вычисления управления

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

## <a name="training-video-manage-calculation-mode"></a>Обучающее видео: управление режимом вычисления

[Смотреть Sudhi Ramamurthy ходить через этот пример на YouTube](https://youtu.be/iw6O8QH01CI).
