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
# <a name="manage-calculation-mode-in-excel"></a>Управление режимом вычисления в Excel

В этом примере показано, как использовать режим [вычисления](/javascript/api/office-scripts/excelscript/excelscript.calculationmode) и вычислять методы в Excel в Интернете с помощью Office Scripts. Сценарий можно попробовать в любом файле Excel.

## <a name="scenario"></a>Сценарий

В Excel в Интернете режим вычисления файла можно управлять программным образом с помощью API. Следующие действия возможны с помощью скриптов Office.

1. Получите режим вычисления.
1. Установите режим вычисления.
1. Вычислять формулы Excel для файлов, задамых в ручном режиме (также именуемого перерасчетом).

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

[![Просмотр пошагового видео об управлении режимом вычислений в Excel в Интернете](../../images/calc-mode-vid.jpg)](https://youtu.be/iw6O8QH01CI "Пошаговая видеокадры об управлении режимом вычислений в Excel в Интернете")
