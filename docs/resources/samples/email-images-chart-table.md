---
title: Отправить по электронной почте изображения диаграммы и таблицы Excel
description: Узнайте, как использовать office Scripts и Power Automate для извлечения и отправки по электронной почте изображений диаграммы и таблицы Excel.
ms.date: 04/01/2021
localization_priority: Normal
ms.openlocfilehash: de3cf16537cb12db45d4d465d367d797d053afc4
ms.sourcegitcommit: 45ffe3dbd2c834b78592ad35928cf8096f5e80bc
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 04/14/2021
ms.locfileid: "51754812"
---
# <a name="use-office-scripts-and-power-automate-to-email-images-of-a-chart-and-table"></a><span data-ttu-id="abd03-103">Использование скриптов Office и power Automate для отправки изображений диаграммы и таблицы по электронной почте</span><span class="sxs-lookup"><span data-stu-id="abd03-103">Use Office Scripts and Power Automate to email images of a chart and table</span></span>

<span data-ttu-id="abd03-104">В этом примере для создания диаграммы используются скрипты Office и Power Automate.</span><span class="sxs-lookup"><span data-stu-id="abd03-104">This sample uses Office Scripts and Power Automate to create a chart.</span></span> <span data-ttu-id="abd03-105">Затем он передает по электронной почте изображения диаграммы и базовой таблицы.</span><span class="sxs-lookup"><span data-stu-id="abd03-105">It then emails images of the chart and its base table.</span></span>

## <a name="example-scenario"></a><span data-ttu-id="abd03-106">Пример сценария</span><span class="sxs-lookup"><span data-stu-id="abd03-106">Example scenario</span></span>

* <span data-ttu-id="abd03-107">Вычислять, чтобы получить последние результаты.</span><span class="sxs-lookup"><span data-stu-id="abd03-107">Calculate to get latest results.</span></span>
* <span data-ttu-id="abd03-108">Создание диаграммы.</span><span class="sxs-lookup"><span data-stu-id="abd03-108">Create chart.</span></span>
* <span data-ttu-id="abd03-109">Получите изображения диаграммы и таблицы.</span><span class="sxs-lookup"><span data-stu-id="abd03-109">Get chart and table images.</span></span>
* <span data-ttu-id="abd03-110">Отправьте изображения по электронной почте с помощью Power Automate.</span><span class="sxs-lookup"><span data-stu-id="abd03-110">Email the images with Power Automate.</span></span>

<span data-ttu-id="abd03-111">_Входные данные_</span><span class="sxs-lookup"><span data-stu-id="abd03-111">_Input data_</span></span>

:::image type="content" source="../../images/input-data.png" alt-text="Таблица, показывающая таблицу входных данных.":::

<span data-ttu-id="abd03-113">_Диаграмма вывода_</span><span class="sxs-lookup"><span data-stu-id="abd03-113">_Output chart_</span></span>

:::image type="content" source="../../images/chart-created.png" alt-text="Диаграмма столбцов, созданная с указанием суммы, которая должна быть засвеяна клиентом.":::

<span data-ttu-id="abd03-115">_Электронная почта, полученная через поток Power Automate_</span><span class="sxs-lookup"><span data-stu-id="abd03-115">_Email that was received through Power Automate flow_</span></span>

:::image type="content" source="../../images/email-received.png" alt-text="Сообщение, отправленное потоком, с указанием диаграммы Excel, встроенной в тело.":::

## <a name="solution"></a><span data-ttu-id="abd03-117">Решение</span><span class="sxs-lookup"><span data-stu-id="abd03-117">Solution</span></span>

<span data-ttu-id="abd03-118">Это решение состоит из двух частей:</span><span class="sxs-lookup"><span data-stu-id="abd03-118">This solution has two parts:</span></span>

1. [<span data-ttu-id="abd03-119">Сценарий Office для вычисления и извлечения диаграммы и таблицы Excel</span><span class="sxs-lookup"><span data-stu-id="abd03-119">An Office Script to calculate and extract Excel chart and table</span></span>](#sample-code-calculate-and-extract-excel-chart-and-table)
1. <span data-ttu-id="abd03-120">Поток Power Automate для вызова скрипта и отправки результатов по электронной почте.</span><span class="sxs-lookup"><span data-stu-id="abd03-120">A Power Automate flow to invoke the script and email the results.</span></span> <span data-ttu-id="abd03-121">Пример этого см. в примере [Create an automated workflow with Power Automate.](../../tutorials/excel-power-automate-returns.md#create-an-automated-workflow-with-power-automate)</span><span class="sxs-lookup"><span data-stu-id="abd03-121">For an example on how to do this, see [Create an automated workflow with Power Automate](../../tutorials/excel-power-automate-returns.md#create-an-automated-workflow-with-power-automate).</span></span>

## <a name="sample-code-calculate-and-extract-excel-chart-and-table"></a><span data-ttu-id="abd03-122">Пример кода: Вычислять и извлекать диаграмму и таблицу Excel</span><span class="sxs-lookup"><span data-stu-id="abd03-122">Sample code: Calculate and extract Excel chart and table</span></span>

<span data-ttu-id="abd03-123">Следующий сценарий вычисляет и извлекает диаграмму и таблицу Excel.</span><span class="sxs-lookup"><span data-stu-id="abd03-123">The following script calculates and extracts an Excel chart and table.</span></span>

<span data-ttu-id="abd03-124">Скачайте пример файла <a href="email-chart-table.xlsx">email-chart-table.xlsx</a> и используйте его с помощью этого скрипта, чтобы попробовать его самостоятельно!</span><span class="sxs-lookup"><span data-stu-id="abd03-124">Download the sample file <a href="email-chart-table.xlsx">email-chart-table.xlsx</a> and use it with this script to try it out yourself!</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook): ReportImages {

  workbook.getApplication().calculate(ExcelScript.CalculationType.full);
  
  let sheet1 = workbook.getWorksheet("Sheet1");
  const table = workbook.getWorksheet('InvoiceAmounts').getTables()[0];
  const rows = table.getRange().getTexts();

  const selectColumns = rows.map((row) => {
    return [row[2], row[5]];
  });
  table.setShowTotals(true);
  selectColumns.splice(selectColumns.length-1, 1);
  console.log(selectColumns);

  workbook.getWorksheet('ChartSheet')?.delete();
  const chartSheet = workbook.addWorksheet('ChartSheet');
  const targetRange = updateRange(chartSheet, selectColumns);

  // Insert chart on sheet 'Sheet1'.
  let chart_2 = chartSheet.addChart(ExcelScript.ChartType.columnClustered, targetRange);
  chart_2.setPosition('D1');
  const chartImage = chart_2.getImage();
  const tableImage = table.getRange().getImage();
  return {
    chartImage,
    tableImage
  }
}

function updateRange(sheet: ExcelScript.Worksheet, data: string[][]): ExcelScript.Range {
  const targetRange = sheet.getRange('A1').getResizedRange(data.length-1, data[0].length-1);
  targetRange.setValues(data);
  return targetRange;
}

interface ReportImages {
  chartImage: string
  tableImage: string
}
```

## <a name="training-video-extract-and-email-images-of-chart-and-table"></a><span data-ttu-id="abd03-125">Обучающее видео: извлечение и отправка изображений диаграммы и таблицы по электронной почте</span><span class="sxs-lookup"><span data-stu-id="abd03-125">Training video: Extract and email images of chart and table</span></span>

<span data-ttu-id="abd03-126">[![Просмотрите пошаговую видеозапись по извлечению и отправке изображений диаграммы и таблицы по электронной почте](../../images/charts-image-vid.jpg)](https://youtu.be/152GJyqc-Kw "Пошаговая видеозапись по извлечению и отправке изображений диаграммы и таблицы по электронной почте")</span><span class="sxs-lookup"><span data-stu-id="abd03-126">[![Watch step-by-step video on how to extract and email images of chart and table](../../images/charts-image-vid.jpg)](https://youtu.be/152GJyqc-Kw "Step-by-step video on how to extract and email images of chart and table")</span></span>
