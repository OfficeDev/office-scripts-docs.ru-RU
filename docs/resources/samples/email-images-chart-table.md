---
title: Отправить по электронной почте изображения Excel и таблицы
description: Узнайте, как использовать Office скрипты и Power Automate для извлечения и отправки по электронной почте изображений Excel диаграммы и таблицы.
ms.date: 04/05/2021
localization_priority: Normal
ms.openlocfilehash: 0265250f7fd885cb4899d0b9493b4285496965ff
ms.sourcegitcommit: 1f003c9924e651600c913d84094506125f1055ab
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 04/26/2021
ms.locfileid: "52026871"
---
# <a name="use-office-scripts-and-power-automate-to-email-images-of-a-chart-and-table"></a><span data-ttu-id="28929-103">Использование Office и Power Automate для отправки изображений диаграммы и таблицы по электронной почте</span><span class="sxs-lookup"><span data-stu-id="28929-103">Use Office Scripts and Power Automate to email images of a chart and table</span></span>

<span data-ttu-id="28929-104">В этом примере Office скрипты и Power Automate для создания диаграммы.</span><span class="sxs-lookup"><span data-stu-id="28929-104">This sample uses Office Scripts and Power Automate to create a chart.</span></span> <span data-ttu-id="28929-105">Затем он передает по электронной почте изображения диаграммы и базовой таблицы.</span><span class="sxs-lookup"><span data-stu-id="28929-105">It then emails images of the chart and its base table.</span></span>

## <a name="example-scenario"></a><span data-ttu-id="28929-106">Пример сценария</span><span class="sxs-lookup"><span data-stu-id="28929-106">Example scenario</span></span>

* <span data-ttu-id="28929-107">Вычислять, чтобы получить последние результаты.</span><span class="sxs-lookup"><span data-stu-id="28929-107">Calculate to get latest results.</span></span>
* <span data-ttu-id="28929-108">Создание диаграммы.</span><span class="sxs-lookup"><span data-stu-id="28929-108">Create chart.</span></span>
* <span data-ttu-id="28929-109">Получите изображения диаграммы и таблицы.</span><span class="sxs-lookup"><span data-stu-id="28929-109">Get chart and table images.</span></span>
* <span data-ttu-id="28929-110">Отправьте изображения по электронной почте Power Automate.</span><span class="sxs-lookup"><span data-stu-id="28929-110">Email the images with Power Automate.</span></span>

<span data-ttu-id="28929-111">_Входные данные_</span><span class="sxs-lookup"><span data-stu-id="28929-111">_Input data_</span></span>

:::image type="content" source="../../images/input-data.png" alt-text="Таблица, показывающая таблицу входных данных.":::

<span data-ttu-id="28929-113">_Диаграмма вывода_</span><span class="sxs-lookup"><span data-stu-id="28929-113">_Output chart_</span></span>

:::image type="content" source="../../images/chart-created.png" alt-text="Диаграмма столбцов, созданная с указанием суммы, которая должна быть засвеяна клиентом.":::

<span data-ttu-id="28929-115">_Электронная почта, полученная Power Automate потока_</span><span class="sxs-lookup"><span data-stu-id="28929-115">_Email that was received through Power Automate flow_</span></span>

:::image type="content" source="../../images/email-received.png" alt-text="Сообщение, отправленное потоком, с указанием Excel, встроенного в тело.":::

## <a name="solution"></a><span data-ttu-id="28929-117">Решение</span><span class="sxs-lookup"><span data-stu-id="28929-117">Solution</span></span>

<span data-ttu-id="28929-118">Это решение состоит из двух частей:</span><span class="sxs-lookup"><span data-stu-id="28929-118">This solution has two parts:</span></span>

1. [<span data-ttu-id="28929-119">Сценарий Office для вычисления и извлечения Excel диаграммы и таблицы</span><span class="sxs-lookup"><span data-stu-id="28929-119">An Office Script to calculate and extract Excel chart and table</span></span>](#sample-code-calculate-and-extract-excel-chart-and-table)
1. <span data-ttu-id="28929-120">Поток Power Automate для вызова скрипта и отправки результатов по электронной почте.</span><span class="sxs-lookup"><span data-stu-id="28929-120">A Power Automate flow to invoke the script and email the results.</span></span> <span data-ttu-id="28929-121">Пример этого см. в примере [Create an automated workflow with Power Automate.](../../tutorials/excel-power-automate-returns.md#create-an-automated-workflow-with-power-automate)</span><span class="sxs-lookup"><span data-stu-id="28929-121">For an example on how to do this, see [Create an automated workflow with Power Automate](../../tutorials/excel-power-automate-returns.md#create-an-automated-workflow-with-power-automate).</span></span>

## <a name="sample-code-calculate-and-extract-excel-chart-and-table"></a><span data-ttu-id="28929-122">Пример кода. Вычислять и извлекать Excel диаграмму и таблицу</span><span class="sxs-lookup"><span data-stu-id="28929-122">Sample code: Calculate and extract Excel chart and table</span></span>

<span data-ttu-id="28929-123">Следующий сценарий вычисляет и извлекает Excel диаграмму и таблицу.</span><span class="sxs-lookup"><span data-stu-id="28929-123">The following script calculates and extracts an Excel chart and table.</span></span>

<span data-ttu-id="28929-124">Скачайте пример файла <a href="email-chart-table.xlsx">email-chart-table.xlsx</a> и используйте его с помощью этого скрипта, чтобы попробовать его самостоятельно!</span><span class="sxs-lookup"><span data-stu-id="28929-124">Download the sample file <a href="email-chart-table.xlsx">email-chart-table.xlsx</a> and use it with this script to try it out yourself!</span></span>

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

## <a name="power-automate-flow-email-the-chart-and-table-images"></a><span data-ttu-id="28929-125">Power Automate потока: отправить по электронной почте изображения диаграммы и таблицы</span><span class="sxs-lookup"><span data-stu-id="28929-125">Power Automate flow: Email the chart and table images</span></span>

<span data-ttu-id="28929-126">Этот поток запускает сценарий и передает возвращаемые изображения по электронной почте.</span><span class="sxs-lookup"><span data-stu-id="28929-126">This flow runs the script and emails the returned images.</span></span>

1. <span data-ttu-id="28929-127">Создайте новый **поток мгновенных облаков.**</span><span class="sxs-lookup"><span data-stu-id="28929-127">Create a new **Instant cloud flow**.</span></span>
1. <span data-ttu-id="28929-128">Выберите **вручную вызвать поток и** нажмите **кнопку Создать**.</span><span class="sxs-lookup"><span data-stu-id="28929-128">Select **Manually trigger a flow** and press **Create**.</span></span>
1. <span data-ttu-id="28929-129">Добавьте новый **шаг,** использующий **соединителю Excel Online (Бизнес)** с действием **Запуска скрипта (предварительного просмотра).**</span><span class="sxs-lookup"><span data-stu-id="28929-129">Add a **New step** that uses the **Excel Online (Business)** connector with the **Run script (preview)** action.</span></span> <span data-ttu-id="28929-130">Используйте следующие значения для действия:</span><span class="sxs-lookup"><span data-stu-id="28929-130">Use the following values for the action:</span></span>
    * <span data-ttu-id="28929-131">**Расположение**: OneDrive для бизнеса</span><span class="sxs-lookup"><span data-stu-id="28929-131">**Location**: OneDrive for Business</span></span>
    * <span data-ttu-id="28929-132">**Библиотека документов**: OneDrive</span><span class="sxs-lookup"><span data-stu-id="28929-132">**Document Library**: OneDrive</span></span>
    * <span data-ttu-id="28929-133">**Файл**: Ваша книга [(выбрана с помощью выбора файла)](../../testing/power-automate-troubleshooting.md#select-workbooks-with-the-file-browser-control)</span><span class="sxs-lookup"><span data-stu-id="28929-133">**File**: Your workbook ([selected with the file chooser](../../testing/power-automate-troubleshooting.md#select-workbooks-with-the-file-browser-control))</span></span>
    * <span data-ttu-id="28929-134">**Сценарий:** имя сценария</span><span class="sxs-lookup"><span data-stu-id="28929-134">**Script**: Your script name</span></span>

    :::image type="content" source="../../images/email-chart-sample-flow-1.png" alt-text="Завершенный соедините Excel Online (Бизнес) в Power Automate.":::
1. <span data-ttu-id="28929-136">В этом примере Outlook как клиент электронной почты.</span><span class="sxs-lookup"><span data-stu-id="28929-136">This sample uses Outlook as the email client.</span></span> <span data-ttu-id="28929-137">Можно использовать любые соединители электронной почты Power Automate поддерживает, но остальные действия предполагают, что вы выбрали Outlook.</span><span class="sxs-lookup"><span data-stu-id="28929-137">You could use any email connector Power Automate supports, but the rest of the steps assume that you chose Outlook.</span></span> <span data-ttu-id="28929-138">Добавьте новый **шаг,** использующий **соединителю Office 365 Outlook** и действие Отправка и электронная почта **(V2).**</span><span class="sxs-lookup"><span data-stu-id="28929-138">Add a **New step** that uses the **Office 365 Outlook** connector and the **Send and email (V2)** action.</span></span> <span data-ttu-id="28929-139">Используйте следующие значения для действия:</span><span class="sxs-lookup"><span data-stu-id="28929-139">Use the following values for the action:</span></span>
    * <span data-ttu-id="28929-140">**Чтобы:** ваша тестовая учетная запись электронной почты (или личная электронная почта)</span><span class="sxs-lookup"><span data-stu-id="28929-140">**To**: Your test email account (or personal email)</span></span>
    * <span data-ttu-id="28929-141">**Тема:** Просмотрите отчетные данные</span><span class="sxs-lookup"><span data-stu-id="28929-141">**Subject**: Please Review Report Data</span></span>
    * <span data-ttu-id="28929-142">Для поля **Body** выберите "Представление кода" () и `</>` введите следующее:</span><span class="sxs-lookup"><span data-stu-id="28929-142">For the **Body** field, select "Code View" (`</>`) and enter the following:</span></span>

    ```HTML
    <p>Please review the following report data:<br>
    <br>
    Chart:<br>
    <br>
    <img src="data:image/png;base64,@{outputs('Run_script')?['body/result/chartImage']}"/>
    <br>
    Data:<br>
    <br>
    <img src="data:image/png;base64,@{outputs('Run_script')?['body/result/tableImage']}"/>
    <br>
    </p>
    ```

    :::image type="content" source="../../images/email-chart-sample-flow-2.png" alt-text="Завершенный соедините Office 365 Outlook в Power Automate.":::
1. <span data-ttu-id="28929-144">Сохраните поток и попробуйте его.</span><span class="sxs-lookup"><span data-stu-id="28929-144">Save the flow and try it out.</span></span>

## <a name="training-video-extract-and-email-images-of-chart-and-table"></a><span data-ttu-id="28929-145">Обучающее видео: извлечение и отправка изображений диаграммы и таблицы по электронной почте</span><span class="sxs-lookup"><span data-stu-id="28929-145">Training video: Extract and email images of chart and table</span></span>

<span data-ttu-id="28929-146">[![Просмотрите пошаговую видеозапись по извлечению и отправке изображений диаграммы и таблицы по электронной почте](../../images/charts-image-vid.jpg)](https://youtu.be/152GJyqc-Kw "Пошаговая видеозапись по извлечению и отправке изображений диаграммы и таблицы по электронной почте")</span><span class="sxs-lookup"><span data-stu-id="28929-146">[![Watch step-by-step video on how to extract and email images of chart and table](../../images/charts-image-vid.jpg)](https://youtu.be/152GJyqc-Kw "Step-by-step video on how to extract and email images of chart and table")</span></span>
