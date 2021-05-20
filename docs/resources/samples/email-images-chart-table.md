---
title: Отправьте изображения диаграммы и таблицы Excel электронной почте
description: Узнайте, как использовать Office и Power Automate для извлечения и электронной почты изображения диаграммы Excel таблицы.
ms.date: 05/06/2021
localization_priority: Normal
ms.openlocfilehash: 54b6b67a0f211f2dc6c881bab17ff23220619e6e
ms.sourcegitcommit: 4687693f02fc90a57ba30c461f35046e02e6f5fb
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 05/19/2021
ms.locfileid: "52545780"
---
# <a name="use-office-scripts-and-power-automate-to-email-images-of-a-chart-and-table"></a><span data-ttu-id="c62fd-103">Используйте Office и Power Automate по электронной почте изображения диаграммы и таблицы</span><span class="sxs-lookup"><span data-stu-id="c62fd-103">Use Office Scripts and Power Automate to email images of a chart and table</span></span>

<span data-ttu-id="c62fd-104">Этот пример использует Office скрипты и Power Automate для создания диаграммы.</span><span class="sxs-lookup"><span data-stu-id="c62fd-104">This sample uses Office Scripts and Power Automate to create a chart.</span></span> <span data-ttu-id="c62fd-105">Затем он отправляют изображения диаграммы и ее базовой таблицы.</span><span class="sxs-lookup"><span data-stu-id="c62fd-105">It then emails images of the chart and its base table.</span></span>

## <a name="example-scenario"></a><span data-ttu-id="c62fd-106">Пример сценария</span><span class="sxs-lookup"><span data-stu-id="c62fd-106">Example scenario</span></span>

* <span data-ttu-id="c62fd-107">Рассчитайте, чтобы получить последние результаты.</span><span class="sxs-lookup"><span data-stu-id="c62fd-107">Calculate to get latest results.</span></span>
* <span data-ttu-id="c62fd-108">Создайте диаграмму.</span><span class="sxs-lookup"><span data-stu-id="c62fd-108">Create chart.</span></span>
* <span data-ttu-id="c62fd-109">Получите изображения диаграммы и таблицы.</span><span class="sxs-lookup"><span data-stu-id="c62fd-109">Get chart and table images.</span></span>
* <span data-ttu-id="c62fd-110">Отправьте изображения по электронной почте Power Automate.</span><span class="sxs-lookup"><span data-stu-id="c62fd-110">Email the images with Power Automate.</span></span>

<span data-ttu-id="c62fd-111">_Входные данные_</span><span class="sxs-lookup"><span data-stu-id="c62fd-111">_Input data_</span></span>

:::image type="content" source="../../images/input-data.png" alt-text="Лист с указанием таблицы входных данных":::

<span data-ttu-id="c62fd-113">_Диаграмма вывода_</span><span class="sxs-lookup"><span data-stu-id="c62fd-113">_Output chart_</span></span>

:::image type="content" source="../../images/chart-created.png" alt-text="Диаграмма столбца создана, показывающая сумму, примеская заказчиком":::

<span data-ttu-id="c62fd-115">_Электронная почта, полученная через Power Automate поток_</span><span class="sxs-lookup"><span data-stu-id="c62fd-115">_Email that was received through Power Automate flow_</span></span>

:::image type="content" source="../../images/email-received.png" alt-text="Письмо, отправленное потоком, показывающим Excel диаграмму, встроенную в тело":::

## <a name="solution"></a><span data-ttu-id="c62fd-117">Решение</span><span class="sxs-lookup"><span data-stu-id="c62fd-117">Solution</span></span>

<span data-ttu-id="c62fd-118">Это решение состоит из двух частей:</span><span class="sxs-lookup"><span data-stu-id="c62fd-118">This solution has two parts:</span></span>

1. [<span data-ttu-id="c62fd-119">Сценарий Office для расчета и извлечения Excel диаграммы и таблицы</span><span class="sxs-lookup"><span data-stu-id="c62fd-119">An Office Script to calculate and extract Excel chart and table</span></span>](#sample-code-calculate-and-extract-excel-chart-and-table)
1. <span data-ttu-id="c62fd-120">Поток Power Automate для вызова скрипта и электронной почты результатов.</span><span class="sxs-lookup"><span data-stu-id="c62fd-120">A Power Automate flow to invoke the script and email the results.</span></span> <span data-ttu-id="c62fd-121">Пример того, как это сделать, смотрите Создание [автоматизированного рабочего процесса с Power Automate.](../../tutorials/excel-power-automate-returns.md#create-an-automated-workflow-with-power-automate)</span><span class="sxs-lookup"><span data-stu-id="c62fd-121">For an example on how to do this, see [Create an automated workflow with Power Automate](../../tutorials/excel-power-automate-returns.md#create-an-automated-workflow-with-power-automate).</span></span>

## <a name="sample-code-calculate-and-extract-excel-chart-and-table"></a><span data-ttu-id="c62fd-122">Пример кода: Рассчитайте и извлекайте Excel диаграмму и таблицу</span><span class="sxs-lookup"><span data-stu-id="c62fd-122">Sample code: Calculate and extract Excel chart and table</span></span>

<span data-ttu-id="c62fd-123">Следующий скрипт вычисляет и извлекает Excel и таблицу.</span><span class="sxs-lookup"><span data-stu-id="c62fd-123">The following script calculates and extracts an Excel chart and table.</span></span>

<span data-ttu-id="c62fd-124">Скачать пример файла <a href="email-chart-table.xlsx">email-chart-table.xlsx</a> использовать его с этим скриптом, чтобы попробовать его самостоятельно!</span><span class="sxs-lookup"><span data-stu-id="c62fd-124">Download the sample file <a href="email-chart-table.xlsx">email-chart-table.xlsx</a> and use it with this script to try it out yourself!</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook): ReportImages {
  // Recalculate the workbook to ensure all tables and charts are updated.
  workbook.getApplication().calculate(ExcelScript.CalculationType.full);
  
  // Get the data from the "InvoiceAmounts" table.
  let sheet1 = workbook.getWorksheet("Sheet1");
  const table = workbook.getWorksheet('InvoiceAmounts').getTables()[0];
  const rows = table.getRange().getTexts();

  // Get only the "Customer Name" and "Amount due" columns, then remove the "Total" row.
  const selectColumns = rows.map((row) => {
    return [row[2], row[5]];
  });
  table.setShowTotals(true);
  selectColumns.splice(selectColumns.length-1, 1);
  console.log(selectColumns);

  // Delete the "ChartSheet" worksheet if it's present, then recreate it.
  workbook.getWorksheet('ChartSheet')?.delete();
  const chartSheet = workbook.addWorksheet('ChartSheet');

  // Add the selected data to the new worksheet.
  const targetRange = chartSheet.getRange('A1').getResizedRange(selectColumns.length-1, selectColumns[0].length-1);
  targetRange.setValues(selectColumns);

  // Insert the chart on sheet 'ChartSheet' at cell "D1".
  let chart_2 = chartSheet.addChart(ExcelScript.ChartType.columnClustered, targetRange);
  chart_2.setPosition('D1');

  // Get images of the chart and table, then return them for a Power Automate flow.
  const chartImage = chart_2.getImage();
  const tableImage = table.getRange().getImage();
  return {chartImage, tableImage};
}

// The interface for table and chart images.
interface ReportImages {
  chartImage: string
  tableImage: string
}
```

## <a name="power-automate-flow-email-the-chart-and-table-images"></a><span data-ttu-id="c62fd-125">Power Automate поток: Электронная почта диаграммы и таблицы изображения</span><span class="sxs-lookup"><span data-stu-id="c62fd-125">Power Automate flow: Email the chart and table images</span></span>

<span data-ttu-id="c62fd-126">Этот поток запускает скрипт и отправляют возвращенные изображения по электронной почте.</span><span class="sxs-lookup"><span data-stu-id="c62fd-126">This flow runs the script and emails the returned images.</span></span>

1. <span data-ttu-id="c62fd-127">Создайте новый **мгновенный поток облаков.**</span><span class="sxs-lookup"><span data-stu-id="c62fd-127">Create a new **Instant cloud flow**.</span></span>
1. <span data-ttu-id="c62fd-128">Выберите **Вручную вызвать поток и** нажмите **Создать**.</span><span class="sxs-lookup"><span data-stu-id="c62fd-128">Select **Manually trigger a flow** and press **Create**.</span></span>
1. <span data-ttu-id="c62fd-129">Добавьте **новый шаг,** который использует **Excel Online (Бизнес)** с **действием сценария** Run.</span><span class="sxs-lookup"><span data-stu-id="c62fd-129">Add a **New step** that uses the **Excel Online (Business)** connector with the **Run script** action.</span></span> <span data-ttu-id="c62fd-130">Используйте следующие значения для действия:</span><span class="sxs-lookup"><span data-stu-id="c62fd-130">Use the following values for the action:</span></span>
    * <span data-ttu-id="c62fd-131">**Расположение**: OneDrive для бизнеса</span><span class="sxs-lookup"><span data-stu-id="c62fd-131">**Location**: OneDrive for Business</span></span>
    * <span data-ttu-id="c62fd-132">**Библиотека документов**: OneDrive</span><span class="sxs-lookup"><span data-stu-id="c62fd-132">**Document Library**: OneDrive</span></span>
    * <span data-ttu-id="c62fd-133">**Файл**: Ваша трудовая [книжка (выбрана с помощью выбранного файла)](../../testing/power-automate-troubleshooting.md#select-workbooks-with-the-file-browser-control)</span><span class="sxs-lookup"><span data-stu-id="c62fd-133">**File**: Your workbook ([selected with the file chooser](../../testing/power-automate-troubleshooting.md#select-workbooks-with-the-file-browser-control))</span></span>
    * <span data-ttu-id="c62fd-134">**Сценарий**: Ваше имя скрипта</span><span class="sxs-lookup"><span data-stu-id="c62fd-134">**Script**: Your script name</span></span>

    :::image type="content" source="../../images/email-chart-sample-flow-1.png" alt-text="Завершенный разъем Excel Online (Бизнес) в Power Automate":::
1. <span data-ttu-id="c62fd-136">Этот образец использует Outlook в качестве почтового клиента.</span><span class="sxs-lookup"><span data-stu-id="c62fd-136">This sample uses Outlook as the email client.</span></span> <span data-ttu-id="c62fd-137">Вы можете использовать любой разъем электронной Power Automate поддерживает, но остальные шаги предполагают, что вы выбрали Outlook.</span><span class="sxs-lookup"><span data-stu-id="c62fd-137">You could use any email connector Power Automate supports, but the rest of the steps assume that you chose Outlook.</span></span> <span data-ttu-id="c62fd-138">Добавьте **новый шаг,** который использует **Office 365 Outlook** и отправить и **отправить (V2)** действий.</span><span class="sxs-lookup"><span data-stu-id="c62fd-138">Add a **New step** that uses the **Office 365 Outlook** connector and the **Send and email (V2)** action.</span></span> <span data-ttu-id="c62fd-139">Используйте следующие значения для действия:</span><span class="sxs-lookup"><span data-stu-id="c62fd-139">Use the following values for the action:</span></span>
    * <span data-ttu-id="c62fd-140">**Для:** Ваш тестовый адрес электронной почты (или личная электронная почта)</span><span class="sxs-lookup"><span data-stu-id="c62fd-140">**To**: Your test email account (or personal email)</span></span>
    * <span data-ttu-id="c62fd-141">**Тема**: Пожалуйста, просмотрите данные отчета</span><span class="sxs-lookup"><span data-stu-id="c62fd-141">**Subject**: Please Review Report Data</span></span>
    * <span data-ttu-id="c62fd-142">Для поля **тела** выберите "Code View" `</>` () и введите следующее:</span><span class="sxs-lookup"><span data-stu-id="c62fd-142">For the **Body** field, select "Code View" (`</>`) and enter the following:</span></span>

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

    :::image type="content" source="../../images/email-chart-sample-flow-2.png" alt-text="Завершенный Office 365 Outlook в Power Automate":::
1. <span data-ttu-id="c62fd-144">Сохранить поток и попробовать его.</span><span class="sxs-lookup"><span data-stu-id="c62fd-144">Save the flow and try it out.</span></span>

## <a name="training-video-extract-and-email-images-of-chart-and-table"></a><span data-ttu-id="c62fd-145">Учебное видео: Выдержка и изображения электронной почты диаграммы и таблицы</span><span class="sxs-lookup"><span data-stu-id="c62fd-145">Training video: Extract and email images of chart and table</span></span>

<span data-ttu-id="c62fd-146">[Смотреть Судхи Рамамурти ходить через этот образец на YouTube](https://youtu.be/152GJyqc-Kw).</span><span class="sxs-lookup"><span data-stu-id="c62fd-146">[Watch Sudhi Ramamurthy walk through this sample on YouTube](https://youtu.be/152GJyqc-Kw).</span></span>
