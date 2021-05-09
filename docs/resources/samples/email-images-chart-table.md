---
title: Отправить по электронной почте изображения Excel и таблицы
description: Узнайте, как использовать Office скрипты и Power Automate для извлечения и отправки по электронной почте изображений Excel диаграммы и таблицы.
ms.date: 05/06/2021
localization_priority: Normal
ms.openlocfilehash: f8b52cbf8c19b93c5fc4288fe97775a25e922ab9
ms.sourcegitcommit: 763d341857bcb209b2f2c278a82fdb63d0e18f0a
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 05/08/2021
ms.locfileid: "52285859"
---
# <a name="use-office-scripts-and-power-automate-to-email-images-of-a-chart-and-table"></a><span data-ttu-id="b5fb0-103">Использование Office и Power Automate для отправки изображений диаграммы и таблицы по электронной почте</span><span class="sxs-lookup"><span data-stu-id="b5fb0-103">Use Office Scripts and Power Automate to email images of a chart and table</span></span>

<span data-ttu-id="b5fb0-104">В этом примере Office скрипты и Power Automate для создания диаграммы.</span><span class="sxs-lookup"><span data-stu-id="b5fb0-104">This sample uses Office Scripts and Power Automate to create a chart.</span></span> <span data-ttu-id="b5fb0-105">Затем он передает по электронной почте изображения диаграммы и базовой таблицы.</span><span class="sxs-lookup"><span data-stu-id="b5fb0-105">It then emails images of the chart and its base table.</span></span>

## <a name="example-scenario"></a><span data-ttu-id="b5fb0-106">Пример сценария</span><span class="sxs-lookup"><span data-stu-id="b5fb0-106">Example scenario</span></span>

* <span data-ttu-id="b5fb0-107">Вычислять, чтобы получить последние результаты.</span><span class="sxs-lookup"><span data-stu-id="b5fb0-107">Calculate to get latest results.</span></span>
* <span data-ttu-id="b5fb0-108">Создание диаграммы.</span><span class="sxs-lookup"><span data-stu-id="b5fb0-108">Create chart.</span></span>
* <span data-ttu-id="b5fb0-109">Получите изображения диаграммы и таблицы.</span><span class="sxs-lookup"><span data-stu-id="b5fb0-109">Get chart and table images.</span></span>
* <span data-ttu-id="b5fb0-110">Отправьте изображения по электронной почте Power Automate.</span><span class="sxs-lookup"><span data-stu-id="b5fb0-110">Email the images with Power Automate.</span></span>

<span data-ttu-id="b5fb0-111">_Входные данные_</span><span class="sxs-lookup"><span data-stu-id="b5fb0-111">_Input data_</span></span>

:::image type="content" source="../../images/input-data.png" alt-text="Таблица, показывающая таблицу входных данных":::

<span data-ttu-id="b5fb0-113">_Диаграмма вывода_</span><span class="sxs-lookup"><span data-stu-id="b5fb0-113">_Output chart_</span></span>

:::image type="content" source="../../images/chart-created.png" alt-text="Диаграмма столбцов, созданная с указанием суммы, за которую должен высмеять клиент":::

<span data-ttu-id="b5fb0-115">_Электронная почта, полученная Power Automate потока_</span><span class="sxs-lookup"><span data-stu-id="b5fb0-115">_Email that was received through Power Automate flow_</span></span>

:::image type="content" source="../../images/email-received.png" alt-text="Сообщение, отправленное потоком, с указанием Excel, встроенного в тело":::

## <a name="solution"></a><span data-ttu-id="b5fb0-117">Решение</span><span class="sxs-lookup"><span data-stu-id="b5fb0-117">Solution</span></span>

<span data-ttu-id="b5fb0-118">Это решение состоит из двух частей:</span><span class="sxs-lookup"><span data-stu-id="b5fb0-118">This solution has two parts:</span></span>

1. [<span data-ttu-id="b5fb0-119">Сценарий Office для вычисления и извлечения Excel диаграммы и таблицы</span><span class="sxs-lookup"><span data-stu-id="b5fb0-119">An Office Script to calculate and extract Excel chart and table</span></span>](#sample-code-calculate-and-extract-excel-chart-and-table)
1. <span data-ttu-id="b5fb0-120">Поток Power Automate для вызова скрипта и отправки результатов по электронной почте.</span><span class="sxs-lookup"><span data-stu-id="b5fb0-120">A Power Automate flow to invoke the script and email the results.</span></span> <span data-ttu-id="b5fb0-121">Пример этого см. в примере [Create an automated workflow with Power Automate.](../../tutorials/excel-power-automate-returns.md#create-an-automated-workflow-with-power-automate)</span><span class="sxs-lookup"><span data-stu-id="b5fb0-121">For an example on how to do this, see [Create an automated workflow with Power Automate](../../tutorials/excel-power-automate-returns.md#create-an-automated-workflow-with-power-automate).</span></span>

## <a name="sample-code-calculate-and-extract-excel-chart-and-table"></a><span data-ttu-id="b5fb0-122">Пример кода. Вычислять и извлекать Excel диаграмму и таблицу</span><span class="sxs-lookup"><span data-stu-id="b5fb0-122">Sample code: Calculate and extract Excel chart and table</span></span>

<span data-ttu-id="b5fb0-123">Следующий сценарий вычисляет и извлекает Excel диаграмму и таблицу.</span><span class="sxs-lookup"><span data-stu-id="b5fb0-123">The following script calculates and extracts an Excel chart and table.</span></span>

<span data-ttu-id="b5fb0-124">Скачайте пример файла <a href="email-chart-table.xlsx">email-chart-table.xlsx</a> и используйте его с помощью этого скрипта, чтобы попробовать его самостоятельно!</span><span class="sxs-lookup"><span data-stu-id="b5fb0-124">Download the sample file <a href="email-chart-table.xlsx">email-chart-table.xlsx</a> and use it with this script to try it out yourself!</span></span>

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

## <a name="power-automate-flow-email-the-chart-and-table-images"></a><span data-ttu-id="b5fb0-125">Power Automate потока: отправить по электронной почте изображения диаграммы и таблицы</span><span class="sxs-lookup"><span data-stu-id="b5fb0-125">Power Automate flow: Email the chart and table images</span></span>

<span data-ttu-id="b5fb0-126">Этот поток запускает сценарий и передает возвращаемые изображения по электронной почте.</span><span class="sxs-lookup"><span data-stu-id="b5fb0-126">This flow runs the script and emails the returned images.</span></span>

1. <span data-ttu-id="b5fb0-127">Создайте новый **поток мгновенных облаков.**</span><span class="sxs-lookup"><span data-stu-id="b5fb0-127">Create a new **Instant cloud flow**.</span></span>
1. <span data-ttu-id="b5fb0-128">Выберите **вручную вызвать поток и** нажмите **кнопку Создать**.</span><span class="sxs-lookup"><span data-stu-id="b5fb0-128">Select **Manually trigger a flow** and press **Create**.</span></span>
1. <span data-ttu-id="b5fb0-129">Добавьте новый **шаг,** использующий **соединителю Excel Online (Бизнес)** с действием **Запуска скрипта (предварительного просмотра).**</span><span class="sxs-lookup"><span data-stu-id="b5fb0-129">Add a **New step** that uses the **Excel Online (Business)** connector with the **Run script (preview)** action.</span></span> <span data-ttu-id="b5fb0-130">Используйте следующие значения для действия:</span><span class="sxs-lookup"><span data-stu-id="b5fb0-130">Use the following values for the action:</span></span>
    * <span data-ttu-id="b5fb0-131">**Расположение**: OneDrive для бизнеса</span><span class="sxs-lookup"><span data-stu-id="b5fb0-131">**Location**: OneDrive for Business</span></span>
    * <span data-ttu-id="b5fb0-132">**Библиотека документов**: OneDrive</span><span class="sxs-lookup"><span data-stu-id="b5fb0-132">**Document Library**: OneDrive</span></span>
    * <span data-ttu-id="b5fb0-133">**Файл**: Ваша книга [(выбрана с помощью выбора файла)](../../testing/power-automate-troubleshooting.md#select-workbooks-with-the-file-browser-control)</span><span class="sxs-lookup"><span data-stu-id="b5fb0-133">**File**: Your workbook ([selected with the file chooser](../../testing/power-automate-troubleshooting.md#select-workbooks-with-the-file-browser-control))</span></span>
    * <span data-ttu-id="b5fb0-134">**Сценарий:** имя сценария</span><span class="sxs-lookup"><span data-stu-id="b5fb0-134">**Script**: Your script name</span></span>

    :::image type="content" source="../../images/email-chart-sample-flow-1.png" alt-text="Завершенный соедините Excel Online (Бизнес) в Power Automate":::
1. <span data-ttu-id="b5fb0-136">В этом примере Outlook как клиент электронной почты.</span><span class="sxs-lookup"><span data-stu-id="b5fb0-136">This sample uses Outlook as the email client.</span></span> <span data-ttu-id="b5fb0-137">Можно использовать любые соединители электронной почты Power Automate поддерживает, но остальные действия предполагают, что вы выбрали Outlook.</span><span class="sxs-lookup"><span data-stu-id="b5fb0-137">You could use any email connector Power Automate supports, but the rest of the steps assume that you chose Outlook.</span></span> <span data-ttu-id="b5fb0-138">Добавьте новый **шаг,** использующий **соединителю Office 365 Outlook** и действие Отправка и электронная почта **(V2).**</span><span class="sxs-lookup"><span data-stu-id="b5fb0-138">Add a **New step** that uses the **Office 365 Outlook** connector and the **Send and email (V2)** action.</span></span> <span data-ttu-id="b5fb0-139">Используйте следующие значения для действия:</span><span class="sxs-lookup"><span data-stu-id="b5fb0-139">Use the following values for the action:</span></span>
    * <span data-ttu-id="b5fb0-140">**Чтобы:** ваша тестовая учетная запись электронной почты (или личная электронная почта)</span><span class="sxs-lookup"><span data-stu-id="b5fb0-140">**To**: Your test email account (or personal email)</span></span>
    * <span data-ttu-id="b5fb0-141">**Тема:** Просмотрите отчетные данные</span><span class="sxs-lookup"><span data-stu-id="b5fb0-141">**Subject**: Please Review Report Data</span></span>
    * <span data-ttu-id="b5fb0-142">Для поля **Body** выберите "Представление кода" () и `</>` введите следующее:</span><span class="sxs-lookup"><span data-stu-id="b5fb0-142">For the **Body** field, select "Code View" (`</>`) and enter the following:</span></span>

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

    :::image type="content" source="../../images/email-chart-sample-flow-2.png" alt-text="Завершенный соедините Office 365 Outlook в Power Automate":::
1. <span data-ttu-id="b5fb0-144">Сохраните поток и попробуйте его.</span><span class="sxs-lookup"><span data-stu-id="b5fb0-144">Save the flow and try it out.</span></span>

## <a name="training-video-extract-and-email-images-of-chart-and-table"></a><span data-ttu-id="b5fb0-145">Обучающее видео: извлечение и отправка изображений диаграммы и таблицы по электронной почте</span><span class="sxs-lookup"><span data-stu-id="b5fb0-145">Training video: Extract and email images of chart and table</span></span>

<span data-ttu-id="b5fb0-146">[Смотреть Sudhi Ramamurthy ходить через этот пример на YouTube](https://youtu.be/152GJyqc-Kw).</span><span class="sxs-lookup"><span data-stu-id="b5fb0-146">[Watch Sudhi Ramamurthy walk through this sample on YouTube](https://youtu.be/152GJyqc-Kw).</span></span>
