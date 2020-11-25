---
title: 'Сценарий примера сценариев Office: анализ загружаемых веб-файлов'
description: Пример, который принимает необработанные данные из Интернета в книгу Excel и определяет исходное расположение, прежде чем упорядочивать эту информацию в таблице.
ms.date: 07/10/2020
localization_priority: Normal
ms.openlocfilehash: adc2cb401830b66b245c0dfcc4441b7ac9c8c61f
ms.sourcegitcommit: 009935c5773761c5833e5857491af47e2c95d851
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 11/25/2020
ms.locfileid: "49408968"
---
# <a name="office-scripts-sample-scenario-analyze-web-downloads"></a><span data-ttu-id="9f639-103">Сценарий примера сценариев Office: анализ загружаемых веб-файлов</span><span class="sxs-lookup"><span data-stu-id="9f639-103">Office Scripts sample scenario: Analyze web downloads</span></span>

<span data-ttu-id="9f639-104">В этом сценарии вы являетесь задачей анализа загрузки отчетов с веб-сайта компании.</span><span class="sxs-lookup"><span data-stu-id="9f639-104">In this scenario, you're tasked with analyzing download reports from your company's website.</span></span> <span data-ttu-id="9f639-105">Цель этого анализа — определить, поступает ли веб-трафик из Соединенных Штатов Америки или других стран мира.</span><span class="sxs-lookup"><span data-stu-id="9f639-105">The goal of this analysis is to determine if the web traffic is coming from the United States or elsewhere in the world.</span></span>

<span data-ttu-id="9f639-106">Ваши коллеги отправляют необработанные данные в вашу книгу.</span><span class="sxs-lookup"><span data-stu-id="9f639-106">Your colleagues upload the raw data to your workbook.</span></span> <span data-ttu-id="9f639-107">В наборе данных каждой недели есть собственный лист.</span><span class="sxs-lookup"><span data-stu-id="9f639-107">Each week's set of data has its own worksheet.</span></span> <span data-ttu-id="9f639-108">Кроме того, существует **сводный** лист с таблицей и диаграммой, в которой показаны тенденции за неделю.</span><span class="sxs-lookup"><span data-stu-id="9f639-108">There is also the **Summary** worksheet with a table and chart that shows week-over-week trends.</span></span>

<span data-ttu-id="9f639-109">Вы разрабатываете сценарий, который анализирует еженедельные загрузки данных на активном листе.</span><span class="sxs-lookup"><span data-stu-id="9f639-109">You'll develop a script that analyzes weekly downloads data in the active worksheet.</span></span> <span data-ttu-id="9f639-110">Он будет анализировать IP-адрес, связанный с каждым загружаемым пакетом, и определять, был ли он передан из США.</span><span class="sxs-lookup"><span data-stu-id="9f639-110">It will parse the IP address associated with each download and determine whether or not it came from the US.</span></span> <span data-ttu-id="9f639-111">Ответ будет вставлен на лист в виде логического значения ("TRUE" или "FALSE"), а условное форматирование будет применено к этим ячейкам.</span><span class="sxs-lookup"><span data-stu-id="9f639-111">The answer will be inserted in the worksheet as a boolean value ("TRUE" or "FALSE") and conditional formatting will be applied to those cells.</span></span> <span data-ttu-id="9f639-112">Результаты размещения IP-адресов будут суммироваться на листе и скопированы в сводную таблицу.</span><span class="sxs-lookup"><span data-stu-id="9f639-112">The IP address location results will be totaled on the worksheet and copied to the summary table.</span></span>

## <a name="scripting-skills-covered"></a><span data-ttu-id="9f639-113">Охваченные навыки работы со сценариями</span><span class="sxs-lookup"><span data-stu-id="9f639-113">Scripting skills covered</span></span>

- <span data-ttu-id="9f639-114">Синтаксический анализ текста</span><span class="sxs-lookup"><span data-stu-id="9f639-114">Text parsing</span></span>
- <span data-ttu-id="9f639-115">Подфункции в скриптах</span><span class="sxs-lookup"><span data-stu-id="9f639-115">Subfunctions in scripts</span></span>
- <span data-ttu-id="9f639-116">Условное форматирование</span><span class="sxs-lookup"><span data-stu-id="9f639-116">Conditional formatting</span></span>
- <span data-ttu-id="9f639-117">Таблицы</span><span class="sxs-lookup"><span data-stu-id="9f639-117">Tables</span></span>

## <a name="demo-video"></a><span data-ttu-id="9f639-118">Демонстрационное видео</span><span class="sxs-lookup"><span data-stu-id="9f639-118">Demo video</span></span>

<span data-ttu-id="9f639-119">В этом примере показана демонстрация при вызове сообщества разработчиков надстроек Office в течение февраля 2020.</span><span class="sxs-lookup"><span data-stu-id="9f639-119">This sample was demoed as part of the Office Add-ins developer community call for February 2020.</span></span>

> [!VIDEO https://www.youtube.com/embed/vPEqbb7t6-Y?start=154]

> [!NOTE]
> <span data-ttu-id="9f639-120">Код, показанный в этом видео, использует устаревшую модель API ( [сценарии Office для асинхронных API](../../develop/excel-async-model.md)).</span><span class="sxs-lookup"><span data-stu-id="9f639-120">The code shown in this video uses an older API model (the [Office Scripts Async APIs](../../develop/excel-async-model.md)).</span></span> <span data-ttu-id="9f639-121">Пример, представленный на этой странице, был обновлен, но код выглядит немного иначе, чем запись.</span><span class="sxs-lookup"><span data-stu-id="9f639-121">The sample presented on this page has been updated, but the code looks a little different from the recording.</span></span> <span data-ttu-id="9f639-122">Изменения не влияют на поведение скрипта или другого контента в демо докладчика.</span><span class="sxs-lookup"><span data-stu-id="9f639-122">The changes don't affect the behavior of the script or the other content in the presenter's demo.</span></span>

## <a name="setup-instructions"></a><span data-ttu-id="9f639-123">Инструкции по настройке</span><span class="sxs-lookup"><span data-stu-id="9f639-123">Setup instructions</span></span>

1. <span data-ttu-id="9f639-124">Скачайте <a href="analyze-web-downloads.xlsx">analyze-web-downloads.xlsx</a> в OneDrive.</span><span class="sxs-lookup"><span data-stu-id="9f639-124">Download <a href="analyze-web-downloads.xlsx">analyze-web-downloads.xlsx</a> to your OneDrive.</span></span>

2. <span data-ttu-id="9f639-125">Откройте книгу с помощью Excel для веб-сайта.</span><span class="sxs-lookup"><span data-stu-id="9f639-125">Open the workbook with Excel for the web.</span></span>

3. <span data-ttu-id="9f639-126">На вкладке **Автоматизация** откройте **Редактор кода**.</span><span class="sxs-lookup"><span data-stu-id="9f639-126">Under the **Automate** tab, open the **Code Editor**.</span></span>

4. <span data-ttu-id="9f639-127">В области задач **Редактор кода** нажмите **новый скрипт** и вставьте следующий скрипт в редактор.</span><span class="sxs-lookup"><span data-stu-id="9f639-127">In the **Code Editor** task pane, press **New Script** and paste the following script into the editor.</span></span>

    ```TypeScript
    function main(workbook: ExcelScript.Workbook) {
      /* Get the Summary worksheet and table.
       * End the script early if either object is not in the workbook.
       */
      let summaryWorksheet = workbook.getWorksheet("Summary");
      if (!summaryWorksheet) {
        console.log("The script expects a worksheet named \"Summary\". Please download the correct template and try again.");
        return;
      }
      let summaryTable = summaryWorksheet.getTable("Table1");
      if (!summaryTable) {
        console.log("The script expects a summary table named \"Table1\". Please download the correct template and try again.");
        return;
      }

      // Get the current worksheet.
      let currentWorksheet = workbook.getActiveWorksheet();
      if (currentWorksheet.getName().toLocaleLowerCase().indexOf("week") !== 0) {
        console.log("Please switch worksheet to one of the weekly data sheets and try again.")
        return;
      }

      // Get the values of the active range of the active worksheet.
      let logRange = currentWorksheet.getUsedRange();

      if (logRange.getColumnCount() !== 8) {
        console.log(`Verify that you are on the correct worksheet. Either the week's data has been already processed or the content is incorrect. The following columns are expected: ${[
          "Time Stamp", "IP Address", "kilobytes", "user agent code", "milliseconds", "Request", "Results", "Referrer"
        ]}`);
        return;
      }
      // Get the range that will contain TRUE/FALSE if the IP address is from the United States (US).
      let isUSColumn = logRange
        .getLastColumn()
        .getOffsetRange(0, 1);

      // Get the values of all the US IP addresses.
      let ipRange = workbook.getWorksheet("USIPAddresses").getUsedRange();
      let ipRangeValues = ipRange.getValues();
      let logRangeValues = logRange.getValues();
      // Remove the first row.
      let topRow = logRangeValues.shift();
      console.log(`Analyzing ${logRangeValues.length} entries.`);

      // Create a new array to contain the boolean representing if this is a US IP address.
      let newCol = [];

      // Go through each row in worksheet and add Boolean.
      for (let i = 0; i < logRangeValues.length; i++) {
        let curRowIP = logRangeValues[i][1];
        if (findIP(ipRangeValues, ipAddressToInteger(curRowIP)) > 0) {
          newCol.push([true]);
        } else {
          newCol.push([false]);
        }
      }

      // Remove the empty column header and add proper heading.
      newCol = [["Is US IP"], ...newCol];

      // Write the result to the spreadsheet.
      console.log(`Adding column to indicate whether IP belongs to US region or not at address: ${isUSColumn.getAddress()}`);
      console.log(newCol.length);
      console.log(newCol);
      isUSColumn.setValues(newCol);

      // Call the local function to add summary data to the worksheet.
      addSummaryData();

      // Call the local function to apply conditional formatting.
      applyConditionalFormatting(isUSColumn);

      // Autofit columns.
      currentWorksheet.getUsedRange().getFormat().autofitColumns();

      // Get the calculated summary data.
      let summaryRangeValues = currentWorksheet.getRange("J2:M2").getValues();

      // Add the corresponding row to the summary table.
      summaryTable.addRow(null, summaryRangeValues[0]);
      console.log("Complete.");
      return;

      /**
       * A function to add summary data on the worksheet.
       */
      function addSummaryData() {
        // Add a summary row and table.
        let summaryHeader = [["Year", "Week", "US", "Other"]];
        let countTrueFormula =
          "=COUNTIF(" + isUSColumn.getAddress() + ', "=TRUE")/' + (newCol.length - 1);
        let countFalseFormula =
          "=COUNTIF(" + isUSColumn.getAddress() + ', "=FALSE")/' + (newCol.length - 1);

        let summaryContent = [
          [
            '=TEXT(A2,"YYYY")',
            '=TEXTJOIN(" ", FALSE, "Wk", WEEKNUM(A2))',
            countTrueFormula,
            countFalseFormula
          ]
        ];
        let summaryHeaderRow = currentWorksheet
          .getRange("J1:M1");
        let summaryContentRow = currentWorksheet
          .getRange("J2:M2");
        console.log("2");

        summaryHeaderRow.setValues(summaryHeader);
        console.log("3");

        summaryContentRow.setValues(summaryContent);
        console.log("4");

        let formats = [[".000", ".000"]];
        summaryContentRow
          .getOffsetRange(0, 2)
          .getResizedRange(0, -2).setNumberFormats(formats);
      }
    }
    /**
     * Apply conditional formatting based on TRUE/FALSE values of the Is US IP column.
     */
    function applyConditionalFormatting(isUSColumn: ExcelScript.Range) {
      // Add conditional formatting to the new column.
      let conditionalFormatTrue = isUSColumn.addConditionalFormat(
        ExcelScript.ConditionalFormatType.cellValue
      );
      let conditionalFormatFalse = isUSColumn.addConditionalFormat(
        ExcelScript.ConditionalFormatType.cellValue
      );
      // Set TRUE to light blue and FALSE to light orange.
      conditionalFormatTrue.getCellValue().getFormat().getFill().setColor("#8FA8DB");
      conditionalFormatTrue.getCellValue().setRule({
        formula1: "=TRUE",
        operator: ExcelScript.ConditionalCellValueOperator.equalTo
      });
      conditionalFormatFalse.getCellValue().getFormat().getFill().setColor("#F8CCAD");
      conditionalFormatFalse.getCellValue().setRule({
        formula1: "=FALSE",
        operator: ExcelScript.ConditionalCellValueOperator.equalTo
      });
    }
    /**
     * Translate an IP address into an integer.
     * @param ipAddress: IP address to verify.
     */
    function ipAddressToInteger(ipAddress: string): number {
      // Split the IP address into octets.
      let octets = ipAddress.split(".");

      // Create a number for each octet and do the math to create the integer value of the IP address.
      let fullNum =
        // Define an arbitrary number for the last octet.
        111 +
        parseInt(octets[2]) * 256 +
        parseInt(octets[1]) * 65536 +
        parseInt(octets[0]) * 16777216;
      return fullNum;
    }
    /**
     * Return the row number where the ip address is found.
     * @param ipLookupTable IP look-up table.
     * @param n IP address to number value.  
     */
    function findIP(ipLookupTable: number[][], n: number): number {
      for (let i = 0; i < ipLookupTable.length; i++) {
        if (ipLookupTable[i][0] <= n && ipLookupTable[i][1] >= n) {
          return i;
        }
      }
      return -1;
    }
    ```

5. <span data-ttu-id="9f639-128">Переименуйте сценарий, чтобы **проанализировать загрузку веб-файлов** и сохранить его.</span><span class="sxs-lookup"><span data-stu-id="9f639-128">Rename the script to **Analyze Web Downloads** and save it.</span></span>

## <a name="running-the-script"></a><span data-ttu-id="9f639-129">Выполнение скрипта</span><span class="sxs-lookup"><span data-stu-id="9f639-129">Running the script</span></span>

<span data-ttu-id="9f639-130">Перейдите к любому листу **недели \* \*** и запустите скрипт **анализа веб-загрузки** .</span><span class="sxs-lookup"><span data-stu-id="9f639-130">Navigate to any of the **Week\*\*** worksheets and run the **Analyze Web Downloads** script.</span></span> <span data-ttu-id="9f639-131">Сценарий применит условное форматирование и расположение лабеллинг к текущему листу.</span><span class="sxs-lookup"><span data-stu-id="9f639-131">The script will apply the conditional formatting and location labelling on the current sheet.</span></span> <span data-ttu-id="9f639-132">Кроме того, будет обновлен лист **сводки** .</span><span class="sxs-lookup"><span data-stu-id="9f639-132">It will also update the **Summary** worksheet.</span></span>

### <a name="before-running-the-script"></a><span data-ttu-id="9f639-133">Перед выполнением скрипта</span><span class="sxs-lookup"><span data-stu-id="9f639-133">Before running the script</span></span>

![Лист, отображающий необработанные данные веб-трафика.](../../images/scenario-analyze-web-downloads-before.png)

### <a name="after-running-the-script"></a><span data-ttu-id="9f639-135">После выполнения скрипта</span><span class="sxs-lookup"><span data-stu-id="9f639-135">After running the script</span></span>

![Лист с отформатированными сведениями о расположении IP с предыдущими строками веб-трафика.](../../images/scenario-analyze-web-downloads-after.png)

![Сводная таблица и диаграмма, в которой перечисляются листы, на которых выполнен сценарий.](../../images/scenario-analyze-web-downloads-table.png)
