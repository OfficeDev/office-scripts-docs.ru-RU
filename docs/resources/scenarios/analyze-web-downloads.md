---
title: 'Office Сценарий примера сценариев: анализ веб-загрузки'
description: Пример, который принимает необработанные данные интернет-трафика в книге Excel и определяет расположение начала, прежде чем организовать эту информацию в таблицу.
ms.date: 06/29/2021
localization_priority: Normal
ms.openlocfilehash: a3ad957492184e358015d6fed5e3850a55f153b6722d1cd02ee8e4f5b2e39f93
ms.sourcegitcommit: 75f7ed8c2d23a104acc293f8ce29ea580b4fcdc5
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 08/11/2021
ms.locfileid: "57846334"
---
# <a name="office-scripts-sample-scenario-analyze-web-downloads"></a>Office Сценарий примера сценариев: анализ веб-загрузки

В этом сценарии вам будет поручено проанализировать отчеты о загрузке с веб-сайта вашей компании. Цель этого анализа состоит в том, чтобы определить, идет ли веб-трафик из США или других стран мира.

Ваши коллеги загружают необработанные данные в книгу. Набор данных каждую неделю имеет свой собственный таблицу. Существует также таблица **сводки** с таблицей и диаграммой, которая отображает тенденции недели за неделю.

Вы разработаете сценарий, который анализирует еженедельные скачивания данных в активном таблице. Он будет размазить IP-адрес, связанный с каждым скачиванием, и определить, поступил ли он из США или нет. Ответ будет вставлен в таблицу как значение boolean ("TRUE" или "FALSE") и условное форматирование будет применено к этим ячейкам. Результаты расположения IP-адресов будут подведены на таблицу и скопированы в сводную таблицу.

## <a name="scripting-skills-covered"></a>Навыки скриптов, охватываемых

- Размыв текста
- Subfunctions in scripts
- Условное форматирование
- таблицы;

## <a name="setup-instructions"></a>Инструкции по настройке

1. Скачайте <a href="analyze-web-downloads.xlsx">analyze-web-downloads.xlsx</a> в OneDrive.

1. Откройте книгу с помощью Excel для Интернета.

1. В **вкладке Automate** выберите **Новый скрипт** и вклеите следующий скрипт в редактор.

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
      let ipRangeValues = ipRange.getValues() as number[][];
      let logRangeValues = logRange.getValues() as string[][];
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
        let summaryHeaderRow = currentWorksheet.getRange("J1:M1");
        let summaryContentRow = currentWorksheet.getRange("J2:M2");
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

1. Переименуйте сценарий **для анализа веб-скачивания** и сохранения его.

## <a name="running-the-script"></a>Выполнение скрипта

Перейдите к любому из таблиц **Недели \* \*** и запустите сценарий **Анализ веб-загрузки.** Сценарий будет применять условное форматирование и маркировку расположения на текущем листе. Кроме того, будет обновлена **таблица Сводка.**

### <a name="before-running-the-script"></a>Перед запуском сценария

:::image type="content" source="../../images/scenario-analyze-web-downloads-before.png" alt-text="Таблица, в которую показаны необработанные данные веб-трафика.":::

### <a name="after-running-the-script"></a>После запуска скрипта

:::image type="content" source="../../images/scenario-analyze-web-downloads-after.png" alt-text="Таблица, которая отображает отформатированные сведения о расположении IP с предыдущими строками веб-трафика.":::

:::image type="content" source="../../images/scenario-analyze-web-downloads-table.png" alt-text="Сводная таблица и диаграмма, в которой обобщены таблицы, на которых был прорабатлан сценарий.":::
