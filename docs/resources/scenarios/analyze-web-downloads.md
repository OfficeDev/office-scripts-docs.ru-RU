---
title: 'Сценарий примера сценариев Office: анализ загружаемых веб-файлов'
description: Пример, который принимает необработанные данные из Интернета в книгу Excel и определяет исходное расположение, прежде чем упорядочивать эту информацию в таблице.
ms.date: 06/25/2020
localization_priority: Normal
ms.openlocfilehash: 425d2af432d6b3c4b7604daf7935d2cc1ec059a8
ms.sourcegitcommit: 30750c4392db3ef057075a5702abb92863c93eda
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/01/2020
ms.locfileid: "44999269"
---
# <a name="office-scripts-sample-scenario-analyze-web-downloads"></a>Сценарий примера сценариев Office: анализ загружаемых веб-файлов

В этом сценарии вы являетесь задачей анализа загрузки отчетов с веб-сайта компании. Цель этого анализа — определить, поступает ли веб-трафик из Соединенных Штатов Америки или других стран мира.

Ваши коллеги отправляют необработанные данные в вашу книгу. В наборе данных каждой недели есть собственный лист. Кроме того, существует **сводный** лист с таблицей и диаграммой, в которой показаны тенденции за неделю.

Вы разрабатываете сценарий, который анализирует еженедельные загрузки данных на активном листе. Он будет анализировать IP-адрес, связанный с каждым загружаемым пакетом, и определять, был ли он передан из США. Ответ будет вставлен на лист в виде логического значения ("TRUE" или "FALSE"), а условное форматирование будет применено к этим ячейкам. Результаты размещения IP-адресов будут суммироваться на листе и скопированы в сводную таблицу.

## <a name="scripting-skills-covered"></a>Охваченные навыки работы со сценариями

- Синтаксический анализ текста
- Подфункции в скриптах
- Условное форматирование
- Таблицы

## <a name="demo-video"></a>Демонстрационное видео

В этом примере показана демонстрация при вызове сообщества разработчиков надстроек Office в течение февраля 2020.

> [!VIDEO https://www.youtube.com/embed/vPEqbb7t6-Y?start=154]

> [!NOTE]
> Код, показанный в этом видео, использует устаревшую модель API ( [сценарии Office для асинхронных API](../../develop/excel-async-model.md)). Пример, представленный на этой странице, был обновлен, но код выглядит немного иначе, чем запись. Изменения не влияют на поведение скрипта или другого контента в демо докладчика.

## <a name="setup-instructions"></a>Инструкции по настройке

1. Скачайте <a href="analyze-web-downloads.xlsx">analyze-web-downloads.xlsx</a> в OneDrive.

2. Откройте книгу с помощью Excel для веб-сайта.

3. На вкладке **Автоматизация** откройте **Редактор кода**.

4. В области задач **Редактор кода** нажмите **новый скрипт** и вставьте следующий скрипт в редактор.

    ```TypeScript
    function main(workbook: ExcelScript.Workbook) {
      // Get the Summary worksheet and table.
      let summaryWorksheet = workbook.getWorksheet("Summary");
      let summaryTable = summaryWorksheet?.getTable("Table1");
      if (!summaryWorksheet || !summaryTable) {
          console.log("The script expects the Summary worksheet with a summary table named Table1. Please download the correct template and try again.");
          return;
      }
  
      // Get the current worksheet.
      let currentWorksheet = workbook.getActiveWorksheet();
      if (!currentWorksheet.getName().toLocaleLowerCase().startsWith("week")) {
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
        conditionalFormatTrue.getCellValue().getFormat().getFill().setColor("#F8CCAD");
        conditionalFormatTrue.getCellValue().setRule({
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

5. Переименуйте сценарий, чтобы **проанализировать загрузку веб-файлов** и сохранить его.

## <a name="running-the-script"></a>Выполнение скрипта

Перейдите к любому листу **недели \* \* ** и запустите скрипт **анализа веб-загрузки** . Сценарий применит условное форматирование и расположение лабеллинг к текущему листу. Кроме того, будет обновлен лист **сводки** .

### <a name="before-running-the-script"></a>Перед выполнением скрипта

![Лист, отображающий необработанные данные веб-трафика.](../../images/scenario-analyze-web-downloads-before.png)

### <a name="after-running-the-script"></a>После выполнения скрипта

![Лист с отформатированными сведениями о расположении IP с предыдущими строками веб-трафика.](../../images/scenario-analyze-web-downloads-after.png)

![Сводная таблица и диаграмма, в которой перечисляются листы, на которых выполнен сценарий.](../../images/scenario-analyze-web-downloads-table.png)
