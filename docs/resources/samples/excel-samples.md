---
title: Основные сценарии Office скриптов в Excel в Интернете
description: Коллекция примеров кода, которые можно использовать с Office скриптами в Excel в Интернете.
ms.date: 03/24/2022
ms.localizationpriority: medium
ms.openlocfilehash: 853b00349b246e74765eb2959b4926fad42f07c5
ms.sourcegitcommit: 7023b9e23499806901a5ecf8ebc460b76887cca6
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/31/2022
ms.locfileid: "64585872"
---
# <a name="basic-scripts-for-office-scripts-in-excel-on-the-web"></a>Основные сценарии Office скриптов в Excel в Интернете

Следующие примеры — это простые сценарии, которые можно попробовать в собственных книгах. Чтобы использовать их в Excel в Интернете:

1. Откройте вкладку **Автоматизировать**.
1. Выберите **Создать сценарий**.
1. Замените весь сценарий образцом по вашему выбору.
1. Выберите **Выполнить** в области задач редактора кода.

## <a name="script-basics"></a>Основы скрипта

В этих примерах демонстрируются основные строительные блоки для Office скриптов. Расширите эти скрипты, чтобы расширить решение и решить распространенные проблемы.

### <a name="read-and-log-one-cell"></a>Чтение и журнал одной ячейки

В этом примере считывать значение **A1** и печатать его на консоли.

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Get the current worksheet.
  let selectedSheet = workbook.getActiveWorksheet();

  // Get the value of cell A1.
  let range = selectedSheet.getRange("A1");
  
  // Print the value of A1.
  console.log(range.getValue());
}
```

### <a name="read-the-active-cell"></a>Чтение активной ячейки

В этом скрипте региструется значение текущей активной ячейки. Если выбрано несколько ячеек, будет зарегистрирована верхняя левая ячейка.

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Get the current active cell in the workbook.
  let cell = workbook.getActiveCell();

  // Log that cell's value.
  console.log(`The current cell's value is ${cell.getValue()}`);
}
```

### <a name="change-an-adjacent-cell"></a>Изменение соседней ячейки

Этот скрипт получает соседние ячейки с использованием относительных ссылок. Обратите внимание, что если активная ячейка находится в верхнем ряду, часть скрипта не работает, так как она ссылается на ячейку выше выбранной в настоящее время.

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Get the currently active cell in the workbook.
  let activeCell = workbook.getActiveCell();
  console.log(`The active cell's address is: ${activeCell.getAddress()}`);

  // Get the cell to the right of the active cell and set its value and color.
  let rightCell = activeCell.getOffsetRange(0,1);
  rightCell.setValue("Right cell");
  console.log(`The right cell's address is: ${rightCell.getAddress()}`);
  rightCell.getFormat().getFont().setColor("Magenta");
  rightCell.getFormat().getFill().setColor("Cyan");

  // Get the cell to the above of the active cell and set its value and color.
  // Note that this operation will fail if the active cell is in the top row.
  let aboveCell = activeCell.getOffsetRange(-1, 0);
  aboveCell.setValue("Above cell");
  console.log(`The above cell's address is: ${aboveCell.getAddress()}`);
  aboveCell.getFormat().getFont().setColor("White");
  aboveCell.getFormat().getFill().setColor("Black");
}
```

### <a name="change-all-adjacent-cells"></a>Изменение всех смежных ячеек

Этот скрипт копирует форматирование в активной ячейке в соседние ячейки. Обратите внимание, что этот скрипт работает только тогда, когда активная ячейка не на краю таблицы.

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Get the active cell.
  let activeCell = workbook.getActiveCell();

  // Get the cell that's one row above and one column to the left of the active cell.
  let cornerCell = activeCell.getOffsetRange(-1,-1);

  // Get a range that includes all the cells surrounding the active cell.
  let surroundingRange = cornerCell.getResizedRange(2, 2)

  // Copy the formatting from the active cell to the new range.
  surroundingRange.copyFrom(
    activeCell, /* The source range. */
    ExcelScript.RangeCopyType.formats /* What to copy. */
    );
}
```

### <a name="change-each-individual-cell-in-a-range"></a>Изменение каждой отдельной ячейки в диапазоне

Этот скрипт цикличен по диапазону выбора в настоящее время. Он очищает текущее форматирование и задает цвет заполнения в каждой ячейке случайным цветом.

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Get the currently selected range.
  let range = workbook.getSelectedRange();

  // Get the size boundaries of the range.
  let rows = range.getRowCount();
  let cols = range.getColumnCount();

  // Clear any existing formatting
  range.clear(ExcelScript.ClearApplyTo.formats);

  // Iterate over the range.
  for (let row = 0; row < rows; row++) {
    for (let col = 0; col < cols; col++) {
      // Generate a random color hex-code.
      let colorString = `#${Math.random().toString(16).substr(-6)}`;

      // Set the color of the current cell to that random hex-code.
      range.getCell(row, col).getFormat().getFill().setColor(colorString);
    }
  }
}
```

### <a name="get-groups-of-cells-based-on-special-criteria"></a>Получить группы ячеек на основе специальных критериев

Этот скрипт получает все пустые ячейки в используемом диапазоне текущего листа. Затем он выделяет все эти ячейки с желтым фоном.

```TypeScript
function main(workbook: ExcelScript.Workbook) {
    // Get the current used range.
    let range = workbook.getActiveWorksheet().getUsedRange();
    
    // Get all the blank cells.
    let blankCells = range.getSpecialCells(ExcelScript.SpecialCellType.blanks);

    // Highlight the blank cells with a yellow background.
    blankCells.getFormat().getFill().setColor("yellow");
}
```

## <a name="collections"></a>Коллекции

Эти примеры работают с коллекциями объектов в книге.

### <a name="iterate-over-collections"></a>Итерации над коллекциями

Этот скрипт получает и записывает имена всех таблиц в книге. Он также задает цвета вкладки случайным цветом.

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Get all the worksheets in the workbook.
  let sheets = workbook.getWorksheets();

  // Get a list of all the worksheet names.
  let names = sheets.map ((sheet) => sheet.getName());

  // Write in the console all the worksheet names and the total count.
  console.log(names);
  console.log(`Total worksheets inside of this workbook: ${sheets.length}`);
  
  // Set the tab color each worksheet to a random color
  for (let sheet of sheets) {
    // Generate a random color hex-code.
    let colorString = `#${Math.random().toString(16).substr(-6)}`;

    // Set the color of the current worksheet's tab to that random hex-code.
    sheet.setTabColor(colorString);
  }
}
```

### <a name="query-and-delete-from-a-collection"></a>Запрос и удаление из коллекции

Этот скрипт создает новую таблицу. Он проверяет существующую копию листа и удаляет его перед созданием нового листа.

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Name of the worksheet to be added.
  let name = "Index";

  // Get any worksheet with that name.
  let sheet = workbook.getWorksheet("Index");
  
  // If `null` wasn't returned, then there's already a worksheet with that name.
  if (sheet) {
    console.log(`Worksheet by the name ${name} already exists. Deleting it.`);
    // Delete the sheet.
    sheet.delete();
  }
  
  // Add a blank worksheet with the name "Index".
  // Note that this code runs regardless of whether an existing sheet was deleted.
  console.log(`Adding the worksheet named ${name}.`);
  let newSheet = workbook.addWorksheet("Index");

  // Switch to the new worksheet.
  newSheet.activate();
}
```

## <a name="dates"></a>Даты

В примерах этого раздела покажите, как использовать объект [даты](https://developer.mozilla.org/docs/web/javascript/reference/global_objects/date) JavaScript.

Следующий пример получает текущую дату и время, а затем записывает эти значения в две ячейки в активном таблице.

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Get the cells at A1 and B1.
  let dateRange = workbook.getActiveWorksheet().getRange("A1");
  let timeRange = workbook.getActiveWorksheet().getRange("B1");

  // Get the current date and time with the JavaScript Date object.
  let date = new Date(Date.now());

  // Add the date string to A1.
  dateRange.setValue(date.toLocaleDateString());

  // Add the time string to B1.
  timeRange.setValue(date.toLocaleTimeString());
}
```

В следующем примере считываю дату, храняную в Excel, и переводим ее на объект JavaScript Date. В качестве ввода для даты JavaScript используется числовый серийный номер даты. Этот серийный номер описан в статье [функции NOW()](https://support.office.com/article/now-function-3337fd29-145a-4347-b2e6-20c904739c46) .

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Read a date at cell A1 from Excel.
  let dateRange = workbook.getActiveWorksheet().getRange("A1");

  // Convert the Excel date to a JavaScript Date object.
  let excelDateValue = dateRange.getValue() as number;
  let javaScriptDate = new Date(Math.round((excelDateValue - 25569) * 86400 * 1000));
  console.log(javaScriptDate);
}
```

## <a name="display-data"></a>Отображение данных

В этих примерах показано, как работать с данными таблицы и предоставить пользователям лучшее представление или организацию.

### <a name="apply-conditional-formatting"></a>Применение условного форматирования

В этом примере применяется условное форматирование к используемой в настоящее время линейке в таблице. Условное форматирование — это зеленое заполнение для 10% значений.

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Get the current worksheet.
  let selectedSheet = workbook.getActiveWorksheet();

  // Get the used range in the worksheet.
  let range = selectedSheet.getUsedRange();

  // Set the fill color to green for the top 10% of values in the range.
  let conditionalFormat = range.addConditionalFormat(ExcelScript.ConditionalFormatType.topBottom)
  conditionalFormat.getTopBottom().getFormat().getFill().setColor("green");
  conditionalFormat.getTopBottom().setRule({
    rank: 10, // The percentage threshold.
    type: ExcelScript.ConditionalTopBottomCriterionType.topPercent // The type of the top/bottom condition.
  });
}
```

### <a name="create-a-sorted-table"></a>Создание отсортировать таблицу

В этом примере создается таблица из используемого диапазона текущего таблицы, а затем сортируются на основе первого столбца.

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Get the current worksheet.
  let selectedSheet = workbook.getActiveWorksheet();

  // Create a table with the used cells.
  let usedRange = selectedSheet.getUsedRange();
  let newTable = selectedSheet.addTable(usedRange, true);

  // Sort the table using the first column.
  newTable.getSort().apply([{ key: 0, ascending: true }]);
}
```

### <a name="log-the-grand-total-values-from-a-pivottable"></a>Журнал значений "Grand Total" из pivotTable

В этом примере находится первый pivotTable в книге и регистрируемые значения в ячейках "Grand Total" (как выделено зеленым цветом на рисунке ниже).

:::image type="content" source="../../images/sample-pivottable-grand-total-row.png" alt-text="PivotTable, показывающий продажи фруктов с строкой Grand Total, выделенной зеленым цветом.":::

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Get the first PivotTable in the workbook.
  let pivotTable = workbook.getPivotTables()[0];

  // Get the names of each data column in the PivotTable.
  let pivotColumnLabelRange = pivotTable.getLayout().getColumnLabelRange();

  // Get the range displaying the pivoted data.
  let pivotDataRange = pivotTable.getLayout().getBodyAndTotalRange();

  // Get the range with the "grand totals" for the PivotTable columns.
  let grandTotalRange = pivotDataRange.getLastRow();

  // Print each of the "Grand Totals" to the console.
  grandTotalRange.getValues()[0].forEach((column, columnIndex) => {
    console.log(`Grand total of ${pivotColumnLabelRange.getValues()[0][columnIndex]}: ${grandTotalRange.getValues()[0][columnIndex]}`);
    // Example log: "Grand total of Sum of Crates Sold Wholesale: 11000"
  });
}
```

### <a name="create-a-drop-down-list-using-data-validation"></a>Создание выпадаемого списка с помощью проверки данных

В этом скрипте создается список выбора для ячейки. Он использует существующие значения выбранного диапазона в качестве выбора для списка.

:::image type="content" source="../../images/sample-data-validation.png" alt-text="Лист, на котором отображается диапазон из трех ячеек, содержащих варианты цвета &quot;красный, синий, зеленый&quot; и рядом с ней, те же варианты, показанные в выпадаемом списке.":::

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Get the values for data validation.
  let selectedRange = workbook.getSelectedRange();
  let rangeValues = selectedRange.getValues();

  // Convert the values into a comma-delimited string.
  let dataValidationListString = "";
  rangeValues.forEach((rangeValueRow) => {
    rangeValueRow.forEach((value) => {
      dataValidationListString += value + ",";
    });
  });

  // Clear the old range.
  selectedRange.clear(ExcelScript.ClearApplyTo.contents);

  // Apply the data validation to the first cell in the selected range.
  let targetCell = selectedRange.getCell(0,0);
  let dataValidation = targetCell.getDataValidation();

  // Set the content of the drop-down list.
  dataValidation.setRule({
      list: {
        inCellDropDown: true,
        source: dataValidationListString
      }
    });
}
```

## <a name="formulas"></a>Формулы

В этих примерах Excel формулы и покажите, как работать с ними в сценариях.

### <a name="single-formula"></a>Единая формула

Этот скрипт задает формулу ячейки, а затем отображает, Excel хранит формулу и значение ячейки отдельно.

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  let selectedSheet = workbook.getActiveWorksheet();

  // Set A1 to 2.
  let a1 = selectedSheet.getRange("A1");
  a1.setValue(2);

  // Set B1 to the formula =(2*A1), which should equal 4.
  let b1 = selectedSheet.getRange("B1")
  b1.setFormula("=(2*A1)");

  // Log the current results for `getFormula` and `getValue` at B1.
  console.log(`B1 - Formula: ${b1.getFormula()} | Value: ${b1.getValue()}`);
}
```

### <a name="handle-a-spill-error-returned-from-a-formula"></a>Обработка ошибки `#SPILL!` , возвращаемой из формулы

Этот скрипт передает диапазон "A1:D2" на "A4:B7" с помощью функции TRANSPOSE. Если переливание приводит к ошибке `#SPILL` , он очищает целевой диапазон и снова применяет формулу.

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  let sheet = workbook.getActiveWorksheet();
  // Use the data in A1:D2 for the sample.
  let dataAddress = "A1:D2"
  let inputRange = sheet.getRange(dataAddress);

  // Place the transposed data starting at A4.
  let targetStartCell = sheet.getRange("A4");

  // Compute the target range.
  let targetRange = targetStartCell.getResizedRange(inputRange.getColumnCount() - 1, inputRange.getRowCount() - 1);

  // Call the transpose helper function.
  targetStartCell.setFormula(`=TRANSPOSE(${dataAddress})`);

  // Check if the range update resulted in a spill error.
  let checkValue = targetStartCell.getValue() as string;
  if (checkValue === '#SPILL!') {
    // Clear the target range and call the transpose function again.
    console.log("Target range has data that is preventing update. Clearing target range.");
    targetRange.clear();
    targetStartCell.setFormula(`=TRANSPOSE(${dataAddress})`);
  }

  // Select the transposed range to highlight it.
  targetRange.select();
}
```

## <a name="suggest-new-samples"></a>Предложить новые примеры

Мы приветствуем предложения по новым образцам. Если существует распространенный сценарий, который поможет другим разработчикам сценариев, сообщите нам об этом в разделе отзывов в нижней части страницы.

## <a name="see-also"></a>См. также

* [Sudhi Ramamurthy's "Range basics" on YouTube](https://youtu.be/4emjkOFdLBA)
* [Office сценарии и сценарии](samples-overview.md)
* [Запись, редактирование и создание сценариев Office в Excel в Интернете](../../tutorials/excel-tutorial.md)
