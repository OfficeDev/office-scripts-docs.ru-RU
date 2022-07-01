---
title: Основные сценарии для сценариев Office в Excel
description: Коллекция примеров кода для использования со скриптами Office в Excel.
ms.date: 06/24/2022
ms.localizationpriority: medium
ms.openlocfilehash: b6588dc4109799a7d615d0bee38c82a2bcd16743
ms.sourcegitcommit: 82fb78e6907b7c3b95c5c53cfc83af4ea1067a78
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/01/2022
ms.locfileid: "66572351"
---
# <a name="basic-scripts-for-office-scripts-in-excel"></a>Основные сценарии для сценариев Office в Excel

Приведенные ниже примеры — это простые сценарии, которые позволяют опробовать собственные книги. Чтобы использовать их в Excel, выполните следующие действия.

1. Откройте книгу в Excel в Интернете.
1. Откройте вкладку **Автоматизировать**.
1. Выберите **Создать сценарий**.
1. Замените весь скрипт примером по вашему выбору.
1. Выберите **"Выполнить** " в области задач редактора кода.

## <a name="script-basics"></a>Основные сведения о скриптах

В этих примерах демонстрируются основные стандартные блоки для сценариев Office. Разверните эти скрипты, чтобы расширить решение и решить распространенные проблемы.

### <a name="read-and-log-one-cell"></a>Чтение и запись в журнал одной ячейки

Этот пример считывает значение **A1** и выводит его на консоль.

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

Этот скрипт регистрирует значение текущей активной ячейки. Если выбрано несколько ячеек, будет зарегистрирована левая верхняя ячейка.

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Get the current active cell in the workbook.
  let cell = workbook.getActiveCell();

  // Log that cell's value.
  console.log(`The current cell's value is ${cell.getValue()}`);
}
```

### <a name="change-an-adjacent-cell"></a>Изменение смежных ячеек

Этот скрипт получает смежные ячейки с использованием относительных ссылок. Обратите внимание, что если активная ячейка находится в верхней строке, часть скрипта завершается ошибкой, так как она ссылается на ячейку выше выбранной в данный момент.

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

Этот скрипт копирует форматирование в активной ячейке в соседние ячейки. Обратите внимание, что этот сценарий работает только в том случае, если активная ячейка не на границе листа.

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

Этот сценарий выполняет цикл по диапазону выбора в данный момент. Он очищает текущее форматирование и задает цвет заливки в каждой ячейке случайным цветом.

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

### <a name="get-groups-of-cells-based-on-special-criteria"></a>Получение групп ячеек на основе специальных критериев

Этот скрипт получает все пустые ячейки в используемом диапазоне текущего листа. Затем он выделяет все ячейки с желтым фоном.

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

### <a name="iterate-over-collections"></a>Итерация по коллекциям

Этот скрипт получает и регистрирует имена всех листов в книге. Он также задает цвета вкладок случайным цветом.

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

Этот скрипт создает новый лист. Он проверяет наличие существующей копии листа и удаляет его перед созданием нового листа.

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

В примерах в этом разделе показано, как использовать объект [даты](https://developer.mozilla.org/docs/web/javascript/reference/global_objects/date) JavaScript.

Следующий пример получает текущую дату и время, а затем записывает эти значения в две ячейки активного листа.

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

В следующем примере считываются даты, хранящиеся в Excel, и преобразуется в объект даты JavaScript. Он использует числовой серийный номер даты в качестве входных данных для даты JavaScript. Этот серийный номер описан в статье о [функции NOW(](https://support.office.com/article/now-function-3337fd29-145a-4347-b2e6-20c904739c46) ).

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

В этих примерах показано, как работать с данными листа и предоставлять пользователям более эффективное представление или организацию.

### <a name="apply-conditional-formatting"></a>Применение условного форматирования

Этот пример применяет условное форматирование к используемму в настоящее время диапазону на листе. Условное форматирование представляет собой зеленую заливку для 10 % значений.

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

### <a name="create-a-sorted-table"></a>Создание отсортированной таблицы

Этот пример создает таблицу из используемого диапазона текущего листа, а затем сортирует ее по первому столбцу.

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

### <a name="filter-a-table"></a>Фильтрация таблицы

Этот пример фильтрует существующую таблицу, используя значения в одном из столбцов.

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Get the table in the workbook named "StationTable".
  const table = workbook.getTable("StationTable");

  // Get the "Station" table column for the filter.
  const stationColumn = table.getColumnByName("Station");

  // Apply a filter to the table that will only show rows 
  // with a value of "Station-1" in the "Station" column.
  stationColumn.getFilter().applyValuesFilter(["Station-1"]);
}
```

> [!TIP]
> Скопируйте отфильтрованную информацию в книге с помощью `Range.copyFrom`. Добавьте следующую строку в конец скрипта, чтобы создать лист с отфильтрованными данными.
>
> ```typescript
>   workbook.addWorksheet().getRange("A1").copyFrom(table.getRange());
> ```

### <a name="log-the-grand-total-values-from-a-pivottable"></a>Занося в журнал значения "Общий итог" из сводная таблица

Этот пример находит первый сводная таблица в книге и регистрирует значения в ячейках "Общий итог" (как выделено зеленым цветом на рисунке ниже).

:::image type="content" source="../../images/sample-pivottable-grand-total-row.png" alt-text="В сводная таблица показана продажа цветов с выделенной зеленой строкой &quot;Общий итог&quot;.":::

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

### <a name="create-a-drop-down-list-using-data-validation"></a>Создание раскрывающегося списка с использованием проверки данных

Этот скрипт создает раскрывающийся список выбора для ячейки. Он использует существующие значения выбранного диапазона в качестве вариантов для списка.

:::image type="content" source="../../images/sample-data-validation.png" alt-text="Лист с диапазоном из трех ячеек с вариантами цвета &quot;красный, синий, зеленый&quot; и рядом с ним те же варианты, что и в раскрывающемся списке.":::

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

В этих примерах используются формулы Excel и показано, как работать с ними в скриптах.

### <a name="single-formula"></a>Одна формула

Этот скрипт задает формулу ячейки, а затем показывает, как Excel хранит формулу и значение ячейки отдельно.

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

### <a name="handle-a-spill-error-returned-from-a-formula"></a>Обработка ошибки `#SPILL!` , возвращаемой формулой

Этот скрипт транспонирует диапазон "A1:D2" в "A4:B7" с помощью функции TRANSPOSE. Если транспонирование приводит к ошибке `#SPILL` , он очищает целевой диапазон и снова применяет формулу.

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

### <a name="replace-all-formulas-with-their-result-values"></a>Замените все формулы значениями результатов.

Этот скрипт заменяет каждую ячейку текущего листа, содержащую формулу, на результат этой формулы. Это означает, что после выполнения скрипта не будут использоваться формулы, а только значения.

```TypeScript
function main(workbook: ExcelScript.Workbook) {
    // Get the ranges with formulas.
    let sheet = workbook.getActiveWorksheet();
    let usedRange = sheet.getUsedRange();
    let formulaCells = usedRange.getSpecialCells(ExcelScript.SpecialCellType.formulas);

    // In each formula range: get the current value, clear the contents, and set the value as the old one.
    // This removes the formula but keeps the result.
    formulaCells.getAreas().forEach((range) => {
      let currentValues = range.getValues();
      range.clear(ExcelScript.ClearApplyTo.contents);
      range.setValues(currentValues);
    });
}
```

## <a name="suggest-new-samples"></a>Предложить новые примеры

Мы будем рады новым примерам. Если существует распространенный сценарий, который может помочь другим разработчикам скриптов, сообщите нам об этом в разделе отзывов в нижней части страницы.

## <a name="see-also"></a>См. также

* [Основные сведения о диапазоне sudhi Мхи Химсти (Sudhi Мхайтхи) на YouTube](https://youtu.be/4emjkOFdLBA)
* [Примеры и сценарии сценариев Office](samples-overview.md)
* [Запись, редактирование и создание сценариев Office в Excel в Интернете](../../tutorials/excel-tutorial.md)
