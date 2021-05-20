---
title: Основные сценарии для Office сценариев в Excel в Интернете
description: Коллекция образцов кода для использования с помощью Office в Excel в Интернете.
ms.date: 05/17/2021
localization_priority: Normal
ms.openlocfilehash: f252934a92126212b9520223826b3b2f5161ed57
ms.sourcegitcommit: 4687693f02fc90a57ba30c461f35046e02e6f5fb
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 05/19/2021
ms.locfileid: "52545761"
---
# <a name="basic-scripts-for-office-scripts-in-excel-on-the-web"></a>Основные сценарии для Office сценариев в Excel в Интернете

Следующие образцы простые сценарии для вас, чтобы попробовать на ваших собственных трудовых книжек. Чтобы использовать их в Excel в Интернете:

1. Откройте вкладку **Автоматизировать**.
2. Редактор **пресс-кода**.
3. Нажмите **новый** сценарий в панели задач редактора Кода.
4. Замените весь скрипт на образец по вашему выбору.
5. **Нажмите** Вы запустите в панели задач редактора Кода.

## <a name="script-basics"></a>Основы сценария

Эти образцы демонстрируют фундаментальные строительные блоки для Office скриптов. Добавьте их в скрипты, чтобы расширить решение и решить общие проблемы.

### <a name="read-and-log-one-cell"></a>Читать и регистрировать одну ячейку

Этот образец считывает значение **A1** и печатает его на консоли.

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

### <a name="read-the-active-cell"></a>Читать активную ячейку

Этот скрипт регистрирует значение текущей активной ячейки. Если выбрано несколько ячеек, в журнал будет зарегистрирована ячейка верхней левой.

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Get the current active cell in the workbook.
  let cell = workbook.getActiveCell();

  // Log that cell's value.
  console.log(`The current cell's value is ${cell.getValue()}`);
}
```

### <a name="change-an-adjacent-cell"></a>Изменение соседней ячейки

Этот скрипт получает смежные ячейки с использованием относительных ссылок. Обратите внимание, что если активная ячейка находится в верхнем ряду, часть скрипта выходит из строя, поскольку она ссылается на ячейку выше выбранной в настоящее время ячейки.

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

Этот скрипт копирует форматирование в активной ячейке в соседние ячейки. Обратите внимание, что этот скрипт работает только тогда, когда активная ячейка не находится на краю листа.

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

Этот скрипт циклов над в настоящее время выбрать диапазон. Он очищает текущее форматирование и устанавливает цвет заполнения в каждой ячейке на случайный цвет.

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

Этот скрипт получает все пустые ячейки в используемом диапазоне текущего листа. Затем он выделяет все эти клетки с желтым фоном.

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

Эти образцы работают с коллекциями объектов в рабочей книге.

### <a name="iterate-over-collections"></a>Итерировать над коллекциями

Этот скрипт получает и регистрирует имена всех листов в рабочей книге. Он также устанавливает их цвета вкладок на случайный цвет.

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

Этот скрипт создает новый лист. Он проверяет существующую копию листа и удаляет его перед созданием нового листа.

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

Образцы в этом разделе показывают, как использовать объект JavaScript [Дата.](https://developer.mozilla.org/docs/web/javascript/reference/global_objects/date)

Следующий пример получает текущую дату и время, а затем записывает эти значения на две ячейки в активном листе.

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

Следующий пример считывает дату, хранящуюся в Excel переводит ее на объект JavaScript Date. Он использует [числовой серийный номер даты в качестве](https://support.office.com/article/now-function-3337fd29-145a-4347-b2e6-20c904739c46) ввода для даты JavaScript.

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

Эти образцы демонстрируют, как работать с данными листа и предоставить пользователям лучшее представление или организацию.

### <a name="apply-conditional-formatting"></a>Применение условного форматирования

Этот пример применяет условное форматирование к используемому в настоящее время диапазону в листе. Условное форматирование является зеленым заливки для верхних 10% значений.

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

Этот пример создает таблицу из используемого диапазона текущего листа, а затем сортирует ее на основе первого столбца.

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

### <a name="log-the-grand-total-values-from-a-pivottable"></a>Войти в значения "Grand Total" с PivotTable

Этот пример находит первый PivotTable в рабочей книге и регистрирует значения в ячейках "Grand Total" (как выделено зеленым цветом на рисунке ниже).

:::image type="content" source="../../images/sample-pivottable-grand-total-row.png" alt-text="PivotTable, показывающий продажи фруктов с Grand Total строки подчеркнул зеленый":::

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

### <a name="create-a-drop-down-list-using-data-validation"></a>Создание списка отсева с использованием проверки данных

Этот скрипт создает список выпадают из списка выбора ячейки. В качестве выбора списка он использует существующие значения выбранного диапазона.

:::image type="content" source="../../images/sample-data-validation.png" alt-text="Лист, показывающий ряд из трех ячеек, содержащих цветовые варианты &quot;красный, синий, зеленый&quot; и рядом с ним, тот же выбор, показанный в списке выпадают":::

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

В этих образцах Excel формулы и покажем, как с ними работать в скриптах.

### <a name="single-formula"></a>Единая формула

Этот скрипт устанавливает формулу ячейки, а затем отображает, Excel хранит формулу и значение ячейки отдельно.

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

### <a name="handle-a-spill-error-returned-from-a-formula"></a>Ручка `#SPILL!` ошибки, возвращенной из формулы

Этот скрипт транспонирует диапазон "A1:D2" на "A4:B7" с помощью функции TRANSPOSE. Если транспонирует `#SPILL` ошибку, он очищает целевой диапазон и применяет формулу снова.

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

## <a name="suggest-new-samples"></a>Предложить новые образцы

Мы приветствуем предложения по новым образцам. Если есть общий сценарий, который поможет другим разработчикам скриптов, пожалуйста, сообщите нам в разделе обратной связи в нижней части страницы.

## <a name="see-also"></a>См. также

* [Судхи Рамамурти "Основы диапазона" на YouTube](https://youtu.be/4emjkOFdLBA)
* [Office Образцы сценариев и сценарии](samples-overview.md)
* [Запись, редактирование и создание сценариев Office в Excel в Интернете](../../tutorials/excel-tutorial.md)
