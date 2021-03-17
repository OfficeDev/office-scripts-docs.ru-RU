---
title: Примеры скриптов Office Scripts в Excel в Интернете
description: Коллекция примеров кода, которые можно использовать с помощью Office Scripts в Excel в Интернете.
ms.date: 02/12/2021
localization_priority: Normal
ms.openlocfilehash: 4f1f6d4e160c42524df3c69228d182f1cb4838c8
ms.sourcegitcommit: 5bde455b06ee2ed007f3e462d8ad485b257774ef
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/17/2021
ms.locfileid: "50837280"
---
# <a name="sample-scripts-for-office-scripts-in-excel-on-the-web-preview"></a><span data-ttu-id="8dfea-103">Примеры скриптов Office Scripts в Excel в Интернете (предварительный просмотр)</span><span class="sxs-lookup"><span data-stu-id="8dfea-103">Sample scripts for Office Scripts in Excel on the web (preview)</span></span>

<span data-ttu-id="8dfea-104">Следующие примеры — это простые сценарии, которые можно попробовать в собственных книгах.</span><span class="sxs-lookup"><span data-stu-id="8dfea-104">The following samples are simple scripts for you to try on your own workbooks.</span></span> <span data-ttu-id="8dfea-105">Чтобы использовать их в Excel в Интернете:</span><span class="sxs-lookup"><span data-stu-id="8dfea-105">To use them in Excel on the web:</span></span>

1. <span data-ttu-id="8dfea-106">Откройте вкладку **Автоматизировать**.</span><span class="sxs-lookup"><span data-stu-id="8dfea-106">Open the **Automate** tab.</span></span>
2. <span data-ttu-id="8dfea-107">Редактор **кода прессы**.</span><span class="sxs-lookup"><span data-stu-id="8dfea-107">Press **Code Editor**.</span></span>
3. <span data-ttu-id="8dfea-108">Нажмите **новый скрипт** в области задач редактора кода.</span><span class="sxs-lookup"><span data-stu-id="8dfea-108">Press **New Script** in the Code Editor's task pane.</span></span>
4. <span data-ttu-id="8dfea-109">Замените весь сценарий образцом по вашему выбору.</span><span class="sxs-lookup"><span data-stu-id="8dfea-109">Replace the entire script with the sample of your choice.</span></span>
5. <span data-ttu-id="8dfea-110">Нажмите **кнопку** Выполнить в области задач редактора кода.</span><span class="sxs-lookup"><span data-stu-id="8dfea-110">Press **Run** in the Code Editor's task pane.</span></span>

[!INCLUDE [Preview note](../includes/preview-note.md)]

## <a name="scripting-basics"></a><span data-ttu-id="8dfea-111">Основы скриптов</span><span class="sxs-lookup"><span data-stu-id="8dfea-111">Scripting basics</span></span>

<span data-ttu-id="8dfea-112">В этих примерах демонстрируются основные строительные блоки для office Scripts.</span><span class="sxs-lookup"><span data-stu-id="8dfea-112">These samples demonstrate fundamental building blocks for Office Scripts.</span></span> <span data-ttu-id="8dfea-113">Добавьте их в скрипты, чтобы расширить решение и решить распространенные проблемы.</span><span class="sxs-lookup"><span data-stu-id="8dfea-113">Add these to your scripts to extend your solution and solve common problems.</span></span>

### <a name="read-and-log-one-cell"></a><span data-ttu-id="8dfea-114">Чтение и журнал одной ячейки</span><span class="sxs-lookup"><span data-stu-id="8dfea-114">Read and log one cell</span></span>

<span data-ttu-id="8dfea-115">В этом примере считывать значение **A1** и печатать его на консоли.</span><span class="sxs-lookup"><span data-stu-id="8dfea-115">This sample reads the value of **A1** and prints it to the console.</span></span>

```typescript
function main(workbook: ExcelScript.Workbook) {
  // Get the current worksheet.
  let selectedSheet = workbook.getActiveWorksheet();

  // Get the value of cell A1.
  let range = selectedSheet.getRange("A1");
  
  // Print the value of A1.
  console.log(range.getValue());
}
```

### <a name="read-the-active-cell"></a><span data-ttu-id="8dfea-116">Чтение активной ячейки</span><span class="sxs-lookup"><span data-stu-id="8dfea-116">Read the active cell</span></span>

<span data-ttu-id="8dfea-117">В этом скрипте региструется значение текущей активной ячейки.</span><span class="sxs-lookup"><span data-stu-id="8dfea-117">This script logs the value of the current active cell.</span></span> <span data-ttu-id="8dfea-118">Если выбрано несколько ячеек, будет зарегистрирована верхняя левая ячейка.</span><span class="sxs-lookup"><span data-stu-id="8dfea-118">If multiple cells are selected, the top-leftmost cell will be logged.</span></span>

```typescript
function main(workbook: ExcelScript.Workbook) {
  // Get the current active cell in the workbook.
  let cell = workbook.getActiveCell();

  // Log that cell's value.
  console.log(`The current cell's value is ${cell.getValue()}`);
}
```

### <a name="change-an-adjacent-cell"></a><span data-ttu-id="8dfea-119">Изменение соседней ячейки</span><span class="sxs-lookup"><span data-stu-id="8dfea-119">Change an adjacent cell</span></span>

<span data-ttu-id="8dfea-120">Этот скрипт получает соседние ячейки с использованием относительных ссылок.</span><span class="sxs-lookup"><span data-stu-id="8dfea-120">This script gets adjacent cells using relative references.</span></span> <span data-ttu-id="8dfea-121">Обратите внимание, что если активная ячейка находится в верхнем ряду, часть скрипта не работает, так как она ссылается на ячейку выше выбранной в настоящее время.</span><span class="sxs-lookup"><span data-stu-id="8dfea-121">Note that if the active cell is on the top row, part of the script fails, because it references the cell above the currently selected one.</span></span>

```typescript
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

### <a name="change-all-adjacent-cells"></a><span data-ttu-id="8dfea-122">Изменение всех смежных ячеек</span><span class="sxs-lookup"><span data-stu-id="8dfea-122">Change all adjacent cells</span></span>

<span data-ttu-id="8dfea-123">Этот скрипт копирует форматирование в активной ячейке в соседние ячейки.</span><span class="sxs-lookup"><span data-stu-id="8dfea-123">This script copies the formatting in the active cell to the neighboring cells.</span></span> <span data-ttu-id="8dfea-124">Обратите внимание, что этот скрипт работает только тогда, когда активная ячейка не на краю таблицы.</span><span class="sxs-lookup"><span data-stu-id="8dfea-124">Note that this script only works when the active cell isn't on an edge of the worksheet.</span></span>

```typescript
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

### <a name="change-each-individual-cell-in-a-range"></a><span data-ttu-id="8dfea-125">Изменение каждой отдельной ячейки в диапазоне</span><span class="sxs-lookup"><span data-stu-id="8dfea-125">Change each individual cell in a range</span></span>

<span data-ttu-id="8dfea-126">Этот скрипт цикличен по диапазону выбора в настоящее время.</span><span class="sxs-lookup"><span data-stu-id="8dfea-126">This script loops over the currently select range.</span></span> <span data-ttu-id="8dfea-127">Он очищает текущее форматирование и задает цвет заполнения в каждой ячейке случайным цветом.</span><span class="sxs-lookup"><span data-stu-id="8dfea-127">It clears the current formatting and sets the fill color in each cell to a random color.</span></span>

```typescript
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

### <a name="get-groups-of-cells-based-on-special-criteria"></a><span data-ttu-id="8dfea-128">Получить группы ячеек на основе специальных критериев</span><span class="sxs-lookup"><span data-stu-id="8dfea-128">Get groups of cells based on special criteria</span></span>

<span data-ttu-id="8dfea-129">Этот скрипт получает все пустые ячейки в используемом диапазоне текущего листа.</span><span class="sxs-lookup"><span data-stu-id="8dfea-129">This script gets all the blank cells in the current worksheet's used range.</span></span> <span data-ttu-id="8dfea-130">Затем он выделяет все эти ячейки с желтым фоном.</span><span class="sxs-lookup"><span data-stu-id="8dfea-130">It then highlights all those cells with a yellow background.</span></span>

```typescript
function main(workbook: ExcelScript.Workbook) {
    // Get the current used range.
    let range = workbook.getActiveWorksheet().getUsedRange();
    
    // Get all the blank cells.
    let blankCells = range.getSpecialCells(ExcelScript.SpecialCellType.blanks);

    // Highlight the blank cells with a yellow background.
    blankCells.getFormat().getFill().setColor("yellow");
}
```

## <a name="collections"></a><span data-ttu-id="8dfea-131">Коллекции</span><span class="sxs-lookup"><span data-stu-id="8dfea-131">Collections</span></span>

<span data-ttu-id="8dfea-132">Эти примеры работают с коллекциями объектов в книге.</span><span class="sxs-lookup"><span data-stu-id="8dfea-132">These samples work with collections of objects in the workbook.</span></span>

### <a name="iterating-over-collections"></a><span data-ttu-id="8dfea-133">Итерирование над коллекциями</span><span class="sxs-lookup"><span data-stu-id="8dfea-133">Iterating over collections</span></span>

<span data-ttu-id="8dfea-134">Этот скрипт получает и записывает имена всех таблиц в книге.</span><span class="sxs-lookup"><span data-stu-id="8dfea-134">This script gets and logs the names of all the worksheets in the workbook.</span></span> <span data-ttu-id="8dfea-135">Он также задает цвета вкладки случайным цветом.</span><span class="sxs-lookup"><span data-stu-id="8dfea-135">It also sets the their tab colors to a random color.</span></span>

```typescript
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

### <a name="querying-and-deleting-from-a-collection"></a><span data-ttu-id="8dfea-136">Запрос и удаление из коллекции</span><span class="sxs-lookup"><span data-stu-id="8dfea-136">Querying and deleting from a collection</span></span>

<span data-ttu-id="8dfea-137">Этот скрипт создает новую таблицу.</span><span class="sxs-lookup"><span data-stu-id="8dfea-137">This script creates a new worksheet.</span></span> <span data-ttu-id="8dfea-138">Он проверяет существующую копию листа и удаляет его перед созданием нового листа.</span><span class="sxs-lookup"><span data-stu-id="8dfea-138">It checks for an existing copy of the worksheet and deletes it before making a new sheet.</span></span>

```typescript
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

## <a name="dates"></a><span data-ttu-id="8dfea-139">Даты</span><span class="sxs-lookup"><span data-stu-id="8dfea-139">Dates</span></span>

<span data-ttu-id="8dfea-140">В примерах этого раздела покажите, как использовать объект [даты](https://developer.mozilla.org/docs/web/javascript/reference/global_objects/date) JavaScript.</span><span class="sxs-lookup"><span data-stu-id="8dfea-140">The samples in this section show how to use the JavaScript [Date](https://developer.mozilla.org/docs/web/javascript/reference/global_objects/date) object.</span></span>

<span data-ttu-id="8dfea-141">Следующий пример получает текущую дату и время, а затем записывает эти значения в две ячейки в активном таблице.</span><span class="sxs-lookup"><span data-stu-id="8dfea-141">The following sample gets the current date and time and then writes those values to two cells in the active worksheet.</span></span>

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

<span data-ttu-id="8dfea-142">В следующем примере считывалась дата, хранимая в Excel, и переводит ее на объект JavaScript Date.</span><span class="sxs-lookup"><span data-stu-id="8dfea-142">The next sample reads a date that's stored in Excel and translates it to a JavaScript Date object.</span></span> <span data-ttu-id="8dfea-143">В качестве ввода [для даты](https://support.office.com/article/now-function-3337fd29-145a-4347-b2e6-20c904739c46) JavaScript используется числовый серийный номер даты.</span><span class="sxs-lookup"><span data-stu-id="8dfea-143">It uses the [date's numeric serial number](https://support.office.com/article/now-function-3337fd29-145a-4347-b2e6-20c904739c46) as input for the JavaScript Date.</span></span>

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

## <a name="display-data"></a><span data-ttu-id="8dfea-144">Отображение данных</span><span class="sxs-lookup"><span data-stu-id="8dfea-144">Display data</span></span>

<span data-ttu-id="8dfea-145">В этих примерах показано, как работать с данными таблицы и предоставить пользователям лучшее представление или организацию.</span><span class="sxs-lookup"><span data-stu-id="8dfea-145">These samples demonstrate how to work with worksheet data and provide users with a better view or organization.</span></span>

### <a name="apply-conditional-formatting"></a><span data-ttu-id="8dfea-146">Применение условного форматирования</span><span class="sxs-lookup"><span data-stu-id="8dfea-146">Apply conditional formatting</span></span>

<span data-ttu-id="8dfea-147">В этом примере применяется условное форматирование к используемой в настоящее время линейке в таблице.</span><span class="sxs-lookup"><span data-stu-id="8dfea-147">This sample applies conditional formatting to the currently used range in the worksheet.</span></span> <span data-ttu-id="8dfea-148">Условное форматирование — это зеленое заполнение для 10% значений.</span><span class="sxs-lookup"><span data-stu-id="8dfea-148">The conditional formatting is a green fill for the top 10% of values.</span></span>

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

### <a name="create-a-sorted-table"></a><span data-ttu-id="8dfea-149">Создание отсортировать таблицу</span><span class="sxs-lookup"><span data-stu-id="8dfea-149">Create a sorted table</span></span>

<span data-ttu-id="8dfea-150">В этом примере создается таблица из используемого диапазона текущего таблицы, а затем сортируются на основе первого столбца.</span><span class="sxs-lookup"><span data-stu-id="8dfea-150">This sample creates a table from the current worksheet's used range, then sorts it based on the first column.</span></span>

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

### <a name="log-the-grand-total-values-from-a-pivottable"></a><span data-ttu-id="8dfea-151">Журнал значений "Grand Total" из pivotTable</span><span class="sxs-lookup"><span data-stu-id="8dfea-151">Log the "Grand Total" values from a PivotTable</span></span>

<span data-ttu-id="8dfea-152">В этом примере находится первый pivotTable в книге и регистрируемые значения в ячейках "Grand Total" (как выделено зеленым цветом на рисунке ниже).</span><span class="sxs-lookup"><span data-stu-id="8dfea-152">This sample finds the first PivotTable in the workbook and logs the values in the "Grand Total" cells (as highlighted in green in the image below).</span></span>

![A fruit sales PivotTable with the Grand Total row highlighted green.](../images/sample-pivottable-grand-total-row.png)

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

### <a name="use-data-validation-to-create-a-drop-down-list"></a><span data-ttu-id="8dfea-154">Использование проверки данных для создания выпадаемого списка</span><span class="sxs-lookup"><span data-stu-id="8dfea-154">Use data validation to create a drop-down list</span></span>

<span data-ttu-id="8dfea-155">В этом скрипте создается список выбора для ячейки.</span><span class="sxs-lookup"><span data-stu-id="8dfea-155">This script creates a drop-down selection list for a cell.</span></span> <span data-ttu-id="8dfea-156">Он использует существующие значения выбранного диапазона в качестве выбора для списка.</span><span class="sxs-lookup"><span data-stu-id="8dfea-156">It uses the existing values of the selected range as the choices for the list.</span></span>

![Набор снимков экрана до и после, который отображает три слова в диапазоне, а затем те же самые слова в отсвеяемом списке.](../images/sample-data-validation.png)

```typescript
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

## <a name="formulas"></a><span data-ttu-id="8dfea-158">Формулы</span><span class="sxs-lookup"><span data-stu-id="8dfea-158">Formulas</span></span>

<span data-ttu-id="8dfea-159">В этих примерах используются формулы Excel и покажите, как работать с ними в сценариях.</span><span class="sxs-lookup"><span data-stu-id="8dfea-159">These samples use Excel formulas and show how to work with them in scripts.</span></span>

### <a name="single-formula"></a><span data-ttu-id="8dfea-160">Единая формула</span><span class="sxs-lookup"><span data-stu-id="8dfea-160">Single formula</span></span>

<span data-ttu-id="8dfea-161">Этот скрипт задает формулу ячейки, а затем отображает, как Excel хранит формулу и значение ячейки отдельно.</span><span class="sxs-lookup"><span data-stu-id="8dfea-161">This script sets a cell's formula, then displays how Excel stores the cell's formula and value separately.</span></span>

```typescript
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

### <a name="spilling-results-from-a-formula"></a><span data-ttu-id="8dfea-162">Проливка результатов из формулы</span><span class="sxs-lookup"><span data-stu-id="8dfea-162">Spilling results from a formula</span></span>

<span data-ttu-id="8dfea-163">Этот скрипт передает диапазон "A1:D2" на "A4:B7" с помощью функции TRANSPOSE.</span><span class="sxs-lookup"><span data-stu-id="8dfea-163">This script transposes the range "A1:D2" to "A4:B7" by using the TRANSPOSE function.</span></span> <span data-ttu-id="8dfea-164">Если переливание приводит к ошибке #SPILL, он очищает целевой диапазон и снова применяет формулу.</span><span class="sxs-lookup"><span data-stu-id="8dfea-164">If the transpose results in a #SPILL error, it clears the target range and applies the formula again.</span></span>

```typescript
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

## <a name="scenario-samples"></a><span data-ttu-id="8dfea-165">Примеры сценариев</span><span class="sxs-lookup"><span data-stu-id="8dfea-165">Scenario samples</span></span>

<span data-ttu-id="8dfea-166">Примеры более крупных решений в реальном мире можно найти в примере сценариев [сценариев Office Scripts.](scenarios/sample-scenario-overview.md)</span><span class="sxs-lookup"><span data-stu-id="8dfea-166">For samples showcasing larger, real-world solutions, visit [Sample scenarios for Office Scripts](scenarios/sample-scenario-overview.md).</span></span>

## <a name="suggest-new-samples"></a><span data-ttu-id="8dfea-167">Предложить новые примеры</span><span class="sxs-lookup"><span data-stu-id="8dfea-167">Suggest new samples</span></span>

<span data-ttu-id="8dfea-168">Мы приветствуем предложения по новым образцам.</span><span class="sxs-lookup"><span data-stu-id="8dfea-168">We welcome suggestions for new samples.</span></span> <span data-ttu-id="8dfea-169">Если существует распространенный сценарий, который поможет другим разработчикам скриптов, сообщите нам в разделе отзывов ниже.</span><span class="sxs-lookup"><span data-stu-id="8dfea-169">If there is a common scenario that would help other script developers, please tell us in the feedback section below.</span></span>
