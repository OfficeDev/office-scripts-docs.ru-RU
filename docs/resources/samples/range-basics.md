---
title: Основы диапазона в сценариях Office
description: Узнайте основы использования объекта Range в скриптах Office.
ms.date: 04/01/2021
localization_priority: Normal
ms.openlocfilehash: 73eeba086aace6262c624de9074ffb301f6532bd
ms.sourcegitcommit: 5d24e77df70aa2c1c982275d53213c2a9323ff86
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 04/02/2021
ms.locfileid: "51571302"
---
# <a name="range-basics"></a><span data-ttu-id="0a781-103">Основы диапазона</span><span class="sxs-lookup"><span data-stu-id="0a781-103">Range basics</span></span>

<span data-ttu-id="0a781-104">`Range` является фундаментальным объектом в объектной модели Office Scripts Excel.</span><span class="sxs-lookup"><span data-stu-id="0a781-104">`Range` is the foundational object within the Office Scripts Excel object model.</span></span> <span data-ttu-id="0a781-105">[API диапазона](/javascript/api/office-scripts/excelscript/excelscript.range) позволяют получать доступ к данным и формату, доступным в сетке, и связывать другие ключевые объекты в Excel, такие как таблицы, таблицы, диаграммы и т.д.</span><span class="sxs-lookup"><span data-stu-id="0a781-105">[Range APIs](/javascript/api/office-scripts/excelscript/excelscript.range) allow access to both data and format available on the grid and link other key objects within Excel such as worksheets, tables, charts, etc.</span></span>

<span data-ttu-id="0a781-106">Диапазон идентифицирован с помощью такого адреса, как "A1:B4" или с помощью имени элемента, который является именем ключа для данного набора ячеек.</span><span class="sxs-lookup"><span data-stu-id="0a781-106">A range is identified using its address such as "A1:B4" or using a named-item, which is a named key for a given set of cells.</span></span> <span data-ttu-id="0a781-107">В объектной модели Excel ячейки и группы ячеек называются _диапазоном._</span><span class="sxs-lookup"><span data-stu-id="0a781-107">In the Excel object model, both a cell and group of cells are referred as _range_.</span></span> <span data-ttu-id="0a781-108">`Range` может содержать атрибуты уровня ячейки, такие как данные внутри ячейки, а также атрибуты уровня ячейки и ячейки, такие как формат, границы и т.д.</span><span class="sxs-lookup"><span data-stu-id="0a781-108">`Range` can contain cell-level attributes such as data within a cell and also cell and cells-level attributes such as format, borders, etc.</span></span>

<span data-ttu-id="0a781-109">`Range` можно также получить с помощью выбора пользователя, состоящего по крайней мере из одной ячейки.</span><span class="sxs-lookup"><span data-stu-id="0a781-109">`Range` can also be obtained via user's selection that consists of at least one cell.</span></span> <span data-ttu-id="0a781-110">При взаимодействии с диапазоном важно сохранить эти отношения между ячейками и диапазонами.</span><span class="sxs-lookup"><span data-stu-id="0a781-110">As you interact with the range, it's important to keep these cell and range relationships clear.</span></span>

<span data-ttu-id="0a781-111">Ниже приводится основной набор геттеров, сеттеров и других полезных методов, наиболее часто используемых в скриптах.</span><span class="sxs-lookup"><span data-stu-id="0a781-111">Following are the core set of getters, setters, and other useful methods most often used in scripts.</span></span> <span data-ttu-id="0a781-112">Это отличная отправная точка для вашего путешествия по API.</span><span class="sxs-lookup"><span data-stu-id="0a781-112">This is a great starting point for your API journey.</span></span> <span data-ttu-id="0a781-113">В более поздних разделах сгруппировка методов и помощь в создании умственной модели при начале разблокирования `Range` API объекта.</span><span class="sxs-lookup"><span data-stu-id="0a781-113">The later sections group the methods and help to build a mental model as you begin to unlock the `Range` object's APIs.</span></span>

## <a name="example-scripts"></a><span data-ttu-id="0a781-114">Примеры скриптов</span><span class="sxs-lookup"><span data-stu-id="0a781-114">Example scripts</span></span>

* [<span data-ttu-id="0a781-115">Базовое чтение и написание</span><span class="sxs-lookup"><span data-stu-id="0a781-115">Basic read and write</span></span>](#basic-read-and-write)
* [<span data-ttu-id="0a781-116">Добавление строки в конце таблицы</span><span class="sxs-lookup"><span data-stu-id="0a781-116">Add row at the end of worksheet</span></span>](#add-row-at-the-end-of-worksheet)
* [<span data-ttu-id="0a781-117">Фильтр четких столбцов</span><span class="sxs-lookup"><span data-stu-id="0a781-117">Clear column filter</span></span>](clear-table-filter-for-active-cell.md)
* [<span data-ttu-id="0a781-118">Цвет каждой ячейки уникальным цветом</span><span class="sxs-lookup"><span data-stu-id="0a781-118">Color each cell with unique color</span></span>](#color-each-cell-with-unique-color)
* [<span data-ttu-id="0a781-119">Диапазон обновлений со значениями с помощью 2-мерного (2D) массива</span><span class="sxs-lookup"><span data-stu-id="0a781-119">Update range with values using 2-dimensional (2D) array</span></span>](#update-range-with-values-using-2d-array)

### <a name="basic-read-and-write"></a><span data-ttu-id="0a781-120">Базовое чтение и написание</span><span class="sxs-lookup"><span data-stu-id="0a781-120">Basic read and write</span></span>

```TypeScript
/**
 * This script demonstrates basic read-write operations on the Range object.
 */
function main(workbook: ExcelScript.Workbook) {
  const cell = workbook.getActiveCell();
  const prevValue = cell.getValue();
  if (prevValue) {
      console.log(`Active cell's value is: ${prevValue}`);
  } else {
      console.log("Setting active cell's value..");
      cell.setValue("Sample");
  }

  // Get cell next to the right column and set its value and fill color.
  const nextCell = cell.getOffsetRange(0,1);
  nextCell.setValue("Next cell");
  console.log(`Next cell's address is: ${nextCell.getAddress()}`);
  console.log("Setting fill color and font color of next cell...");
  nextCell.getFormat().getFill().setColor("Magenta");
  nextCell.getFormat().getFill().setColor("Cyan");

  // Get the target range address to update with 2-dimensional value.
  const dataRange = nextCell.getOffsetRange(1, 0).getResizedRange(2, 1);
  const DATA = [
    [10, 7],
    [8, 15],
    [12, 1]
  ];
  console.log(`Updating range ${dataRange.getAddress()} with values: ${DATA}`);
  dataRange.setValues(DATA);

  // Formula range.
  const formulaRange = dataRange.getOffsetRange(3, 0).getRow(0);
  console.log(`Updating formula for range: ${formulaRange.getAddress()}`)
  // Since relative formula is being set, we can set the formula of the entire range to the same value.
  formulaRange.setFormulaR1C1("=SUM(R[-3]C:R[-1]C)");
  console.log(`Updating number format for range: ${formulaRange.getAddress()}`)
  // Since the number format is common to the entire range, we can set it to a common format.
  formulaRange.setNumberFormat("0.00");
  return;
}
```

### <a name="add-row-at-the-end-of-worksheet"></a><span data-ttu-id="0a781-121">Добавление строки в конце таблицы</span><span class="sxs-lookup"><span data-stu-id="0a781-121">Add row at the end of worksheet</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
    const sheet = workbook.getWorksheet('Sheet5');
    const data = ['2016', 'Bikes', 'Seats', '1500', .05];
    addRow(sheet, data);
    return;
}

function addRow(sheet: ExcelScript.Worksheet, data: (string | number | boolean)[]): void {

    const usedRange = sheet.getUsedRange();
    let startCell: ExcelScript.Range;
    // If the sheet is empty, then use A1 as starting cell for the update.
    if (usedRange) {
      startCell = usedRange.getLastRow().getCell(0, 0).getOffsetRange(1, 0);
    } else {
      startCell = sheet.getRange('A1');
    }
    console.log(startCell.getAddress());
    const targetRange = startCell.getResizedRange(0, data.length - 1);
    targetRange.setValues([data]);
    return;
}
```

### <a name="color-each-cell-with-unique-color"></a><span data-ttu-id="0a781-122">Цвет каждой ячейки уникальным цветом</span><span class="sxs-lookup"><span data-stu-id="0a781-122">Color each cell with unique color</span></span>

```TypeScript
/**
 * This sample demonstrates how to iterate over a selected range and set cell property.
   It colors each cell within the selected range with a random color.
 */
function main(workbook: ExcelScript.Workbook) {

    const syncStart = new Date().getTime();
    // Get selected range
    const range = workbook.getSelectedRange();
    const rows = range.getRowCount();
    const cols = range.getColumnCount();
    console.log("Start");

    // Color each cell with random color.
    for (let row = 0; row < rows; row++) {
        for (let col = 0; col < cols; col++) {
            range
                .getCell(row, col)
                .getFormat()
                .getFill()
                .setColor(`#${Math.random().toString(16).substr(-6)}`);
        }
    }

    console.log("End");
    const syncEnd = new Date().getTime();
    console.log("Completed, took: " + (syncEnd - syncStart) / 1000 + " Sec");
}
```

### <a name="update-range-with-values-using-2d-array"></a><span data-ttu-id="0a781-123">Диапазон обновлений со значениями с помощью массива 2D</span><span class="sxs-lookup"><span data-stu-id="0a781-123">Update range with values using 2D array</span></span>

<span data-ttu-id="0a781-124">Динамически вычисляется измерение диапазона для обновления на основе значений массива 2D.</span><span class="sxs-lookup"><span data-stu-id="0a781-124">Dynamically calculates the range dimension to update based on 2D array values.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  const currentCell = workbook.getActiveCell();
  let inputRange = computeTargetRange(currentCell, DATA);
  // Set range values.
  console.log(inputRange.getAddress());
  inputRange.setValues(DATA);
  // Call a helper function to place border around the range.
  borderAround(inputRange);
}

/**
 * A helper function that computes the target range given the target range's starting cell and selected range. 
 */
function computeTargetRange(targetCell: ExcelScript.Range, data: string[][]): ExcelScript.Range {
  const targetRange = targetCell.getResizedRange(data.length - 1, data[0].length - 1);
  return targetRange;
}

/**
 * A helper function that places a border around the range.
 */
function borderAround(range: ExcelScript.Range): void {
  range.getFormat().getRangeBorder(ExcelScript.BorderIndex.edgeLeft).setStyle(ExcelScript.BorderLineStyle.dash);
  range.getFormat().getRangeBorder(ExcelScript.BorderIndex.edgeRight).setStyle(ExcelScript.BorderLineStyle.dash);
  range.getFormat().getRangeBorder(ExcelScript.BorderIndex.edgeTop).setStyle(ExcelScript.BorderLineStyle.dash);
  range.getFormat().getRangeBorder(ExcelScript.BorderIndex.edgeBottom).setStyle(ExcelScript.BorderLineStyle.dash);
  return;
}

// Values used for range setup.
const DATA = [
  ['Item', 'Bread', 'Donuts', 'Cookies', 'Cakes', 'Pies'],
  ['Amount', '2', '1.5', '4', '12', '26']
]
```

## <a name="training-videos-range-basics"></a><span data-ttu-id="0a781-125">Обучающие видео: основы диапазона</span><span class="sxs-lookup"><span data-stu-id="0a781-125">Training videos: Range basics</span></span>

<span data-ttu-id="0a781-126">_Основы диапазона_</span><span class="sxs-lookup"><span data-stu-id="0a781-126">_Range basics_</span></span>

<span data-ttu-id="0a781-127">[![Просмотр пошагового видео в базовых диапазонах](../../images/rangebasics-vid.png)](https://youtu.be/4emjkOFdLBA "Пошаговая видеозапись об основах диапазона")</span><span class="sxs-lookup"><span data-stu-id="0a781-127">[![Watch step-by-step video on Range basics](../../images/rangebasics-vid.png)](https://youtu.be/4emjkOFdLBA "Step-by-step video on Range basics")</span></span>

<span data-ttu-id="0a781-128">_Добавление строки в конце таблицы_</span><span class="sxs-lookup"><span data-stu-id="0a781-128">_Add row at the end of worksheet_</span></span>

<span data-ttu-id="0a781-129">[![Просмотрите пошаговую видеозапись добавления строки в конце таблицы](../../images/rangebasics-addrow-vid.png)](https://youtu.be/RgtUar013D0 "Пошаговая видеозапись добавления строки в конце таблицы")</span><span class="sxs-lookup"><span data-stu-id="0a781-129">[![Watch step-by-step video on how to add a row at the end of a worksheet](../../images/rangebasics-addrow-vid.png)](https://youtu.be/RgtUar013D0 "Step-by-step video on how to add a row at the end of a worksheet")</span></span>

## <a name="methods-that-return-some-range-metadata"></a><span data-ttu-id="0a781-130">Методы, возвращая метаданные некоторых диапазонов</span><span class="sxs-lookup"><span data-stu-id="0a781-130">Methods that return some range metadata</span></span>

* <span data-ttu-id="0a781-131">getAddress(), getAddressLocal()</span><span class="sxs-lookup"><span data-stu-id="0a781-131">getAddress(), getAddressLocal()</span></span>
* <span data-ttu-id="0a781-132">getCellCount()</span><span class="sxs-lookup"><span data-stu-id="0a781-132">getCellCount()</span></span>
* <span data-ttu-id="0a781-133">getRowCount(), getColumnCount()</span><span class="sxs-lookup"><span data-stu-id="0a781-133">getRowCount(), getColumnCount()</span></span>

## <a name="methods-that-return-dataconstants-associated-with-a-given-range"></a><span data-ttu-id="0a781-134">Методы возврата данных и констант, связанных с заданным диапазоном</span><span class="sxs-lookup"><span data-stu-id="0a781-134">Methods that return data/constants associated with a given range</span></span>

### <a name="returned-as-single-cell-value"></a><span data-ttu-id="0a781-135">Возвращено в качестве одноклеточного значения</span><span class="sxs-lookup"><span data-stu-id="0a781-135">Returned as single cell value</span></span>

* <span data-ttu-id="0a781-136">getFormula(), getFormulaLocal()</span><span class="sxs-lookup"><span data-stu-id="0a781-136">getFormula(), getFormulaLocal()</span></span>
* <span data-ttu-id="0a781-137">getFormulaR1C1()</span><span class="sxs-lookup"><span data-stu-id="0a781-137">getFormulaR1C1()</span></span>
* <span data-ttu-id="0a781-138">getNumberFormat(), getNumberFormatLocal()</span><span class="sxs-lookup"><span data-stu-id="0a781-138">getNumberFormat(), getNumberFormatLocal()</span></span>
* <span data-ttu-id="0a781-139">getText()</span><span class="sxs-lookup"><span data-stu-id="0a781-139">getText()</span></span>
* <span data-ttu-id="0a781-140">getValue()</span><span class="sxs-lookup"><span data-stu-id="0a781-140">getValue()</span></span>
* <span data-ttu-id="0a781-141">getValueType()</span><span class="sxs-lookup"><span data-stu-id="0a781-141">getValueType()</span></span>

### <a name="returned-as-2d-arrays-whole-range"></a><span data-ttu-id="0a781-142">Возвращается в качестве 2D-массивов (весь диапазон)</span><span class="sxs-lookup"><span data-stu-id="0a781-142">Returned as 2D arrays (whole range)</span></span>

* <span data-ttu-id="0a781-143">getFormulas(), getFormulasLocal()</span><span class="sxs-lookup"><span data-stu-id="0a781-143">getFormulas(), getFormulasLocal()</span></span>
* <span data-ttu-id="0a781-144">getFormulasR1C1()</span><span class="sxs-lookup"><span data-stu-id="0a781-144">getFormulasR1C1()</span></span>
* <span data-ttu-id="0a781-145">getNumberFormatCategories()</span><span class="sxs-lookup"><span data-stu-id="0a781-145">getNumberFormatCategories()</span></span>
* <span data-ttu-id="0a781-146">getNumberFormats(), getNumberFormatsLocal()</span><span class="sxs-lookup"><span data-stu-id="0a781-146">getNumberFormats(), getNumberFormatsLocal()</span></span>
* <span data-ttu-id="0a781-147">getTexts()</span><span class="sxs-lookup"><span data-stu-id="0a781-147">getTexts()</span></span>
* <span data-ttu-id="0a781-148">getValues()</span><span class="sxs-lookup"><span data-stu-id="0a781-148">getValues()</span></span>
* <span data-ttu-id="0a781-149">getValueTypes()</span><span class="sxs-lookup"><span data-stu-id="0a781-149">getValueTypes()</span></span>
* <span data-ttu-id="0a781-150">getHidden()</span><span class="sxs-lookup"><span data-stu-id="0a781-150">getHidden()</span></span>
* <span data-ttu-id="0a781-151">getIsEntireRow()</span><span class="sxs-lookup"><span data-stu-id="0a781-151">getIsEntireRow()</span></span>
* <span data-ttu-id="0a781-152">getIsEntireColumn()</span><span class="sxs-lookup"><span data-stu-id="0a781-152">getIsEntireColumn()</span></span>

## <a name="methods-that-return-other-range-object"></a><span data-ttu-id="0a781-153">Методы, возвращая другой объект диапазона</span><span class="sxs-lookup"><span data-stu-id="0a781-153">Methods that return other range object</span></span>

* <span data-ttu-id="0a781-154">getSurroundingRegion() — аналогично CurrentRegion в VBA</span><span class="sxs-lookup"><span data-stu-id="0a781-154">getSurroundingRegion() -- similar to CurrentRegion in VBA</span></span>
* <span data-ttu-id="0a781-155">getCell (строка, столбец)</span><span class="sxs-lookup"><span data-stu-id="0a781-155">getCell(row, column)</span></span>
* <span data-ttu-id="0a781-156">getColumn(column)</span><span class="sxs-lookup"><span data-stu-id="0a781-156">getColumn(column)</span></span>
* <span data-ttu-id="0a781-157">getColumnHidden()</span><span class="sxs-lookup"><span data-stu-id="0a781-157">getColumnHidden()</span></span>
* <span data-ttu-id="0a781-158">getColumnsAfter (count)</span><span class="sxs-lookup"><span data-stu-id="0a781-158">getColumnsAfter(count)</span></span>
* <span data-ttu-id="0a781-159">getColumnsBefore (count)</span><span class="sxs-lookup"><span data-stu-id="0a781-159">getColumnsBefore(count)</span></span>
* <span data-ttu-id="0a781-160">getEntireColumn()</span><span class="sxs-lookup"><span data-stu-id="0a781-160">getEntireColumn()</span></span>
* <span data-ttu-id="0a781-161">getEntireRow()</span><span class="sxs-lookup"><span data-stu-id="0a781-161">getEntireRow()</span></span>
* <span data-ttu-id="0a781-162">getLastCell()</span><span class="sxs-lookup"><span data-stu-id="0a781-162">getLastCell()</span></span>
* <span data-ttu-id="0a781-163">getLastColumn()</span><span class="sxs-lookup"><span data-stu-id="0a781-163">getLastColumn()</span></span>
* <span data-ttu-id="0a781-164">getLastRow()</span><span class="sxs-lookup"><span data-stu-id="0a781-164">getLastRow()</span></span>
* <span data-ttu-id="0a781-165">getRow (row)</span><span class="sxs-lookup"><span data-stu-id="0a781-165">getRow(row)</span></span>
* <span data-ttu-id="0a781-166">getRowHidden()</span><span class="sxs-lookup"><span data-stu-id="0a781-166">getRowHidden()</span></span>
* <span data-ttu-id="0a781-167">getRowsAbove (count)</span><span class="sxs-lookup"><span data-stu-id="0a781-167">getRowsAbove(count)</span></span>
* <span data-ttu-id="0a781-168">getRowsBelow (count)</span><span class="sxs-lookup"><span data-stu-id="0a781-168">getRowsBelow(count)</span></span>

<span data-ttu-id="0a781-169">**Важные и интересные**</span><span class="sxs-lookup"><span data-stu-id="0a781-169">**Important/Interesting**</span></span>

* <span data-ttu-id="0a781-170">_книга_.getSelectedRange()</span><span class="sxs-lookup"><span data-stu-id="0a781-170">_workbook_.getSelectedRange()</span></span>
* <span data-ttu-id="0a781-171">_книга_.getActiveCell()</span><span class="sxs-lookup"><span data-stu-id="0a781-171">_workbook_.getActiveCell()</span></span>
* <span data-ttu-id="0a781-172">getUsedRange (valuesOnly)</span><span class="sxs-lookup"><span data-stu-id="0a781-172">getUsedRange(valuesOnly)</span></span>
* <span data-ttu-id="0a781-173">getAbsoluteResizedRange (numRows, numColumns)</span><span class="sxs-lookup"><span data-stu-id="0a781-173">getAbsoluteResizedRange(numRows, numColumns)</span></span>
* <span data-ttu-id="0a781-174">getOffsetRange (rowOffset, columnOffset)</span><span class="sxs-lookup"><span data-stu-id="0a781-174">getOffsetRange(rowOffset, columnOffset)</span></span>
* <span data-ttu-id="0a781-175">getResizedRange (deltaRows, deltaColumns)</span><span class="sxs-lookup"><span data-stu-id="0a781-175">getResizedRange(deltaRows, deltaColumns)</span></span>

## <a name="methods-that-return-a-range-object-in-relation-to-another-range-object"></a><span data-ttu-id="0a781-176">Методы, возвращая объект диапазона по отношению к другому объекту диапазона</span><span class="sxs-lookup"><span data-stu-id="0a781-176">Methods that return a range object in relation to another range object</span></span>

* <span data-ttu-id="0a781-177">getBoundingRect (anotherRange)</span><span class="sxs-lookup"><span data-stu-id="0a781-177">getBoundingRect(anotherRange)</span></span>
* <span data-ttu-id="0a781-178">getIntersection (anotherRange)</span><span class="sxs-lookup"><span data-stu-id="0a781-178">getIntersection(anotherRange)</span></span>

## <a name="methods-that-return-other-objects-non-range-objects"></a><span data-ttu-id="0a781-179">Методы, возвращая другие объекты (не в диапазоне объектов)</span><span class="sxs-lookup"><span data-stu-id="0a781-179">Methods that return other objects (non-range objects)</span></span>

* <span data-ttu-id="0a781-180">getDirectPrecedents()</span><span class="sxs-lookup"><span data-stu-id="0a781-180">getDirectPrecedents()</span></span>
* <span data-ttu-id="0a781-181">getWorksheet()</span><span class="sxs-lookup"><span data-stu-id="0a781-181">getWorksheet()</span></span>
* <span data-ttu-id="0a781-182">getTables(fullyContained)</span><span class="sxs-lookup"><span data-stu-id="0a781-182">getTables(fullyContained)</span></span>
* <span data-ttu-id="0a781-183">getPivotTables(fullyContained)</span><span class="sxs-lookup"><span data-stu-id="0a781-183">getPivotTables(fullyContained)</span></span>
* <span data-ttu-id="0a781-184">getDataValidation()</span><span class="sxs-lookup"><span data-stu-id="0a781-184">getDataValidation()</span></span>
* <span data-ttu-id="0a781-185">getPredefinedCellStyle()</span><span class="sxs-lookup"><span data-stu-id="0a781-185">getPredefinedCellStyle()</span></span>

## <a name="set-methods"></a><span data-ttu-id="0a781-186">Настройка методов</span><span class="sxs-lookup"><span data-stu-id="0a781-186">Set methods</span></span>

### <a name="singular-cell-set-methods"></a><span data-ttu-id="0a781-187">Методы набора сингулярных клеток</span><span class="sxs-lookup"><span data-stu-id="0a781-187">Singular cell set methods</span></span>

* <span data-ttu-id="0a781-188">setFormula(formula)</span><span class="sxs-lookup"><span data-stu-id="0a781-188">setFormula(formula)</span></span>
* <span data-ttu-id="0a781-189">setFormulaLocal (formulaLocal)</span><span class="sxs-lookup"><span data-stu-id="0a781-189">setFormulaLocal(formulaLocal)</span></span>
* <span data-ttu-id="0a781-190">setFormulaR1C1 (formulaR1C1)</span><span class="sxs-lookup"><span data-stu-id="0a781-190">setFormulaR1C1(formulaR1C1)</span></span>
* <span data-ttu-id="0a781-191">setNumberFormatLocal (numberFormatLocal)</span><span class="sxs-lookup"><span data-stu-id="0a781-191">setNumberFormatLocal(numberFormatLocal)</span></span>
* <span data-ttu-id="0a781-192">setValue(value)</span><span class="sxs-lookup"><span data-stu-id="0a781-192">setValue(value)</span></span>

### <a name="2d--entire-range-set-methods"></a><span data-ttu-id="0a781-193">Методы набора диапазонов 2D /всего диапазона</span><span class="sxs-lookup"><span data-stu-id="0a781-193">2D / entire range set methods</span></span>

* <span data-ttu-id="0a781-194">setFormulas(formulas)</span><span class="sxs-lookup"><span data-stu-id="0a781-194">setFormulas(formulas)</span></span>
* <span data-ttu-id="0a781-195">setFormulasLocal (formulasLocal)</span><span class="sxs-lookup"><span data-stu-id="0a781-195">setFormulasLocal(formulasLocal)</span></span>
* <span data-ttu-id="0a781-196">setFormulasR1C1 (formulasR1C1)</span><span class="sxs-lookup"><span data-stu-id="0a781-196">setFormulasR1C1(formulasR1C1)</span></span>
* <span data-ttu-id="0a781-197">setNumberFormat (numberFormat)</span><span class="sxs-lookup"><span data-stu-id="0a781-197">setNumberFormat(numberFormat)</span></span>
* <span data-ttu-id="0a781-198">setNumberFormats (numberFormats)</span><span class="sxs-lookup"><span data-stu-id="0a781-198">setNumberFormats(numberFormats)</span></span>
* <span data-ttu-id="0a781-199">setNumberFormatsLocal (numberFormatsLocal)</span><span class="sxs-lookup"><span data-stu-id="0a781-199">setNumberFormatsLocal(numberFormatsLocal)</span></span>
* <span data-ttu-id="0a781-200">setValues(values)</span><span class="sxs-lookup"><span data-stu-id="0a781-200">setValues(values)</span></span>

## <a name="other-methods"></a><span data-ttu-id="0a781-201">Другие методы</span><span class="sxs-lookup"><span data-stu-id="0a781-201">Other methods</span></span>

* <span data-ttu-id="0a781-202">слияние (поперек)</span><span class="sxs-lookup"><span data-stu-id="0a781-202">merge(across)</span></span>
* <span data-ttu-id="0a781-203">unmerge()</span><span class="sxs-lookup"><span data-stu-id="0a781-203">unmerge()</span></span>

## <a name="coming-soon"></a><span data-ttu-id="0a781-204">Скоро</span><span class="sxs-lookup"><span data-stu-id="0a781-204">Coming soon</span></span>

* <span data-ttu-id="0a781-205">API края диапазона</span><span class="sxs-lookup"><span data-stu-id="0a781-205">Range edge APIs</span></span>
