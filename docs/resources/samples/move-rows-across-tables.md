---
title: Перемещение строк по таблицам с Office скриптами
description: Узнайте, как перемещать строки по таблицам, экономя фильтры, а затем обрабатывая и повторно примыкая к фильтрам.
ms.date: 06/29/2021
localization_priority: Normal
ms.openlocfilehash: 860521de166108d5a8355ea246c1bfe77e0e064b
ms.sourcegitcommit: 211c157ca746e266eeb079f5fa1925a1e35ab702
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/07/2021
ms.locfileid: "53313759"
---
# <a name="move-rows-across-tables-by-saving-filters-then-processing-and-reapplying-the-filters"></a><span data-ttu-id="3f783-103">Перемещение строк по таблицам с сохранением фильтров, а затем обработкой и повторной обработкой фильтров</span><span class="sxs-lookup"><span data-stu-id="3f783-103">Move rows across tables by saving filters, then processing and reapplying the filters</span></span>

<span data-ttu-id="3f783-104">В этом скрипте делается следующее:</span><span class="sxs-lookup"><span data-stu-id="3f783-104">This script does the following:</span></span>

* <span data-ttu-id="3f783-105">Выбирает строки из первой таблицы, где значение в столбце равно _некоторому значению._</span><span class="sxs-lookup"><span data-stu-id="3f783-105">Selects rows from the source table where the value in a column is equal to _some value_.</span></span>
* <span data-ttu-id="3f783-106">Перемещает все выбранные строки в другую (целевую) таблицу на другой таблице.</span><span class="sxs-lookup"><span data-stu-id="3f783-106">Moves all selected rows into another (target) table on another worksheet.</span></span>
* <span data-ttu-id="3f783-107">Повторное повторное появление соответствующих фильтров в исходных таблицах.</span><span class="sxs-lookup"><span data-stu-id="3f783-107">Reapplies the relevant filters on the source table.</span></span>

:::image type="content" source="../../images/table-filter-before-after.png" alt-text="Скриншоты книги до и после.":::

## <a name="sample-excel-file"></a><span data-ttu-id="3f783-109">Пример Excel файла</span><span class="sxs-lookup"><span data-stu-id="3f783-109">Sample Excel file</span></span>

<span data-ttu-id="3f783-110">Скачайте <a href="input-table-filters.xlsx"> файлinput-table-filters.xlsx</a> для готовой к использованию книги.</span><span class="sxs-lookup"><span data-stu-id="3f783-110">Download the file <a href="input-table-filters.xlsx">input-table-filters.xlsx</a> for a ready-to-use workbook.</span></span> <span data-ttu-id="3f783-111">Добавьте следующий скрипт, чтобы попробовать пример самостоятельно!</span><span class="sxs-lookup"><span data-stu-id="3f783-111">Add the following script to try the sample yourself!</span></span>

## <a name="sample-code-move-rows-using-range-values"></a><span data-ttu-id="3f783-112">Пример кода. Перемещение строк с использованием значений диапазона</span><span class="sxs-lookup"><span data-stu-id="3f783-112">Sample code: Move rows using range values</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {

  // You can change these names to match the data in your workbook.
  const TargetTableName = 'Table1';
  const SourceTableName = 'Table2';
  const IndexOfColumnToFilterOn = 1;
  const NameOfColumnToFilterOn = 'Category';
  const ValueToFilterOn = 'Clothing';

  // Get the Table objects.
  let targetTable = workbook.getTable(TargetTableName);
  let sourceTable = workbook.getTable(SourceTableName);

  // If either table is missing, report that information and stop the script.
  if (!targetTable || !sourceTable) {
    console.log(`Tables missing - Check to make sure both source (${TargetTableName}) and target table (${SourceTableName}) are present before running the script. `);
    return;
  }

  // Save the filter criteria.
  const tableFilters = {};
  // For each table column, collect the filter criteria on that column.
  sourceTable.getColumns().forEach((column) => {
    let colFilterCriteria = column.getFilter().getCriteria();
    if (colFilterCriteria) {
      tableFilters[column.getName()] = colFilterCriteria;
    }
  });

  // Get all the data from the table.
  const sourceRange = sourceTable.getRangeBetweenHeaderAndTotal();
  const dataRows: (number | string | boolean)[][] = sourceTable.getRangeBetweenHeaderAndTotal().getValues();

  // Create variables to hold the rows to be moved and their addresses.
  let rowsToMoveValues: (number | string | boolean)[][] = [];
  let rowAddressToRemove: string[] = [];

  // Get the data values from the source table.
  for (let i = 0; i < dataRows.length; i++) { 
    if (dataRows[i][IndexOfColumnToFilterOn] === ValueToFilterOn) {
      rowsToMoveValues.push(dataRows[i]);

      // Get the intersection between table address and the entire row where we found the match. This provides the address of the range to remove.
      let address = sourceRange.getIntersection(sourceRange.getCell(i,0).getEntireRow()).getAddress();
      rowAddressToRemove.push(address);
    }
  }

  // If there are no data rows to process, end the script.
  if (rowsToMoveValues.length < 1) {
    console.log('No rows selected from the source table match the filter criteria.');
    return;
  }

  console.log(`Adding ${rowsToMoveValues.length} rows to target table.`);

  // Insert rows at the end of target table.
  targetTable.addRows(-1, rowsToMoveValues)

  // Remove the rows from the source table.
  const sheet = sourceTable.getWorksheet();

  // Remove all filters before removing rows.
  sourceTable.getAutoFilter().clearCriteria();

  // Important: Remove the rows starting at the bottom of the table.
  // Otherwise, the lower rows change position before they are deleted.
  console.log(`Removing ${rowAddressToRemove.length} rows from the source table.`);
  rowAddressToRemove.reverse().forEach((address) => {
    sheet.getRange(address).delete(ExcelScript.DeleteShiftDirection.up);
  });

  // Reapply the original filters. 
  Object.keys(tableFilters).forEach((columnName) => {
      sourceTable.getColumnByName(columnName).getFilter().apply(tableFilters[columnName]);
    });
}
```

## <a name="training-video-move-rows-across-tables"></a><span data-ttu-id="3f783-113">Обучающее видео: перемещение строк по таблицам</span><span class="sxs-lookup"><span data-stu-id="3f783-113">Training video: Move rows across tables</span></span>

<span data-ttu-id="3f783-114">[Смотреть Sudhi Ramamurthy ходить через этот пример на YouTube](https://youtu.be/_3t3Pk4i2L0).</span><span class="sxs-lookup"><span data-stu-id="3f783-114">[Watch Sudhi Ramamurthy walk through this sample on YouTube](https://youtu.be/_3t3Pk4i2L0).</span></span> <span data-ttu-id="3f783-115">В решении видео показано два сценария.</span><span class="sxs-lookup"><span data-stu-id="3f783-115">There are two scripts shown in the video's solution.</span></span> <span data-ttu-id="3f783-116">Основное отличие состоит в выборе строк.</span><span class="sxs-lookup"><span data-stu-id="3f783-116">The main difference is how the rows are selected.</span></span>

* <span data-ttu-id="3f783-117">В первом варианте строки выбираются путем применения фильтра таблицы и чтения видимого диапазона.</span><span class="sxs-lookup"><span data-stu-id="3f783-117">In the first variant, the rows are selected by applying the table filter and reading the visible range.</span></span>
* <span data-ttu-id="3f783-118">Во втором случае строки выбираются путем чтения значений и извлечения значений строк (что используется в примере на этой странице).</span><span class="sxs-lookup"><span data-stu-id="3f783-118">In the second, the rows are selected by reading the values and extracting the row values (which is what the sample on this page uses).</span></span>
