---
title: Перемещение строк по таблицам с Office скриптами
description: Узнайте, как перемещать строки по таблицам, экономя фильтры, а затем обрабатывая и повторно примыкая к фильтрам.
ms.date: 06/29/2021
ms.localizationpriority: medium
ms.openlocfilehash: bffdb17516016d159e61586c116d764f7bb8f3fc
ms.sourcegitcommit: d3ed4bdeeba805d97c930394e172e8306a0cf484
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/15/2021
ms.locfileid: "59334971"
---
# <a name="move-rows-across-tables"></a>Перемещение строк по таблицам

В этом скрипте делается следующее:

* Выбирает строки из первой таблицы, где значение в столбце равно некоторому значению `FILTER_VALUE` (в скрипте).
* Перемещает все выбранные строки в целевую таблицу в другой таблице.
* Повторно передает соответствующие фильтры в исходный стол.

## <a name="sample-excel-file"></a>Пример Excel файла

Скачайте <a href="input-table-filters.xlsx"> файлinput-table-filters.xlsx</a> для готовой к использованию книги. Добавьте следующий скрипт, чтобы попробовать пример самостоятельно!

## <a name="sample-code-move-rows-using-range-values"></a>Пример кода. Перемещение строк с использованием значений диапазона

```TypeScript
function main(workbook: ExcelScript.Workbook) {

  // You can change these names to match the data in your workbook.
  const TARGET_TABLE_NAME = 'Table1';
  const SOURCE_TABLE_NAME = 'Table2';

  // Select what will be moved between tables.
  const FILTER_COLUMN_INDEX = 1;
  const FILTER_VALUE = 'Clothing';

  // Get the Table objects.
  let targetTable = workbook.getTable(TARGET_TABLE_NAME);
  let sourceTable = workbook.getTable(SOURCE_TABLE_NAME);

  // If either table is missing, report that information and stop the script.
  if (!targetTable || !sourceTable) {
    console.log(`Tables missing - Check to make sure both source (${TARGET_TABLE_NAME}) and target table (${SOURCE_TABLE_NAME}) are present before running the script. `);
    return;
  }

  // Save the filter criteria currently on the source table.
  const originalTableFilters = {};
  // For each table column, collect the filter criteria on that column.
  sourceTable.getColumns().forEach((column) => {
    let originalColumnFilter = column.getFilter().getCriteria();
    if (originalColumnFilter) {
      originalTableFilters[column.getName()] = originalColumnFilter;
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
    if (dataRows[i][FILTER_COLUMN_INDEX] === FILTER_VALUE) {
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
  Object.keys(originalTableFilters).forEach((columnName) => {
      sourceTable.getColumnByName(columnName).getFilter().apply(originalTableFilters[columnName]);
    });
}
```

## <a name="training-video-move-rows-across-tables"></a>Обучающее видео: перемещение строк по таблицам

[Смотреть Sudhi Ramamurthy ходить через этот пример на YouTube](https://youtu.be/_3t3Pk4i2L0). В решении видео показано два сценария. Основное отличие состоит в выборе строк.

* В первом варианте строки выбираются путем применения фильтра таблицы и чтения видимого диапазона.
* Во втором случае строки выбираются путем чтения значений и извлечения значений строк (что используется в примере на этой странице).
