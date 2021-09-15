---
title: Фильтр четких столбцов таблицы на основе расположения активных клеток
description: Узнайте, как очистить фильтр столбца таблицы на основе активного расположения ячейки.
ms.date: 06/29/2021
ms.localizationpriority: medium
ms.openlocfilehash: bb01292017a027e41230d786337b5bf53293a20c
ms.sourcegitcommit: d3ed4bdeeba805d97c930394e172e8306a0cf484
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/15/2021
ms.locfileid: "59332987"
---
# <a name="clear-table-column-filter-based-on-active-cell-location"></a>Фильтр четких столбцов таблицы на основе расположения активных клеток

В этом примере фильтр столбца таблицы очищается в зависимости от расположения активной ячейки. Скрипт определяет, является ли ячейка частью таблицы, определяет столбец таблицы и очищает фильтр, применяемый на ней.

Если вы хотите узнать больше о том, как сохранить фильтр до его очистки (и повторно применить позже), см. в таблице [Перемещение](move-rows-across-tables.md)строк по таблицам путем сохранения фильтров , более расширенный пример.

_Перед очисткой фильтра столбца (обратите внимание на активную ячейку)_

:::image type="content" source="../../images/before-filter-applied.png" alt-text="Активная ячейка перед очисткой фильтра столбца.":::

_После очистки фильтра столбца_

:::image type="content" source="../../images/after-filter-cleared.png" alt-text="Активная ячейка после очистки фильтра столбца.":::

## <a name="sample-excel-file"></a>Пример Excel файла

Скачайте <a href="table-with-filter.xlsx">table-with-filter.xlsx</a> для готовой к использованию книги. Добавьте следующий скрипт, чтобы попробовать пример самостоятельно!

## <a name="sample-code-clear-table-column-filter-based-on-active-cell"></a>Пример кода: фильтр столбцов ясной таблицы на основе активной ячейки

Следующий сценарий очищает фильтр столбца таблицы на основе расположения активных ячеь и может применяться к любому Excel с таблицей.

```TypeScript
function main(workbook: ExcelScript.Workbook) {
    // Get the active cell.
    const cell = workbook.getActiveCell();

    // Get all tables associated with that cell.
    const tables = cell.getTables();
    
    // If there is no table on the selection, end the script.
    if (tables.length !== 1) {
      console.log("The selection is not in a table.");
      return;
    }

    // Get the first table associated with the active cell.
    const currentTable = tables[0];

    // Log key information about the table.
    console.log(currentTable.getName());
    console.log(currentTable.getRange().getAddress());

    // Get the table header above the current cell by referencing its column.
    const entireColumn = cell.getEntireColumn();
    const intersect = entireColumn.getIntersection(currentTable.getRange());
    console.log(intersect.getAddress());

    const headerCellValue = intersect.getCell(0,0).getValue() as string;
    console.log(headerCellValue);

    // Get the TableColumn object matching that header.
    const tableColumn = currentTable.getColumnByName(headerCellValue);

    // Clear the filter on that table column.
    tableColumn.getFilter().clear();
}
```
