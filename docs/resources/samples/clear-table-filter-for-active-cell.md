---
title: Фильтр четких столбцов таблицы на основе расположения активных клеток
description: Узнайте, как очистить фильтр столбца таблицы на основе активного расположения ячейки.
ms.date: 03/04/2021
localization_priority: Normal
ms.openlocfilehash: bbca4adce1de2cfade2c4f84273bf0bc06b5cc4b
ms.sourcegitcommit: f7a7aebfb687f2a35dbed07ed62ff352a114525a
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 05/06/2021
ms.locfileid: "52232503"
---
# <a name="clear-table-column-filter-based-on-active-cell-location"></a>Фильтр четких столбцов таблицы на основе расположения активных клеток

В этом примере фильтр столбца таблицы очищается в зависимости от расположения активной ячейки. Скрипт определяет, является ли ячейка частью таблицы, определяет столбец таблицы и очищает фильтр, применяемый на ней.

Если вы хотите узнать больше о том, как сохранить фильтр до его очистки (и повторно применить позже), см. в таблице [Перемещение](move-rows-across-tables.md)строк по таблицам путем сохранения фильтров , более расширенный пример.

_Перед очисткой фильтра столбца (обратите внимание на активную ячейку)_

:::image type="content" source="../../images/before-filter-applied.png" alt-text="Активная ячейка перед очисткой фильтра столбца":::

_После очистки фильтра столбца_

:::image type="content" source="../../images/after-filter-cleared.png" alt-text="Активная ячейка после очистки фильтра столбца":::

## <a name="sample-code-clear-table-column-filter-based-on-active-cell"></a>Пример кода: фильтр столбцов ясной таблицы на основе активной ячейки

Следующий сценарий очищает фильтр столбца таблицы на основе расположения активных ячеь и может применяться к любому Excel с таблицей. Для удобства можно скачать и использовать <a href="table-with-filter.xlsx">table-with-filter.xlsx. </a>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
    // Get active cell.
    const cell = workbook.getActiveCell();

    // Get all tables associated with that cell.
    const tables = cell.getTables();
    
    // If there is no table on the selection, return/exit.
    if (tables.length !== 1) {
      console.log("The selection is not in a table.");
      return;
    }

    // Get table (since it is already determined that there is only
    // a single table part of the selection).
    const currentTable = tables[0];

    console.log(currentTable.getName());
    console.log(currentTable.getRange().getAddress());

    const entireCol = cell.getEntireColumn();
    const intersect = entireCol.getIntersection(currentTable.getRange());
    console.log(intersect.getAddress());

    const headerCellValue = intersect.getCell(0,0).getValue() as string;
    console.log(headerCellValue);

    // Get column.
    const col = currentTable.getColumnByName(headerCellValue);

    // Clear filter.
    col.getFilter().clear();
}
```
