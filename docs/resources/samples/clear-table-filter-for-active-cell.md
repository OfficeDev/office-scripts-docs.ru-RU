---
title: Удаление фильтров столбцов таблицы
description: Узнайте, как очистить фильтр столбцов таблицы на основе активного расположения ячейки.
ms.date: 07/15/2022
ms.localizationpriority: medium
ms.openlocfilehash: 21a79abfdd4aeac79af4a0f9ea4a581d45b9706b
ms.sourcegitcommit: dd632402cb46ec8407a1c98456f1bc9ab96ffa46
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/21/2022
ms.locfileid: "66918813"
---
# <a name="remove-table-column-filters"></a>Удаление фильтров столбцов таблицы

В этом примере фильтры удаляются из столбца таблицы в зависимости от расположения активной ячейки. Скрипт определяет, является ли ячейка частью таблицы, определяет столбец таблицы и очищает все примененные к ней фильтры.

Если вы хотите узнать больше о том, как сохранить фильтр перед его очисткой (и повторно применить позже), см. раздел "Перемещение строк между таблицами путем сохранения фильтров [", более](move-rows-across-tables.md) сложный пример.

## <a name="sample-excel-file"></a>Пример файла Excel

<a href="table-with-filter.xlsx"> Скачайтеtable-with-filter.xlsx</a> для готовой к использованию книги. Добавьте следующий скрипт, чтобы попробовать пример самостоятельно!

## <a name="sample-code-clear-table-column-filter-based-on-active-cell"></a>Пример кода: очистка фильтра столбцов таблицы на основе активной ячейки

Следующий сценарий очищает фильтр столбцов таблицы на основе активного расположения ячейки и может применяться к любому файлу Excel с таблицей.

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Get the active cell.
  const cell = workbook.getActiveCell();

  // Get the tables associated with that cell.
  // Since tables can't overlap, this will be one table at most.
  const currentTable = cell.getTables()[0];

  // If there is no table on the selection, end the script.
  if (!currentTable) {
    console.log("The selection is not in a table.");
    return;
  }

  // Get the table header above the current cell by referencing its column.
  const entireColumn = cell.getEntireColumn();
  const intersect = entireColumn.getIntersection(currentTable.getRange());
  const headerCellValue = intersect.getCell(0, 0).getValue() as string;

  // Get the TableColumn object matching that header.
  const tableColumn = currentTable.getColumnByName(headerCellValue);

  // Clear the filters on that table column.
  tableColumn.getFilter().clear();
}
```

## <a name="before-clearing-column-filter-notice-the-active-cell"></a>Перед очисткой фильтра столбцов (обратите внимание на активную ячейку)

:::image type="content" source="../../images/before-filter-applied.png" alt-text="Активная ячейка перед очисткой фильтра столбцов.":::

## <a name="after-clearing-column-filter"></a>После очистки фильтра столбцов

:::image type="content" source="../../images/after-filter-cleared.png" alt-text="Активная ячейка после очистки фильтра столбцов.":::
