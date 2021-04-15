---
title: Фильтр четких столбцов таблицы на основе расположения активных клеток
description: Узнайте, как очистить фильтр столбца таблицы на основе активного расположения ячейки.
ms.date: 03/04/2021
localization_priority: Normal
ms.openlocfilehash: 4f8353fb5480812b7b63e7a9b3ffb11ece2a8c6c
ms.sourcegitcommit: 45ffe3dbd2c834b78592ad35928cf8096f5e80bc
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 04/14/2021
ms.locfileid: "51755086"
---
# <a name="clear-table-column-filter-based-on-active-cell-location"></a>Фильтр четких столбцов таблицы на основе расположения активных клеток

В этом примере фильтр столбца таблицы очищается в зависимости от расположения активной ячейки. Скрипт определяет, является ли ячейка частью таблицы, определяет столбец таблицы и очищает фильтр, применяемый на ней.

Если вы хотите узнать больше о том, как сохранить фильтр до его очистки (и повторно применить позже), см. в таблице [Перемещение](move-rows-across-tables.md)строк по таблицам путем сохранения фильтров , более расширенный пример.

_Перед очисткой фильтра столбца (обратите внимание на активную ячейку)_

:::image type="content" source="../../images/before-filter-applied.png" alt-text="Активная ячейка перед очисткой фильтра столбца.":::

_После очистки фильтра столбца_

:::image type="content" source="../../images/after-filter-cleared.png" alt-text="Активная ячейка после очистки фильтра столбца.":::

## <a name="sample-code-clear-table-column-filter-based-on-active-cell"></a>Пример кода: фильтр столбцов ясной таблицы на основе активной ячейки

Следующий сценарий очищает фильтр столбца таблицы в зависимости от расположения активных клеток и может применяться к любому файлу Excel со таблицей. Для удобства можно скачать и использовать <a href="table-with-filter.xlsx">table-with-filter.xlsx. </a>

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

## <a name="training-video-clear-table-column-filter-based-on-active-cell-location"></a>Обучающее видео: фильтр столбцов clear table based on active cell location

Пример работы с диапазонами см. в примере обучающих видео [range basics.](range-basics.md#training-videos-range-basics)
