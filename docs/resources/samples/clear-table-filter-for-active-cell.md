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
# <a name="clear-table-column-filter-based-on-active-cell-location"></a><span data-ttu-id="c03ab-103">Фильтр четких столбцов таблицы на основе расположения активных клеток</span><span class="sxs-lookup"><span data-stu-id="c03ab-103">Clear table column filter based on active cell location</span></span>

<span data-ttu-id="c03ab-104">В этом примере фильтр столбца таблицы очищается в зависимости от расположения активной ячейки.</span><span class="sxs-lookup"><span data-stu-id="c03ab-104">This sample clears the table column filter based on the active cell location.</span></span> <span data-ttu-id="c03ab-105">Скрипт определяет, является ли ячейка частью таблицы, определяет столбец таблицы и очищает фильтр, применяемый на ней.</span><span class="sxs-lookup"><span data-stu-id="c03ab-105">The script detects if the cell is part of a table, determines the table column, and clears any filter that are applied on it.</span></span>

<span data-ttu-id="c03ab-106">Если вы хотите узнать больше о том, как сохранить фильтр до его очистки (и повторно применить позже), см. в таблице [Перемещение](move-rows-across-tables.md)строк по таблицам путем сохранения фильтров , более расширенный пример.</span><span class="sxs-lookup"><span data-stu-id="c03ab-106">If you wish to learn more about how to save the filter prior to clearing it (and re-apply later), see [Move rows across tables by saving filters](move-rows-across-tables.md), a more advanced sample.</span></span>

<span data-ttu-id="c03ab-107">_Перед очисткой фильтра столбца (обратите внимание на активную ячейку)_</span><span class="sxs-lookup"><span data-stu-id="c03ab-107">_Before clearing column filter (notice the active cell)_</span></span>

:::image type="content" source="../../images/before-filter-applied.png" alt-text="Активная ячейка перед очисткой фильтра столбца":::

<span data-ttu-id="c03ab-109">_После очистки фильтра столбца_</span><span class="sxs-lookup"><span data-stu-id="c03ab-109">_After clearing column filter_</span></span>

:::image type="content" source="../../images/after-filter-cleared.png" alt-text="Активная ячейка после очистки фильтра столбца":::

## <a name="sample-code-clear-table-column-filter-based-on-active-cell"></a><span data-ttu-id="c03ab-111">Пример кода: фильтр столбцов ясной таблицы на основе активной ячейки</span><span class="sxs-lookup"><span data-stu-id="c03ab-111">Sample code: Clear table column filter based on active cell</span></span>

<span data-ttu-id="c03ab-112">Следующий сценарий очищает фильтр столбца таблицы на основе расположения активных ячеь и может применяться к любому Excel с таблицей.</span><span class="sxs-lookup"><span data-stu-id="c03ab-112">The following script clears the table column filter based on active cell location and can be applied to any Excel file with a table.</span></span> <span data-ttu-id="c03ab-113">Для удобства можно скачать и использовать <a href="table-with-filter.xlsx">table-with-filter.xlsx. </a></span><span class="sxs-lookup"><span data-stu-id="c03ab-113">For convenience, you can download and use <a href="table-with-filter.xlsx">table-with-filter.xlsx</a>.</span></span>

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
