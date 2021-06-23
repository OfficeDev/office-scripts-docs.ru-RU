---
title: Фильтр четких столбцов таблицы на основе расположения активных клеток
description: Узнайте, как очистить фильтр столбца таблицы на основе активного расположения ячейки.
ms.date: 05/03/2021
localization_priority: Normal
ms.openlocfilehash: d6f267b433be9a0ddf44edf53ed92a136eb2ded6
ms.sourcegitcommit: 4693c8f79428ec74695328275703af0ba1bfea8f
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 06/23/2021
ms.locfileid: "53074440"
---
# <a name="clear-table-column-filter-based-on-active-cell-location"></a><span data-ttu-id="21576-103">Фильтр четких столбцов таблицы на основе расположения активных клеток</span><span class="sxs-lookup"><span data-stu-id="21576-103">Clear table column filter based on active cell location</span></span>

<span data-ttu-id="21576-104">В этом примере фильтр столбца таблицы очищается в зависимости от расположения активной ячейки.</span><span class="sxs-lookup"><span data-stu-id="21576-104">This sample clears the table column filter based on the active cell location.</span></span> <span data-ttu-id="21576-105">Скрипт определяет, является ли ячейка частью таблицы, определяет столбец таблицы и очищает фильтр, применяемый на ней.</span><span class="sxs-lookup"><span data-stu-id="21576-105">The script detects if the cell is part of a table, determines the table column, and clears any filter that are applied on it.</span></span>

<span data-ttu-id="21576-106">Если вы хотите узнать больше о том, как сохранить фильтр до его очистки (и повторно применить позже), см. в таблице [Перемещение](move-rows-across-tables.md)строк по таблицам путем сохранения фильтров , более расширенный пример.</span><span class="sxs-lookup"><span data-stu-id="21576-106">If you wish to learn more about how to save the filter prior to clearing it (and re-apply later), see [Move rows across tables by saving filters](move-rows-across-tables.md), a more advanced sample.</span></span>

<span data-ttu-id="21576-107">_Перед очисткой фильтра столбца (обратите внимание на активную ячейку)_</span><span class="sxs-lookup"><span data-stu-id="21576-107">_Before clearing column filter (notice the active cell)_</span></span>

:::image type="content" source="../../images/before-filter-applied.png" alt-text="Активная ячейка перед очисткой фильтра столбца.":::

<span data-ttu-id="21576-109">_После очистки фильтра столбца_</span><span class="sxs-lookup"><span data-stu-id="21576-109">_After clearing column filter_</span></span>

:::image type="content" source="../../images/after-filter-cleared.png" alt-text="Активная ячейка после очистки фильтра столбца.":::

## <a name="sample-code-clear-table-column-filter-based-on-active-cell"></a><span data-ttu-id="21576-111">Пример кода: фильтр столбцов ясной таблицы на основе активной ячейки</span><span class="sxs-lookup"><span data-stu-id="21576-111">Sample code: Clear table column filter based on active cell</span></span>

<span data-ttu-id="21576-112">Следующий сценарий очищает фильтр столбца таблицы на основе расположения активных ячеь и может применяться к любому Excel с таблицей.</span><span class="sxs-lookup"><span data-stu-id="21576-112">The following script clears the table column filter based on active cell location and can be applied to any Excel file with a table.</span></span> <span data-ttu-id="21576-113">Для удобства можно скачать и использовать <a href="table-with-filter.xlsx">table-with-filter.xlsx. </a></span><span class="sxs-lookup"><span data-stu-id="21576-113">For convenience, you can download and use <a href="table-with-filter.xlsx">table-with-filter.xlsx</a>.</span></span>

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
