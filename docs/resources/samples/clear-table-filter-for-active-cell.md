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
# <a name="clear-table-column-filter-based-on-active-cell-location"></a><span data-ttu-id="76c3f-103">Фильтр четких столбцов таблицы на основе расположения активных клеток</span><span class="sxs-lookup"><span data-stu-id="76c3f-103">Clear table column filter based on active cell location</span></span>

<span data-ttu-id="76c3f-104">В этом примере фильтр столбца таблицы очищается в зависимости от расположения активной ячейки.</span><span class="sxs-lookup"><span data-stu-id="76c3f-104">This sample clears the table column filter based on the active cell location.</span></span> <span data-ttu-id="76c3f-105">Скрипт определяет, является ли ячейка частью таблицы, определяет столбец таблицы и очищает фильтр, применяемый на ней.</span><span class="sxs-lookup"><span data-stu-id="76c3f-105">The script detects if the cell is part of a table, determines the table column, and clears any filter that are applied on it.</span></span>

<span data-ttu-id="76c3f-106">Если вы хотите узнать больше о том, как сохранить фильтр до его очистки (и повторно применить позже), см. в таблице [Перемещение](move-rows-across-tables.md)строк по таблицам путем сохранения фильтров , более расширенный пример.</span><span class="sxs-lookup"><span data-stu-id="76c3f-106">If you wish to learn more about how to save the filter prior to clearing it (and re-apply later), see [Move rows across tables by saving filters](move-rows-across-tables.md), a more advanced sample.</span></span>

<span data-ttu-id="76c3f-107">_Перед очисткой фильтра столбца (обратите внимание на активную ячейку)_</span><span class="sxs-lookup"><span data-stu-id="76c3f-107">_Before clearing column filter (notice the active cell)_</span></span>

:::image type="content" source="../../images/before-filter-applied.png" alt-text="Активная ячейка перед очисткой фильтра столбца.":::

<span data-ttu-id="76c3f-109">_После очистки фильтра столбца_</span><span class="sxs-lookup"><span data-stu-id="76c3f-109">_After clearing column filter_</span></span>

:::image type="content" source="../../images/after-filter-cleared.png" alt-text="Активная ячейка после очистки фильтра столбца.":::

## <a name="sample-code-clear-table-column-filter-based-on-active-cell"></a><span data-ttu-id="76c3f-111">Пример кода: фильтр столбцов ясной таблицы на основе активной ячейки</span><span class="sxs-lookup"><span data-stu-id="76c3f-111">Sample code: Clear table column filter based on active cell</span></span>

<span data-ttu-id="76c3f-112">Следующий сценарий очищает фильтр столбца таблицы в зависимости от расположения активных клеток и может применяться к любому файлу Excel со таблицей.</span><span class="sxs-lookup"><span data-stu-id="76c3f-112">The following script clears the table column filter based on active cell location and can be applied to any Excel file with a table.</span></span> <span data-ttu-id="76c3f-113">Для удобства можно скачать и использовать <a href="table-with-filter.xlsx">table-with-filter.xlsx. </a></span><span class="sxs-lookup"><span data-stu-id="76c3f-113">For convenience, you can download and use <a href="table-with-filter.xlsx">table-with-filter.xlsx</a>.</span></span>

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

## <a name="training-video-clear-table-column-filter-based-on-active-cell-location"></a><span data-ttu-id="76c3f-114">Обучающее видео: фильтр столбцов clear table based on active cell location</span><span class="sxs-lookup"><span data-stu-id="76c3f-114">Training video: Clear table column filter based on active cell location</span></span>

<span data-ttu-id="76c3f-115">Пример работы с диапазонами см. в примере обучающих видео [range basics.](range-basics.md#training-videos-range-basics)</span><span class="sxs-lookup"><span data-stu-id="76c3f-115">For an example of how to work with ranges, see [Range basics training videos](range-basics.md#training-videos-range-basics).</span></span>
