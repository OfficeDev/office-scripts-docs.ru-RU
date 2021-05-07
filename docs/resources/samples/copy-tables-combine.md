---
title: Объединяйте данные из нескольких Excel таблиц в одну таблицу
description: Узнайте, как использовать Office скрипты для объединения данных из нескольких Excel таблиц в одну таблицу.
ms.date: 04/28/2021
localization_priority: Normal
ms.openlocfilehash: ac8c7d0a3f0f4f3d7d3217ffac31aff1a5595d17
ms.sourcegitcommit: f7a7aebfb687f2a35dbed07ed62ff352a114525a
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 05/06/2021
ms.locfileid: "52232447"
---
# <a name="combine-data-from-multiple-excel-tables-into-a-single-table"></a><span data-ttu-id="0ae20-103">Объединяйте данные из нескольких Excel таблиц в одну таблицу</span><span class="sxs-lookup"><span data-stu-id="0ae20-103">Combine data from multiple Excel tables into a single table</span></span>

<span data-ttu-id="0ae20-104">Этот пример объединяет данные из нескольких Excel таблиц в одну таблицу, которая включает все строки.</span><span class="sxs-lookup"><span data-stu-id="0ae20-104">This sample combines data from multiple Excel tables into a single table that includes all the rows.</span></span> <span data-ttu-id="0ae20-105">Предполагается, что все используемые таблицы имеют ту же структуру.</span><span class="sxs-lookup"><span data-stu-id="0ae20-105">It assumes that all tables being used have the same structure.</span></span>

<span data-ttu-id="0ae20-106">Существует два варианта этого сценария:</span><span class="sxs-lookup"><span data-stu-id="0ae20-106">There are two variations of this script:</span></span>

1. <span data-ttu-id="0ae20-107">Первый [скрипт объединяет](#sample-code-combine-data-from-multiple-excel-tables-into-a-single-table) все таблицы в Excel файле.</span><span class="sxs-lookup"><span data-stu-id="0ae20-107">The [first script](#sample-code-combine-data-from-multiple-excel-tables-into-a-single-table) combines all tables in the Excel file.</span></span>
1. <span data-ttu-id="0ae20-108">Второй [сценарий](#sample-code-combine-data-from-multiple-excel-tables-in-select-worksheets-into-a-single-table) выборочно получает таблицы в наборе таблиц.</span><span class="sxs-lookup"><span data-stu-id="0ae20-108">The [second script](#sample-code-combine-data-from-multiple-excel-tables-in-select-worksheets-into-a-single-table) selectively gets tables within a set of worksheets.</span></span>

## <a name="sample-code-combine-data-from-multiple-excel-tables-into-a-single-table"></a><span data-ttu-id="0ae20-109">Пример кода. Объединяйте данные из нескольких Excel таблиц в одну таблицу</span><span class="sxs-lookup"><span data-stu-id="0ae20-109">Sample code: Combine data from multiple Excel tables into a single table</span></span>

<span data-ttu-id="0ae20-110">Скачайте пример файла <a href="tables-copy.xlsx">tables-copy.xlsx</a> и используйте его со следующим скриптом, чтобы попробовать его самостоятельно!</span><span class="sxs-lookup"><span data-stu-id="0ae20-110">Download the sample file <a href="tables-copy.xlsx">tables-copy.xlsx</a> and use it with the following script to try it out yourself!</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
    workbook.getWorksheet('Combined')?.delete();
    const newSheet = workbook.addWorksheet('Combined');
    
    const tables = workbook.getTables();    
    const headerValues = tables[0].getHeaderRowRange().getTexts();
    console.log(headerValues);
    const targetRange = updateRange(newSheet, headerValues);
    const combinedTable = newSheet.addTable(targetRange.getAddress(), true);
    for (let table of tables) {      
      let dataValues = table.getRangeBetweenHeaderAndTotal().getTexts();
      let rowCount = table.getRowCount();
      if (rowCount > 0) {
        combinedTable.addRows(-1, dataValues);
      }
    }
}

function updateRange(sheet: ExcelScript.Worksheet, data: string[][]): ExcelScript.Range {
  const targetRange = sheet.getRange('A1').getResizedRange(data.length-1, data[0].length-1);
  targetRange.setValues(data);
  return targetRange;
}
```

## <a name="sample-code-combine-data-from-multiple-excel-tables-in-select-worksheets-into-a-single-table"></a><span data-ttu-id="0ae20-111">Пример кода. Объединяйте данные из нескольких Excel таблиц в отдельных таблицах в одну таблицу</span><span class="sxs-lookup"><span data-stu-id="0ae20-111">Sample code: Combine data from multiple Excel tables in select worksheets into a single table</span></span>

<span data-ttu-id="0ae20-112">Скачайте пример файла <a href="tables-select-copy.xlsx">tables-select-copy.xlsx</a> и используйте его со следующим скриптом, чтобы попробовать его самостоятельно!</span><span class="sxs-lookup"><span data-stu-id="0ae20-112">Download the sample file <a href="tables-select-copy.xlsx">tables-select-copy.xlsx</a> and use it with the following script to try it out yourself!</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
    const sheetNames = ['Sheet1', 'Sheet2', 'Sheet3'];
    
    workbook.getWorksheet('Combined')?.delete();
    const newSheet = workbook.addWorksheet('Combined');
    let targetTableCreated = false;
    let combinedTable;
    sheetNames.forEach((sheet) => {
      const tables = workbook.getWorksheet(sheet).getTables();
      if (!targetTableCreated) {
        const headerValues = tables[0].getHeaderRowRange().getTexts();
        const targetRange = updateRange(newSheet, headerValues);
        combinedTable = newSheet.addTable(targetRange.getAddress(), true);
        targetTableCreated = true;
      }      
      for (let table of tables) {
        let dataValues = table.getRangeBetweenHeaderAndTotal().getTexts();
        let rowCount = table.getRowCount();
        if (rowCount > 0) {
        combinedTable.addRows(-1, dataValues);
        }
      }
    })
}

function updateRange(sheet: ExcelScript.Worksheet, data: string[][]): ExcelScript.Range {
  const targetRange = sheet.getRange('A1').getResizedRange(data.length-1, data[0].length-1);
  targetRange.setValues(data);
  return targetRange;
}
```

## <a name="training-video-combine-data-from-multiple-excel-tables-into-a-single-table"></a><span data-ttu-id="0ae20-113">Обучающее видео. Объединяйте данные из нескольких Excel таблиц в одну таблицу</span><span class="sxs-lookup"><span data-stu-id="0ae20-113">Training video: Combine data from multiple Excel tables into a single table</span></span>

<span data-ttu-id="0ae20-114">[Смотреть Sudhi Ramamurthy ходить через этот пример на YouTube](https://youtu.be/di-8JukK3Lc).</span><span class="sxs-lookup"><span data-stu-id="0ae20-114">[Watch Sudhi Ramamurthy walk through this sample on YouTube](https://youtu.be/di-8JukK3Lc).</span></span>
