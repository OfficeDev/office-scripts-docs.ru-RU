---
title: Объединяйте данные из нескольких таблиц Excel в одну таблицу
description: Узнайте, как использовать скрипты Office для объединения данных из нескольких таблиц Excel в одну таблицу.
ms.date: 03/30/2021
localization_priority: Normal
ms.openlocfilehash: 2f3f7232216f686946861d8c2cdec44013333ec7
ms.sourcegitcommit: 5d24e77df70aa2c1c982275d53213c2a9323ff86
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 04/02/2021
ms.locfileid: "51571416"
---
# <a name="combine-data-from-multiple-excel-tables-into-a-single-table"></a><span data-ttu-id="cf636-103">Объединяйте данные из нескольких таблиц Excel в одну таблицу</span><span class="sxs-lookup"><span data-stu-id="cf636-103">Combine data from multiple Excel tables into a single table</span></span>

<span data-ttu-id="cf636-104">Этот пример объединяет данные из нескольких таблиц Excel в одну таблицу, которая включает все строки.</span><span class="sxs-lookup"><span data-stu-id="cf636-104">This sample combines data from multiple Excel tables into a single table that includes all the rows.</span></span> <span data-ttu-id="cf636-105">Предполагается, что все используемые таблицы имеют ту же структуру.</span><span class="sxs-lookup"><span data-stu-id="cf636-105">It assumes that all tables being used have the same structure.</span></span>

<span data-ttu-id="cf636-106">Существует два варианта этого сценария:</span><span class="sxs-lookup"><span data-stu-id="cf636-106">There are two variations of this script:</span></span>

1. <span data-ttu-id="cf636-107">Первый [сценарий объединяет](#sample-code-combine-data-from-multiple-excel-tables-into-a-single-table) все таблицы в файле Excel.</span><span class="sxs-lookup"><span data-stu-id="cf636-107">The [first script](#sample-code-combine-data-from-multiple-excel-tables-into-a-single-table) combines all tables in the Excel file.</span></span>
1. <span data-ttu-id="cf636-108">Второй [сценарий](#sample-code-combine-data-from-multiple-excel-tables-in-select-worksheets-into-a-single-table) выборочно получает таблицы в наборе таблиц.</span><span class="sxs-lookup"><span data-stu-id="cf636-108">The [second script](#sample-code-combine-data-from-multiple-excel-tables-in-select-worksheets-into-a-single-table) selectively gets tables within a set of worksheets.</span></span>

## <a name="sample-code-combine-data-from-multiple-excel-tables-into-a-single-table"></a><span data-ttu-id="cf636-109">Пример кода. Объединяйте данные из нескольких таблиц Excel в одну таблицу</span><span class="sxs-lookup"><span data-stu-id="cf636-109">Sample code: Combine data from multiple Excel tables into a single table</span></span>

<span data-ttu-id="cf636-110">Скачайте пример файла <a href="tables-copy.xlsx">tables-copy.xlsx</a> и используйте его со следующим скриптом, чтобы попробовать его самостоятельно!</span><span class="sxs-lookup"><span data-stu-id="cf636-110">Download the sample file <a href="tables-copy.xlsx">tables-copy.xlsx</a> and use it with the following script to try it out yourself!</span></span>

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

## <a name="sample-code-combine-data-from-multiple-excel-tables-in-select-worksheets-into-a-single-table"></a><span data-ttu-id="cf636-111">Пример кода. Объединяйте данные из нескольких таблиц Excel в отдельных листах в одну таблицу</span><span class="sxs-lookup"><span data-stu-id="cf636-111">Sample code: Combine data from multiple Excel tables in select worksheets into a single table</span></span>

<span data-ttu-id="cf636-112">Скачайте пример файла <a href="tables-select-copy.xlsx">tables-select-copy.xlsx</a> и используйте его со следующим скриптом, чтобы попробовать его самостоятельно!</span><span class="sxs-lookup"><span data-stu-id="cf636-112">Download the sample file <a href="tables-select-copy.xlsx">tables-select-copy.xlsx</a> and use it with the following script to try it out yourself!</span></span>

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

## <a name="training-video-combine-data-from-multiple-excel-tables-into-a-single-table"></a><span data-ttu-id="cf636-113">Обучающее видео. Объединяйте данные из нескольких таблиц Excel в одну таблицу</span><span class="sxs-lookup"><span data-stu-id="cf636-113">Training video: Combine data from multiple Excel tables into a single table</span></span>

<span data-ttu-id="cf636-114">[![Просмотрите пошаговую видеозапись объединения данных из нескольких таблиц Excel в одну таблицу](../../images/merge-tables-vid.jpg)](https://youtu.be/di-8JukK3Lc "Пошаговая видеозапись объединения данных из нескольких таблиц Excel в одну таблицу")</span><span class="sxs-lookup"><span data-stu-id="cf636-114">[![Watch step-by-step video on how to combine data from multiple Excel tables into a single table](../../images/merge-tables-vid.jpg)](https://youtu.be/di-8JukK3Lc "Step-by-step video on how to combine data from multiple Excel tables into a single table")</span></span>
