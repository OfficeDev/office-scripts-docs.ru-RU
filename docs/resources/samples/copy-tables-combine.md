---
title: Объединяйте данные из нескольких Excel таблиц в одну таблицу
description: Узнайте, как использовать Office скрипты для объединения данных из нескольких Excel таблиц в одну таблицу.
ms.date: 06/29/2021
ms.localizationpriority: medium
ms.openlocfilehash: 490727f39a497dd1d2e31f2fac938b6d518012a5
ms.sourcegitcommit: d3ed4bdeeba805d97c930394e172e8306a0cf484
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/15/2021
ms.locfileid: "59330774"
---
# <a name="combine-data-from-multiple-excel-tables-into-a-single-table"></a>Объединяйте данные из нескольких Excel таблиц в одну таблицу

Этот пример объединяет данные из нескольких Excel таблиц в одну таблицу, которая включает все строки. Предполагается, что все используемые таблицы имеют ту же структуру.

Существует два варианта этого сценария:

1. Первый [скрипт объединяет](#sample-code-combine-data-from-multiple-excel-tables-into-a-single-table) все таблицы в Excel файле.
1. Второй [сценарий](#sample-code-combine-data-from-multiple-excel-tables-in-select-worksheets-into-a-single-table) выборочно получает таблицы в наборе таблиц.

## <a name="sample-excel-file"></a>Пример Excel файла

Скачайте <a href="tables-copy.xlsx">tables-copy.xlsx</a> для готовой к использованию книги. Добавьте следующие скрипты, чтобы попробовать пример самостоятельно!

## <a name="sample-code-combine-data-from-multiple-excel-tables-into-a-single-table"></a>Пример кода. Объединяйте данные из нескольких Excel таблиц в одну таблицу

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Delete the "Combined" worksheet, if it's present.
  workbook.getWorksheet('Combined')?.delete();

  // Create a new worksheet named "Combined" for the combined table.
  const newSheet = workbook.addWorksheet('Combined');
  
  // Get the header values for the first table in the workbook.
  // This also saves the table list before we add the new, combined table.
  const tables = workbook.getTables();    
  const headerValues = tables[0].getHeaderRowRange().getTexts();
  console.log(headerValues);

  // Copy the headers on a new worksheet to an equal-sized range.
  const targetRange = newSheet.getRange('A1').getResizedRange(headerValues.length-1, headerValues[0].length-1);
  targetRange.setValues(headerValues);

  // Add the data from each table in the workbook to the new table.
  const combinedTable = newSheet.addTable(targetRange.getAddress(), true);
  for (let table of tables) {      
    let dataValues = table.getRangeBetweenHeaderAndTotal().getTexts();
    let rowCount = table.getRowCount();

    // If the table is not empty, add its rows to the combined table.
    if (rowCount > 0) {
      combinedTable.addRows(-1, dataValues);
    }
  }
}
```

## <a name="sample-code-combine-data-from-multiple-excel-tables-in-select-worksheets-into-a-single-table"></a>Пример кода. Объединяйте данные из нескольких Excel таблиц в отдельных таблицах в одну таблицу

Скачайте пример файла <a href="tables-select-copy.xlsx">tables-select-copy.xlsx</a> и используйте его со следующим скриптом, чтобы попробовать его самостоятельно!

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Set the worksheet names to get tables from.
  const sheetNames = ['Sheet1', 'Sheet2', 'Sheet3'];
    
  // Delete the "Combined" worksheet, if it's present.
  workbook.getWorksheet('Combined')?.delete();

  // Create a new worksheet named "Combined" for the combined table.
  const newSheet = workbook.addWorksheet('Combined');

  // Create a new table with the same headers as the other tables.
  const headerValues = workbook.getWorksheet(sheetNames[0]).getTables()[0].getHeaderRowRange().getTexts();
  const targetRange = newSheet.getRange('A1').getResizedRange(headerValues.length-1, headerValues[0].length-1);
  targetRange.setValues(headerValues);
  const combinedTable = newSheet.addTable(targetRange.getAddress(), true);

  // Go through each listed worksheet and get their tables.
  sheetNames.forEach((sheet) => {
    const tables = workbook.getWorksheet(sheet).getTables();     
    for (let table of tables) {
      // Get the rows from the tables.
      let dataValues = table.getRangeBetweenHeaderAndTotal().getTexts();
      let rowCount = table.getRowCount();

      // If there's data in the table, add it to the combined table.
      if (rowCount > 0) {
          combinedTable.addRows(-1, dataValues);
      }
    }
  });
}
```

## <a name="training-video-combine-data-from-multiple-excel-tables-into-a-single-table"></a>Обучающее видео. Объединяйте данные из нескольких Excel таблиц в одну таблицу

[Смотреть Sudhi Ramamurthy ходить через этот пример на YouTube](https://youtu.be/di-8JukK3Lc).
