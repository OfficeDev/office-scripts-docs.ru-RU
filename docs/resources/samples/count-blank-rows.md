---
title: Подсчет пустых строк на листах
description: Узнайте, как использовать Office скрипты, чтобы определить, есть ли пустые строки вместо данных в листах, а затем сообщить о том, сколько строк будет использоваться в потоке Power Automate.
ms.date: 06/29/2021
ms.localizationpriority: medium
ms.openlocfilehash: 1ae513928b885994dc7f6d1b8ad66d694b61e7b7
ms.sourcegitcommit: 7023b9e23499806901a5ecf8ebc460b76887cca6
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/31/2022
ms.locfileid: "64585564"
---
# <a name="count-blank-rows-on-sheets"></a>Подсчет пустых строк на листах

Этот проект включает два сценария:

* [Подсчитайте пустые строки на заданном](#sample-code-count-blank-rows-on-a-given-sheet) листе: пересекает используемый диапазон на заданном листе и возвращает количество пустых строк.
* [Подсчитайте пустые строки на всех листах](#sample-code-count-blank-rows-on-all-sheets): пересекает используемый  диапазон на всех листах и возвращает количество пустых строк.

> [!NOTE]
> Для нашего скрипта пустая строка — это строка, в которой нет данных. Строка может иметь форматирование.

_Этот лист возвращает количество 4 пустых строк_

:::image type="content" source="../../images/blank-rows.png" alt-text="Лист с данными с пустыми строками.":::

_Этот лист возвращает количество 0 пустых строк (все строки имеют некоторые данные)_

:::image type="content" source="../../images/no-blank-rows.png" alt-text="Лист, на котором отображаются данные без пустых строк.":::

## <a name="sample-code-count-blank-rows-on-a-given-sheet"></a>Пример кода. Подсчитайте пустые строки на заданном листе

```TypeScript
function main(workbook: ExcelScript.Workbook): number
{
  // Get the worksheet named "Sheet1".
  const sheet = workbook.getWorksheet('Sheet1'); 
  
  // Get the entire data range.
  const range = sheet.getUsedRange(true);

  // If the used range is empty, end the script.
  if (!range) {
    console.log(`No data on this sheet.`);
    return;
  }
  
  // Log the address of the used range.
  console.log(`Used range for the worksheet: ${range.getAddress()}`);
    
  // Look through the values in the range for blank rows.
  const values = range.getValues();
  let emptyRows = 0;
  for (let row of values) {
    let emptyRow = true;
    
    // Look at every cell in the row for one with a value.
    for (let cell of row) {
      if (cell.toString().length > 0) {
        emptyRow = false
      }
    }

    // If no cell had a value, the row is empty.
    if (emptyRow) {
      emptyRows++;
    }
  }

  // Log the number of empty rows.
  console.log(`Total empty rows: ${emptyRows}`);

  // Return the number of empty rows for use in a Power Automate flow.
  return emptyRows;
}
```

## <a name="sample-code-count-blank-rows-on-all-sheets"></a>Пример кода: количество пустых строк на всех листах

```TypeScript
function main(workbook: ExcelScript.Workbook): number
{
  // Loop through every worksheet in the workbook.
  const sheets = workbook.getWorksheets();
  let emptyRows = 0;
  for (let sheet of sheets) {     
    // Get the entire data range.
    const range = sheet.getUsedRange(true);
  
    // If the used range is empty, skip to the next worksheet.
    if (!range) {
      console.log(`No data on this sheet.`);
      continue;
    }
    
    // Log the address of the used range.
    console.log(`Used range for the worksheet: ${range.getAddress()}`);
      
    // Look through the values in the range for blank rows.
    const values = range.getValues();
    for (let row of values) {
      let emptyRow = true;
      
      // Look at every cell in the row for one with a value.
      for (let cell of row) {
        if (cell.toString().length > 0) {
          emptyRow = false
        }
      }
  
      // If no cell had a value, the row is empty.
      if (emptyRow) {
        emptyRows++;
      }
    }
  }

  // Log the number of empty rows.
  console.log(`Total empty rows: ${emptyRows}`);

  // Return the number of empty rows for use in a Power Automate flow.
  return emptyRows;
}
```
