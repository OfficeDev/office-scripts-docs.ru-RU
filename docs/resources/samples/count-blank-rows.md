---
title: Подсчет пустых строк на листах
description: Узнайте, как использовать Office скрипты, чтобы определить, есть ли пустые строки вместо данных в листах, а затем сообщить о том, сколько строк будет использоваться в потоке Power Automate.
ms.date: 06/29/2021
localization_priority: Normal
ms.openlocfilehash: e5b60779d2ca2de5f4cf4e03ddd6ff7372515ad6
ms.sourcegitcommit: 211c157ca746e266eeb079f5fa1925a1e35ab702
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/07/2021
ms.locfileid: "53313808"
---
# <a name="count-blank-rows-on-sheets"></a><span data-ttu-id="31153-103">Подсчет пустых строк на листах</span><span class="sxs-lookup"><span data-stu-id="31153-103">Count blank rows on sheets</span></span>

<span data-ttu-id="31153-104">Этот проект включает два сценария:</span><span class="sxs-lookup"><span data-stu-id="31153-104">This project includes two scripts:</span></span>

* <span data-ttu-id="31153-105">[Подсчитайте пустые строки на заданном листе:](#sample-code-count-blank-rows-on-a-given-sheet)пересекает используемый диапазон на заданном листе и возвращает количество пустых строк.</span><span class="sxs-lookup"><span data-stu-id="31153-105">[Count blank rows on a given sheet](#sample-code-count-blank-rows-on-a-given-sheet): Traverses the used range on a given worksheet and returns a blank row count.</span></span>
* <span data-ttu-id="31153-106">[Подсчитайте пустые строки](#sample-code-count-blank-rows-on-all-sheets)на всех листах: пересекает используемый диапазон на всех листах и возвращает количество пустых строк. </span><span class="sxs-lookup"><span data-stu-id="31153-106">[Count blank rows on all sheets](#sample-code-count-blank-rows-on-all-sheets): Traverses the used range on _all of the worksheets_ and returns a blank row count.</span></span>

> [!NOTE]
> <span data-ttu-id="31153-107">Для нашего скрипта пустая строка — это строка, в которой нет данных.</span><span class="sxs-lookup"><span data-stu-id="31153-107">For our script, a blank row is any row where there's no data.</span></span> <span data-ttu-id="31153-108">Строка может иметь форматирование.</span><span class="sxs-lookup"><span data-stu-id="31153-108">The row can have formatting.</span></span>

<span data-ttu-id="31153-109">_Этот лист возвращает количество 4 пустых строк_</span><span class="sxs-lookup"><span data-stu-id="31153-109">_This sheet returns count of 4 blank rows_</span></span>

:::image type="content" source="../../images/blank-rows.png" alt-text="Лист с данными с пустыми строками.":::

<span data-ttu-id="31153-111">_Этот лист возвращает количество 0 пустых строк (все строки имеют некоторые данные)_</span><span class="sxs-lookup"><span data-stu-id="31153-111">_This sheet returns count of 0 blank rows (all rows have some data)_</span></span>

:::image type="content" source="../../images/no-blank-rows.png" alt-text="Лист, на котором отображаются данные без пустых строк.":::

## <a name="sample-code-count-blank-rows-on-a-given-sheet"></a><span data-ttu-id="31153-113">Пример кода. Подсчитайте пустые строки на заданном листе</span><span class="sxs-lookup"><span data-stu-id="31153-113">Sample code: Count blank rows on a given sheet</span></span>

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

## <a name="sample-code-count-blank-rows-on-all-sheets"></a><span data-ttu-id="31153-114">Пример кода: количество пустых строк на всех листах</span><span class="sxs-lookup"><span data-stu-id="31153-114">Sample code: Count blank rows on all sheets</span></span>

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
