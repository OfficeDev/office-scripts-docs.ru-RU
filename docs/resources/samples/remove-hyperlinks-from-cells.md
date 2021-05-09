---
title: Удаление гиперссылки из каждой ячейки в Excel таблицы
description: Узнайте, как использовать Office скрипты для удаления гиперссылки из каждой ячейки в Excel таблицы.
ms.date: 05/06/2021
localization_priority: Normal
ms.openlocfilehash: 048d01691377a7086bdba9ceb87ca98839cfa4d1
ms.sourcegitcommit: 763d341857bcb209b2f2c278a82fdb63d0e18f0a
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 05/08/2021
ms.locfileid: "52285803"
---
# <a name="remove-hyperlinks-from-each-cell-in-an-excel-worksheet"></a><span data-ttu-id="af9e0-103">Удаление гиперссылки из каждой ячейки в Excel таблицы</span><span class="sxs-lookup"><span data-stu-id="af9e0-103">Remove hyperlinks from each cell in an Excel worksheet</span></span>

 <span data-ttu-id="af9e0-104">Этот пример очищает все гиперссылки из текущего таблицы.</span><span class="sxs-lookup"><span data-stu-id="af9e0-104">This sample clears all of the hyperlinks from the current worksheet.</span></span> <span data-ttu-id="af9e0-105">Она пересекает таблицу, и если есть гиперссылка, связанная с ячейкой, она очищает гиперссылку, но сохраняет значение ячейки как есть.</span><span class="sxs-lookup"><span data-stu-id="af9e0-105">It traverses the worksheet and if there is any hyperlink associated with the cell, it clears the hyperlink yet retains the cell value as is.</span></span> <span data-ttu-id="af9e0-106">Кроме того, записи времени, необходимого для завершения обхода.</span><span class="sxs-lookup"><span data-stu-id="af9e0-106">Also logs the time it takes to complete traversal.</span></span>

> [!NOTE]
> <span data-ttu-id="af9e0-107">Это работает только в том случае, если < 10k.</span><span class="sxs-lookup"><span data-stu-id="af9e0-107">This only works if the cell count is < 10k.</span></span>

## <a name="sample-code-remove-hyperlinks"></a><span data-ttu-id="af9e0-108">Пример кода. Удаление гиперссылки</span><span class="sxs-lookup"><span data-stu-id="af9e0-108">Sample code: Remove hyperlinks</span></span>

<span data-ttu-id="af9e0-109">Скачайте файл <a href="remove-hyperlinks.xlsx">remove-hyperlinks.xlsx, </a> используемый в этом примере, и попробуйте его самостоятельно!</span><span class="sxs-lookup"><span data-stu-id="af9e0-109">Download the file <a href="remove-hyperlinks.xlsx">remove-hyperlinks.xlsx</a> used in this sample and try it out yourself!</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook, sheetName: string = 'Sheet1') {
  // Get the active worksheet. 
  let sheet = workbook.getWorksheet(sheetName);

  // Get the used range to operate on.
  // For large ranges (over 10000 entries), consider splitting the operation into batches for performance.
  const targetRange = sheet.getUsedRange(true);
  console.log(`Target Range to clear hyperlinks from: ${targetRange.getAddress()}`);

  const rowCount = targetRange.getRowCount();
  const colCount = targetRange.getColumnCount();
  console.log(`Searching for hyperlinks in ${targetRange.getAddress()} which contains ${(rowCount * colCount)} cells`);

  // Go through each individual cell looking for a hyperlink. 
  // This allows us to limit the formatting changes to only the cells with hyperlink formatting.
  let clearedCount = 0;
  for (let i = 0; i < rowCount; i++) {
    for (let j = 0; j < colCount; j++) {
      const cell = targetRange.getCell(i, j);
      const hyperlink = cell.getHyperlink();
      if (hyperlink) {
        cell.clear(ExcelScript.ClearApplyTo.hyperlinks);
        cell.getFormat().getFont().setUnderline(ExcelScript.RangeUnderlineStyle.none);
        cell.getFormat().getFont().setColor('Black');
        clearedCount++;
      }
    }
  }

  console.log(`Done. Cleared hyperlinks from ${clearedCount} cells`);
}
```

## <a name="training-video-remove-hyperlinks-from-each-cell-in-an-excel-worksheet"></a><span data-ttu-id="af9e0-110">Обучающее видео: Удаление гиперссылки из каждой ячейки в Excel таблицы</span><span class="sxs-lookup"><span data-stu-id="af9e0-110">Training video: Remove hyperlinks from each cell in an Excel worksheet</span></span>

<span data-ttu-id="af9e0-111">[Смотреть Sudhi Ramamurthy ходить через этот пример на YouTube](https://youtu.be/v20fdinxpHU).</span><span class="sxs-lookup"><span data-stu-id="af9e0-111">[Watch Sudhi Ramamurthy walk through this sample on YouTube](https://youtu.be/v20fdinxpHU).</span></span>
