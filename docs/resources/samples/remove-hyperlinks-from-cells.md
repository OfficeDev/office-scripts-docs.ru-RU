---
title: Удаление гиперссылки из каждой ячейки в Excel таблицы
description: Узнайте, как использовать Office скрипты для удаления гиперссылки из каждой ячейки в Excel таблицы.
ms.date: 06/29/2021
ms.localizationpriority: medium
ms.openlocfilehash: c318fc9b413f31c1c75c2b4b4bfd31312a7810b5
ms.sourcegitcommit: 7023b9e23499806901a5ecf8ebc460b76887cca6
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/31/2022
ms.locfileid: "64585795"
---
# <a name="remove-hyperlinks-from-each-cell-in-an-excel-worksheet"></a>Удаление гиперссылки из каждой ячейки в Excel таблицы

 Этот пример очищает все гиперссылки из текущего таблицы. Она пересекает таблицу, и если есть гиперссылка, связанная с ячейкой, она очищает гиперссылку, но сохраняет значение ячейки как есть. Кроме того, записи времени, необходимого для завершения обхода.

> [!NOTE]
> Это работает только в том случае, если < 10k.

## <a name="sample-excel-file"></a>Пример Excel файла

Скачайте <a href="remove-hyperlinks.xlsx"> файлremove-hyperlinks.xlsx</a> для готовой к использованию книги. Добавьте следующий скрипт, чтобы попробовать пример самостоятельно!

## <a name="sample-code-remove-hyperlinks"></a>Пример кода. Удаление гиперссылки

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

## <a name="training-video-remove-hyperlinks-from-each-cell-in-an-excel-worksheet"></a>Обучающее видео: Удаление гиперссылки из каждой ячейки в Excel таблицы

[Посмотрите, как суди Рамамурти (Sudhi Ramamurthy) пройдите этот пример на YouTube](https://youtu.be/v20fdinxpHU).
