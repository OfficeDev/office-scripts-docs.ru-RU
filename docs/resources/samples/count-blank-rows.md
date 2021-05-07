---
title: Подсчет пустых строк на листах
description: Узнайте, как использовать Office скрипты, чтобы определить, есть ли пустые строки вместо данных в листах, а затем сообщить о том, сколько строк будет использоваться в потоке Power Automate.
ms.date: 03/31/2021
localization_priority: Normal
ms.openlocfilehash: db84f2446c168f867c325a05129fe982c9645731
ms.sourcegitcommit: f7a7aebfb687f2a35dbed07ed62ff352a114525a
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 05/06/2021
ms.locfileid: "52232587"
---
# <a name="count-blank-rows-on-sheets"></a>Подсчет пустых строк на листах

Этот проект включает два сценария:

* [Подсчитайте пустые строки на заданном листе:](#sample-code-count-blank-rows-on-a-given-sheet)пересекает используемый диапазон на заданном листе и возвращает количество пустых строк.
* [Подсчитайте пустые строки](#sample-code-count-blank-rows-on-all-sheets)на всех листах: пересекает используемый диапазон на всех листах и возвращает количество пустых строк. 

> [!NOTE]
> Для нашего скрипта пустая строка — это строка, в которой нет данных. Строка может иметь форматирование.

_Этот лист возвращает количество 4 пустых строк_

:::image type="content" source="../../images/blank-rows.png" alt-text="Лист с данными с пустыми строками":::

_Этот лист возвращает количество 0 пустых строк (все строки имеют некоторые данные)_

:::image type="content" source="../../images/no-blank-rows.png" alt-text="Лист с данными без пустых строк":::

## <a name="sample-code-count-blank-rows-on-a-given-sheet"></a>Пример кода. Подсчитайте пустые строки на заданном листе

```TypeScript
function main(workbook: ExcelScript.Workbook): number
{
  const sheet = workbook.getWorksheet('Sheet1'); 
  // Getting the active worksheet is not suitable for a script used by Power Automate.
  // const sheet = workbook.getActiveWorksheet();
  
  const range = sheet.getUsedRange(true); // Get value only.
  if (!range) {
    console.log(`No data on this sheet. `);
    return;
  }
  console.log(`Used range for the worksheet: ${range.getAddress()}`);
  const values = range.getValues();
  let emptyRows = 0;
  for (let row of values) {
    let len = 0; 
    for (let cell of row) {
      len = len + cell.toString().length;
    }
    if (len === 0) { 
      emptyRows++;
    }
  }
  console.log(`Total empty row: ` + emptyRows);
  return emptyRows;
}
```

## <a name="sample-code-count-blank-rows-on-all-sheets"></a>Пример кода: количество пустых строк на всех листах

```TypeScript
function main(workbook: ExcelScript.Workbook): number
{
  const sheets = workbook.getWorksheets();
  let emptyRows = 0;
  for (let sheet of sheets) { 
    const range = sheet.getUsedRange(true); // Get value only.
    if (!range) {
      console.log(`No data on this sheet. `);
      continue;
    }
    console.log(`Used range for the worksheet ${sheet.getName()}: ${range.getAddress()}`);
    const values = range.getValues();

    for (let row of values) {
      let len = 0;
      for (let cell of row) {
        len = len + cell.toString().length;
      }
      if (len === 0) {
        emptyRows++;
      }
    }
  }
  console.log(`Total empty row: ` + emptyRows);
  return emptyRows;
}
```

## <a name="use-with-power-automate"></a>Использование с Power Automate

:::image type="content" source="../../images/use-in-power-automate.png" alt-text="Поток Power Automate, показывающий, как настроить запуск Office скрипта":::
