---
title: Перекрестная ссылка и формат Excel файла
description: Узнайте, как использовать Office и Power Automate для перекрестной ссылки и формата Excel файла.
ms.date: 04/28/2021
localization_priority: Normal
ms.openlocfilehash: 858fe561c1a82f471bc3c0f43d81e457fb02b627
ms.sourcegitcommit: f7a7aebfb687f2a35dbed07ed62ff352a114525a
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 05/06/2021
ms.locfileid: "52232384"
---
# <a name="cross-reference-and-format-an-excel-file"></a>Перекрестная ссылка и формат Excel файла

Это решение показывает, как Excel двух файлов можно перекрестно ссылаться и форматирование с помощью Office и Power Automate.

В проекте реализуется следующее:

1. Извлекает данные событий из <a href="events.xlsx">events.xlsx</a> с помощью одного действия скрипта Run.
1. Передает эти данные во второй Excel, содержащий данные транзакций событий, и использует эти данные для базовой проверки данных и форматирования отсутствующих или неправильных данных с помощью Office Scripts.
1. По электронной почте результат передается рецензенту.

Дополнительные сведения см. в перекрестной ссылке и [форматирования двух Excel с помощью Office Scripts.](https://powerusers.microsoft.com/t5/Power-Automate-Cookbook/Cross-Reference-and-formatting-two-Excel-files-using-Office/td-p/728535)

## <a name="sample-excel-files"></a>Пример Excel файлов

Скачайте следующие файлы, используемые в этом решении, чтобы попробовать его самостоятельно!

1. <a href="events.xlsx">events.xlsx</a>
1. <a href="event-transactions.xlsx">event-transactions.xlsx</a>

## <a name="sample-code-get-event-data"></a>Пример кода: получить данные событий

```TypeScript
function main(workbook: ExcelScript.Workbook): EventData[] {
    let table = workbook.getWorksheet('Keys').getTables()[0];
    let range = table.getRangeBetweenHeaderAndTotal();
    let rows = range.getValues();
    let records: EventData[] = [];
    for (let row of rows) {
        let [event, date, location, capacity] = row;
        records.push({
            event: event as string,
            date: date as number, 
            location: location as string,
            capacity: capacity as number
        })
    }
    console.log(JSON.stringify(records))
    return records;
}

interface EventData {
    event: string
    date: number
    location: string
    capacity: number
}
```

## <a name="sample-code-validate-event-transactions"></a>Пример кода. Проверка транзакций событий

```TypeScript
function main(workbook: ExcelScript.Workbook, keys: string): string {
    let table = workbook.getWorksheet('Transactions').getTables()[0];
    let range = table.getRangeBetweenHeaderAndTotal();
    range.clear(ExcelScript.ClearApplyTo.formats);
  
    let overallMatch = true;
  
    table.getColumnByName('Date').getRangeBetweenHeaderAndTotal().setNumberFormatLocal("yyyy-mm-dd;@");
    table.getColumnByName('Capacity').getRangeBetweenHeaderAndTotal().getFormat()
      .setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);
    let rows = range.getValues();
    let keysObject = JSON.parse(keys) as EventData[];
    for (let i=0; i < rows.length; i++){
      let row = rows[i];
      let [event, date, location, capacity] = row;
      let match = false;
      for (let keyObject of keysObject){
        if (keyObject.event === event) {
          match = true;
          if (keyObject.date !== date) {
            overallMatch = false;
            range.getCell(i, 1).getFormat()
              .getFill()
              .setColor("FFFF00");
          }
          if (keyObject.location !== location) {
            overallMatch = false;
            range.getCell(i, 2).getFormat()
              .getFill()
              .setColor("FFFF00");
          }
          if (keyObject.capacity !== capacity) {
            overallMatch = false;
            range.getCell(i, 3).getFormat()
              .getFill()
              .setColor("FFFF00");
          }   
          break;             
        }
      }
      if (!match) {
        overallMatch = false;
        range.getCell(i, 0).getFormat()
          .getFill()
          .setColor("FFFF00");      
      }
  
    }
    let returnString = "All the data is in the right order.";
    if (overallMatch === false) {
      returnString = "Mismatch found. Data requires your review.";
    }
    console.log("Returning: " + returnString);
    return returnString;
}

interface EventData {
event: string
date: number
location: string
capacity: number
}
```

## <a name="training-video-cross-reference-and-format-an-excel-file"></a>Обучающее видео: перекрестная ссылка и формат Excel файла

[Смотреть Sudhi Ramamurthy ходить через этот пример на YouTube](https://youtu.be/dVwqBf483qo").
