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
# <a name="cross-reference-and-format-an-excel-file"></a><span data-ttu-id="cbe07-103">Перекрестная ссылка и формат Excel файла</span><span class="sxs-lookup"><span data-stu-id="cbe07-103">Cross-reference and format an Excel file</span></span>

<span data-ttu-id="cbe07-104">Это решение показывает, как Excel двух файлов можно перекрестно ссылаться и форматирование с помощью Office и Power Automate.</span><span class="sxs-lookup"><span data-stu-id="cbe07-104">This solution shows how two Excel files can be cross-referenced and formatted using Office Scripts and Power Automate.</span></span>

<span data-ttu-id="cbe07-105">В проекте реализуется следующее:</span><span class="sxs-lookup"><span data-stu-id="cbe07-105">The project achieves the following:</span></span>

1. <span data-ttu-id="cbe07-106">Извлекает данные событий из <a href="events.xlsx">events.xlsx</a> с помощью одного действия скрипта Run.</span><span class="sxs-lookup"><span data-stu-id="cbe07-106">Extracts event data from <a href="events.xlsx">events.xlsx</a> using one Run script action.</span></span>
1. <span data-ttu-id="cbe07-107">Передает эти данные во второй Excel, содержащий данные транзакций событий, и использует эти данные для базовой проверки данных и форматирования отсутствующих или неправильных данных с помощью Office Scripts.</span><span class="sxs-lookup"><span data-stu-id="cbe07-107">Passes that data to the second Excel file containing event transaction data and uses that data to do basic validation of data and formatting of missing or incorrect data using Office Scripts.</span></span>
1. <span data-ttu-id="cbe07-108">По электронной почте результат передается рецензенту.</span><span class="sxs-lookup"><span data-stu-id="cbe07-108">Emails the result to a reviewer.</span></span>

<span data-ttu-id="cbe07-109">Дополнительные сведения см. в перекрестной ссылке и [форматирования двух Excel с помощью Office Scripts.](https://powerusers.microsoft.com/t5/Power-Automate-Cookbook/Cross-Reference-and-formatting-two-Excel-files-using-Office/td-p/728535)</span><span class="sxs-lookup"><span data-stu-id="cbe07-109">For further details, see [Cross Reference and formatting two Excel files using Office Scripts](https://powerusers.microsoft.com/t5/Power-Automate-Cookbook/Cross-Reference-and-formatting-two-Excel-files-using-Office/td-p/728535).</span></span>

## <a name="sample-excel-files"></a><span data-ttu-id="cbe07-110">Пример Excel файлов</span><span class="sxs-lookup"><span data-stu-id="cbe07-110">Sample Excel files</span></span>

<span data-ttu-id="cbe07-111">Скачайте следующие файлы, используемые в этом решении, чтобы попробовать его самостоятельно!</span><span class="sxs-lookup"><span data-stu-id="cbe07-111">Download the following files used in this solution to try it out yourself!</span></span>

1. <span data-ttu-id="cbe07-112"><a href="events.xlsx">events.xlsx</a></span><span class="sxs-lookup"><span data-stu-id="cbe07-112"><a href="events.xlsx">events.xlsx</a></span></span>
1. <span data-ttu-id="cbe07-113"><a href="event-transactions.xlsx">event-transactions.xlsx</a></span><span class="sxs-lookup"><span data-stu-id="cbe07-113"><a href="event-transactions.xlsx">event-transactions.xlsx</a></span></span>

## <a name="sample-code-get-event-data"></a><span data-ttu-id="cbe07-114">Пример кода: получить данные событий</span><span class="sxs-lookup"><span data-stu-id="cbe07-114">Sample code: Get event data</span></span>

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

## <a name="sample-code-validate-event-transactions"></a><span data-ttu-id="cbe07-115">Пример кода. Проверка транзакций событий</span><span class="sxs-lookup"><span data-stu-id="cbe07-115">Sample code: Validate event transactions</span></span>

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

## <a name="training-video-cross-reference-and-format-an-excel-file"></a><span data-ttu-id="cbe07-116">Обучающее видео: перекрестная ссылка и формат Excel файла</span><span class="sxs-lookup"><span data-stu-id="cbe07-116">Training video: Cross-reference and format an Excel file</span></span>

<span data-ttu-id="cbe07-117">[Смотреть Sudhi Ramamurthy ходить через этот пример на YouTube](https://youtu.be/dVwqBf483qo").</span><span class="sxs-lookup"><span data-stu-id="cbe07-117">[Watch Sudhi Ramamurthy walk through this sample on YouTube](https://youtu.be/dVwqBf483qo").</span></span>
