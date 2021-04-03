---
title: Перекрестная ссылка и формат файла Excel
description: Узнайте, как использовать office Scripts и Power Automate для перекрестной ссылки и формата файла Excel.
ms.date: 03/18/2021
localization_priority: Normal
ms.openlocfilehash: 287de604733b7e6a126d0c81cb4e23351e558c61
ms.sourcegitcommit: 5d24e77df70aa2c1c982275d53213c2a9323ff86
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 04/02/2021
ms.locfileid: "51571512"
---
# <a name="cross-reference-and-format-an-excel-file"></a><span data-ttu-id="c8d55-103">Перекрестная ссылка и формат файла Excel</span><span class="sxs-lookup"><span data-stu-id="c8d55-103">Cross-reference and format an Excel file</span></span>

<span data-ttu-id="c8d55-104">В этом решении показано, как два файла Excel можно перекрестно ссылаться и форматировать с помощью office Scripts и Power Automate.</span><span class="sxs-lookup"><span data-stu-id="c8d55-104">This solution shows how two Excel files can be cross-referenced and formatted using Office Scripts and Power Automate.</span></span>

<span data-ttu-id="c8d55-105">В проекте реализуется следующее:</span><span class="sxs-lookup"><span data-stu-id="c8d55-105">The project achieves the following:</span></span>

1. <span data-ttu-id="c8d55-106">Извлекает данные событий из <a href="events.xlsx">events.xlsx</a> с помощью одного действия скрипта Run.</span><span class="sxs-lookup"><span data-stu-id="c8d55-106">Extracts event data from <a href="events.xlsx">events.xlsx</a> using one Run script action.</span></span>
1. <span data-ttu-id="c8d55-107">Передает эти данные во второй файл Excel, содержащий данные транзакций событий, и использует эти данные для базовой проверки данных и форматирования отсутствующих или неправильных данных с помощью скриптов Office.</span><span class="sxs-lookup"><span data-stu-id="c8d55-107">Passes that data to the second Excel file containing event transaction data and uses that data to do basic validation of data and formatting of missing or incorrect data using Office Scripts.</span></span>
1. <span data-ttu-id="c8d55-108">По электронной почте результат передается рецензенту.</span><span class="sxs-lookup"><span data-stu-id="c8d55-108">Emails the result to a reviewer.</span></span>

<span data-ttu-id="c8d55-109">Дополнительные сведения см. в [перекрестной ссылке и форматирования двух файлов Excel с помощью скриптов Office.](https://powerusers.microsoft.com/t5/Power-Automate-Cookbook/Cross-Reference-and-formatting-two-Excel-files-using-Office/td-p/728535)</span><span class="sxs-lookup"><span data-stu-id="c8d55-109">For further details, see [Cross Reference and formatting two Excel files using Office Scripts](https://powerusers.microsoft.com/t5/Power-Automate-Cookbook/Cross-Reference-and-formatting-two-Excel-files-using-Office/td-p/728535).</span></span>

## <a name="sample-excel-files"></a><span data-ttu-id="c8d55-110">Примеры файлов Excel</span><span class="sxs-lookup"><span data-stu-id="c8d55-110">Sample Excel files</span></span>

<span data-ttu-id="c8d55-111">Скачайте следующие файлы, используемые в этом решении, чтобы попробовать его самостоятельно!</span><span class="sxs-lookup"><span data-stu-id="c8d55-111">Download the following files used in this solution to try it out yourself!</span></span>

1. <span data-ttu-id="c8d55-112"><a href="events.xlsx">events.xlsx</a></span><span class="sxs-lookup"><span data-stu-id="c8d55-112"><a href="events.xlsx">events.xlsx</a></span></span>
1. <span data-ttu-id="c8d55-113"><a href="event-transactions.xlsx">event-transactions.xlsx</a></span><span class="sxs-lookup"><span data-stu-id="c8d55-113"><a href="event-transactions.xlsx">event-transactions.xlsx</a></span></span>

## <a name="sample-code-get-event-data"></a><span data-ttu-id="c8d55-114">Пример кода: получить данные событий</span><span class="sxs-lookup"><span data-stu-id="c8d55-114">Sample code: Get event data</span></span>

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

## <a name="sample-code-validate-event-transactions"></a><span data-ttu-id="c8d55-115">Пример кода. Проверка транзакций событий</span><span class="sxs-lookup"><span data-stu-id="c8d55-115">Sample code: Validate event transactions</span></span>

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

## <a name="training-video-cross-reference-and-format-an-excel-file"></a><span data-ttu-id="c8d55-116">Обучающее видео: перекрестная ссылка и формат файла Excel</span><span class="sxs-lookup"><span data-stu-id="c8d55-116">Training video: Cross-reference and format an Excel file</span></span>

<span data-ttu-id="c8d55-117">[![Просмотр пошагового видео о перекрестной ссылке и формате файла Excel](../../images/cross-ref-tables-vid.jpg)](https://youtu.be/dVwqBf483qo "Пошаговая видеозапись о перекрестной ссылке и формате файла Excel")</span><span class="sxs-lookup"><span data-stu-id="c8d55-117">[![Watch step-by-step video on how to cross-reference and format an Excel file](../../images/cross-ref-tables-vid.jpg)](https://youtu.be/dVwqBf483qo "Step-by-step video on how to cross-reference and format an Excel file")</span></span>
