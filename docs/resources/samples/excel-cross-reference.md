---
title: Перекрестная ссылка и формат Excel файла
description: Узнайте, как Office скрипты и Power Automate для перекрестной ссылки и форматировать Excel файл.
ms.date: 05/06/2021
localization_priority: Normal
ROBOTS: NOINDEX
ms.openlocfilehash: f07395eb4e6c77b7aee3776e3252d135bc690a6f
ms.sourcegitcommit: 4687693f02fc90a57ba30c461f35046e02e6f5fb
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 05/19/2021
ms.locfileid: "52545768"
---
# <a name="cross-reference-and-format-an-excel-file"></a><span data-ttu-id="3875f-103">Перекрестная ссылка и формат Excel файла</span><span class="sxs-lookup"><span data-stu-id="3875f-103">Cross-reference and format an Excel file</span></span>

<span data-ttu-id="3875f-104">Это решение показывает, как Excel файлы могут быть перекрестными ссылками и отформатированы с помощью Office скриптов и Power Automate.</span><span class="sxs-lookup"><span data-stu-id="3875f-104">This solution shows how two Excel files can be cross-referenced and formatted using Office Scripts and Power Automate.</span></span>

<span data-ttu-id="3875f-105">Проект достигает следующих результатов:</span><span class="sxs-lookup"><span data-stu-id="3875f-105">The project achieves the following:</span></span>

1. <span data-ttu-id="3875f-106">Извлекает данные событий из <a href="events.xlsx">events.xlsxс </a> помощью одного действия сценария Run.</span><span class="sxs-lookup"><span data-stu-id="3875f-106">Extracts event data from <a href="events.xlsx">events.xlsx</a> using one Run script action.</span></span>
1. <span data-ttu-id="3875f-107">Передает эти данные во второй файл Excel, содержащий данные о транзакциях событий, и использует эти данные для базовой проверки данных и форматирования недостающих или неверных данных с помощью Office скриптов.</span><span class="sxs-lookup"><span data-stu-id="3875f-107">Passes that data to the second Excel file containing event transaction data and uses that data to do basic validation of data and formatting of missing or incorrect data using Office Scripts.</span></span>
1. <span data-ttu-id="3875f-108">Электронная почта результат рецензента.</span><span class="sxs-lookup"><span data-stu-id="3875f-108">Emails the result to a reviewer.</span></span>

<span data-ttu-id="3875f-109">Для получения более подробной [информации с Office Excel м.](https://powerusers.microsoft.com/t5/Power-Automate-Cookbook/Cross-Reference-and-formatting-two-Excel-files-using-Office/td-p/728535)</span><span class="sxs-lookup"><span data-stu-id="3875f-109">For further details, see [Cross Reference and formatting two Excel files using Office Scripts](https://powerusers.microsoft.com/t5/Power-Automate-Cookbook/Cross-Reference-and-formatting-two-Excel-files-using-Office/td-p/728535).</span></span>

## <a name="sample-excel-files"></a><span data-ttu-id="3875f-110">Пример Excel файлов</span><span class="sxs-lookup"><span data-stu-id="3875f-110">Sample Excel files</span></span>

<span data-ttu-id="3875f-111">Скачать следующие файлы, используемые в этом решении, чтобы попробовать его самостоятельно!</span><span class="sxs-lookup"><span data-stu-id="3875f-111">Download the following files used in this solution to try it out yourself!</span></span>

1. <span data-ttu-id="3875f-112"><a href="events.xlsx">events.xlsx</a></span><span class="sxs-lookup"><span data-stu-id="3875f-112"><a href="events.xlsx">events.xlsx</a></span></span>
1. <span data-ttu-id="3875f-113"><a href="event-transactions.xlsx">event-transactions.xlsx</a></span><span class="sxs-lookup"><span data-stu-id="3875f-113"><a href="event-transactions.xlsx">event-transactions.xlsx</a></span></span>

## <a name="sample-code-get-event-data"></a><span data-ttu-id="3875f-114">Пример кода: Получить данные о событиях</span><span class="sxs-lookup"><span data-stu-id="3875f-114">Sample code: Get event data</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook): EventData[] {
  // Get the first table in the "Keys" worksheet.
  let table = workbook.getWorksheet('Keys').getTables()[0];
  
  // Get the rows in the event table.
  let range = table.getRangeBetweenHeaderAndTotal();
  let rows = range.getValues();

  // Save each row as an EventData object. This lets them be passed through Power Automate.
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

  // Log the event data to the console and return it for a flow.
  console.log(JSON.stringify(records));
  return records;
}

// An interface representing a row of event data.
interface EventData {
  event: string
  date: number
  location: string
  capacity: number
}
```

## <a name="sample-code-validate-event-transactions"></a><span data-ttu-id="3875f-115">Пример кода: Проверка транзакций событий</span><span class="sxs-lookup"><span data-stu-id="3875f-115">Sample code: Validate event transactions</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook, keys: string): string {
  // Get the first table in the "Transactions" worksheet.
  let table = workbook.getWorksheet('Transactions').getTables()[0];

  // Clear the existing formatting in the table.
  let range = table.getRangeBetweenHeaderAndTotal();
  range.clear(ExcelScript.ClearApplyTo.formats);
    
 // Apply some basic formatting for readability.
  table.getColumnByName('Date').getRangeBetweenHeaderAndTotal().setNumberFormatLocal("yyyy-mm-dd;@");
  table.getColumnByName('Capacity').getRangeBetweenHeaderAndTotal().getFormat()
    .setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);

  // Compare the data in the table to the keys passed into the script.
  let keysObject = JSON.parse(keys) as EventData[];
  let overallMatch = true;

  // Iterate over every row.
  let rows = range.getValues();
  for (let i = 0; i < rows.length; i++) {
    let row = rows[i];
    let [event, date, location, capacity] = row;
    let match = false;

    // Look at each key provided for a matching Event ID.
    for (let keyObject of keysObject) {
      if (keyObject.event === event) {
        match = true;

        // If there's a match on the event ID, look for things that don't match and highlight them.
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

    // If no matching Event ID is found, highlight the Event ID's cell.
    if (!match) {
      overallMatch = false;
      range.getCell(i, 0).getFormat()
        .getFill()
        .setColor("FFFF00");      
    }  
  }

  // Choose a message to send to the user.
  let returnString = "All the data is in the right order.";
  if (overallMatch === false) {
    returnString = "Mismatch found. Data requires your review.";
  }
  console.log("Returning: " + returnString);
  return returnString;
}

// An interface representing a row of event data.
interface EventData {
  event: string
  date: number
  location: string
  capacity: number
}
```

## <a name="training-video-cross-reference-and-format-an-excel-file"></a><span data-ttu-id="3875f-116">Учебное видео: Кросс-справка и формат Excel файл</span><span class="sxs-lookup"><span data-stu-id="3875f-116">Training video: Cross-reference and format an Excel file</span></span>

<span data-ttu-id="3875f-117">[Смотреть Судхи Рамамурти ходить через этот образец на YouTube](https://youtu.be/dVwqBf483qo").</span><span class="sxs-lookup"><span data-stu-id="3875f-117">[Watch Sudhi Ramamurthy walk through this sample on YouTube](https://youtu.be/dVwqBf483qo").</span></span>
