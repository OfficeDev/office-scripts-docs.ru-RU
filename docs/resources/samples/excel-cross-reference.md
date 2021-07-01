---
title: Перекрестные Excel файлы с Power Automate
description: Узнайте, как использовать Office и Power Automate для перекрестной ссылки и формата Excel файла.
ms.date: 06/25/2021
localization_priority: Normal
ms.openlocfilehash: 89c4a5fa5dcff21681fa20cd4118447d39d9b6da
ms.sourcegitcommit: a063b3faf6c1b7c294bd6a73e46845b352f2a22d
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 06/29/2021
ms.locfileid: "53202877"
---
# <a name="cross-reference-excel-files-with-power-automate"></a><span data-ttu-id="32789-103">Перекрестные Excel файлы с Power Automate</span><span class="sxs-lookup"><span data-stu-id="32789-103">Cross-reference Excel files with Power Automate</span></span>

<span data-ttu-id="32789-104">В этом решении показано, как сравнить данные между двумя Excel файлами, чтобы найти несоответствия.</span><span class="sxs-lookup"><span data-stu-id="32789-104">This solution shows how to compare data across two Excel files to find discrepancies.</span></span> <span data-ttu-id="32789-105">Он использует Office скрипты для анализа данных и Power Automate для связи между книгами.</span><span class="sxs-lookup"><span data-stu-id="32789-105">It uses Office Scripts to analyze data and Power Automate to communicate between the workbooks.</span></span>

## <a name="example-scenario"></a><span data-ttu-id="32789-106">Пример сценария</span><span class="sxs-lookup"><span data-stu-id="32789-106">Example scenario</span></span>

<span data-ttu-id="32789-107">Вы координатор событий, который составляет расписание докладчиков для предстоящих конференций.</span><span class="sxs-lookup"><span data-stu-id="32789-107">You're an event coordinator who is scheduling speakers for upcoming conferences.</span></span> <span data-ttu-id="32789-108">Данные событий будут храниться в одной таблице, а регистры динамиков - в другой.</span><span class="sxs-lookup"><span data-stu-id="32789-108">You keep the event data in one spreadsheet and the speaker registrations in another.</span></span> <span data-ttu-id="32789-109">Чтобы обеспечить синхронизацию двух книг, для выделения потенциальных проблем используется поток с Office скриптами.</span><span class="sxs-lookup"><span data-stu-id="32789-109">To ensure the two workbooks are kept in sync, you use a flow with Office Scripts to highlight any potential problems.</span></span>

## <a name="sample-excel-files"></a><span data-ttu-id="32789-110">Пример Excel файлов</span><span class="sxs-lookup"><span data-stu-id="32789-110">Sample Excel files</span></span>

<span data-ttu-id="32789-111">Скачайте следующие файлы, используемые в этом решении, чтобы попробовать его самостоятельно!</span><span class="sxs-lookup"><span data-stu-id="32789-111">Download the following files used in this solution to try it out yourself!</span></span>

1. <span data-ttu-id="32789-112"><a href="event-data.xlsx">event-data.xlsx</a></span><span class="sxs-lookup"><span data-stu-id="32789-112"><a href="event-data.xlsx">event-data.xlsx</a></span></span>
1. <span data-ttu-id="32789-113"><a href="speaker-registrations.xlsx">speaker-registrations.xlsx</a></span><span class="sxs-lookup"><span data-stu-id="32789-113"><a href="speaker-registrations.xlsx">speaker-registrations.xlsx</a></span></span>

## <a name="sample-code-get-event-data"></a><span data-ttu-id="32789-114">Пример кода: получить данные событий</span><span class="sxs-lookup"><span data-stu-id="32789-114">Sample code: Get event data</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook): string {
  // Get the first table in the "Keys" worksheet.
  let table = workbook.getWorksheet('Keys').getTables()[0];

  // Get the rows in the event table.
  let range = table.getRangeBetweenHeaderAndTotal();
  let rows = range.getValues();

  // Save each row as an EventData object. This lets them be passed through Power Automate.
  let records: EventData[] = [];
  for (let row of rows) {
    let [eventId, date, location, capacity] = row;
    records.push({
      eventId: eventId as string,
      date: date as number,
      location: location as string,
      capacity: capacity as number
    })
  }

  // Log the event data to the console and return it for a flow.
  let stringResult = JSON.stringify(records);
  console.log(stringResult);
  return stringResult;
}

// An interface representing a row of event data.
interface EventData {
  eventId: string
  date: number
  location: string
  capacity: number
}
```

## <a name="sample-code-validate-speaker-registrations"></a><span data-ttu-id="32789-115">Пример кода: Проверка регистрации спикеров</span><span class="sxs-lookup"><span data-stu-id="32789-115">Sample code: Validate speaker registrations</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook, keys: string): string {
  // Get the first table in the "Transactions" worksheet.
  let table = workbook.getWorksheet('Transactions').getTables()[0];

  // Clear the existing formatting in the table.
  let range = table.getRangeBetweenHeaderAndTotal();
  range.clear(ExcelScript.ClearApplyTo.formats);

  // Compare the data in the table to the keys passed into the script.
  let keysObject = JSON.parse(keys) as EventData[];
  let speakerSlotsRemaining = keysObject.map(value => value.capacity);
  let overallMatch = true;

  // Iterate over every row looking for differences from the other worksheet.
  let rows = range.getValues();
  for (let i = 0; i < rows.length; i++) {
    let row = rows[i];
    let [eventId, date, location, capacity] = row;
    let match = false;

    // Look at each key provided for a matching Event ID.
    for (let keyIndex = 0; keyIndex < keysObject.length; keyIndex++) {
      let event = keysObject[keyIndex];
      if (event.eventId === eventId) {
        match = true;
        speakerSlotsRemaining[keyIndex]--;
        // If there's a match on the event ID, look for things that don't match and highlight them.
        if (event.date !== date) {
          overallMatch = false;
          range.getCell(i, 1).getFormat()
            .getFill()
            .setColor("FFFF00");
        }
        if (event.location !== location) {
          overallMatch = false;
          range.getCell(i, 2).getFormat()
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
  } else if (speakerSlotsRemaining.find(remaining => remaining < 0)){
    returnString = "Event potentially overbooked. Please review."
  }

  console.log("Returning: " + returnString);
  return returnString;
}

// An interface representing a row of event data.
interface EventData {
  eventId: string
  date: number
  location: string
  capacity: number
}
```

## <a name="power-automate-flow-check-for-inconsistencies-across-the-workbooks"></a><span data-ttu-id="32789-116">Power Automate: проверка несоответствий в книгах</span><span class="sxs-lookup"><span data-stu-id="32789-116">Power Automate flow: Check for inconsistencies across the workbooks</span></span>

<span data-ttu-id="32789-117">Этот поток извлекает сведения о событиях из первой книги и использует эти данные для проверки второй книги.</span><span class="sxs-lookup"><span data-stu-id="32789-117">This flow extracts the event information from the first workbook and uses that data to validate the second workbook.</span></span>

1. <span data-ttu-id="32789-118">Вопишите [Power Automate](https://flow.microsoft.com) и создайте новый поток **мгновенных облаков.**</span><span class="sxs-lookup"><span data-stu-id="32789-118">Sign into [Power Automate](https://flow.microsoft.com) and create a new **Instant cloud flow**.</span></span>
1. <span data-ttu-id="32789-119">Выберите **вручную вызвать поток и** нажмите **кнопку Создать**.</span><span class="sxs-lookup"><span data-stu-id="32789-119">Select **Manually trigger a flow** and press **Create**.</span></span>
1. <span data-ttu-id="32789-120">Добавьте новый **шаг,** использующий **соединителю Excel Online (Бизнес)** с действием **сценария Run.**</span><span class="sxs-lookup"><span data-stu-id="32789-120">Add a **New step** that uses the **Excel Online (Business)** connector with the **Run script** action.</span></span> <span data-ttu-id="32789-121">Используйте следующие значения для действия:</span><span class="sxs-lookup"><span data-stu-id="32789-121">Use the following values for the action:</span></span>
    * <span data-ttu-id="32789-122">**Расположение**: OneDrive для бизнеса</span><span class="sxs-lookup"><span data-stu-id="32789-122">**Location**: OneDrive for Business</span></span>
    * <span data-ttu-id="32789-123">**Библиотека документов**: OneDrive</span><span class="sxs-lookup"><span data-stu-id="32789-123">**Document Library**: OneDrive</span></span>
    * <span data-ttu-id="32789-124">**Файл**: event-data.xlsx [(выбранный с помощью выбора файла)](../../testing/power-automate-troubleshooting.md#select-workbooks-with-the-file-browser-control)</span><span class="sxs-lookup"><span data-stu-id="32789-124">**File**: event-data.xlsx ([selected with the file chooser](../../testing/power-automate-troubleshooting.md#select-workbooks-with-the-file-browser-control))</span></span>
    * <span data-ttu-id="32789-125">**Сценарий:** Получить данные событий</span><span class="sxs-lookup"><span data-stu-id="32789-125">**Script**: Get event data</span></span>

    :::image type="content" source="../../images/cross-reference-flow-1.png" alt-text="Завершенный соедините Excel Online (Бизнес) для первого сценария в Power Automate.":::

1. <span data-ttu-id="32789-127">Добавьте второй **новый** шаг, использующий **соединителю Excel Online (Бизнес)** с действием **сценария Run.**</span><span class="sxs-lookup"><span data-stu-id="32789-127">Add a second **New step** that uses the **Excel Online (Business)** connector with the **Run script** action.</span></span> <span data-ttu-id="32789-128">Используйте следующие значения для действия:</span><span class="sxs-lookup"><span data-stu-id="32789-128">Use the following values for the action:</span></span>
    * <span data-ttu-id="32789-129">**Расположение**: OneDrive для бизнеса</span><span class="sxs-lookup"><span data-stu-id="32789-129">**Location**: OneDrive for Business</span></span>
    * <span data-ttu-id="32789-130">**Библиотека документов**: OneDrive</span><span class="sxs-lookup"><span data-stu-id="32789-130">**Document Library**: OneDrive</span></span>
    * <span data-ttu-id="32789-131">**Файл**: speaker-registration.xlsx [(выбранный с помощью выбора файла)](../../testing/power-automate-troubleshooting.md#select-workbooks-with-the-file-browser-control)</span><span class="sxs-lookup"><span data-stu-id="32789-131">**File**: speaker-registration.xlsx ([selected with the file chooser](../../testing/power-automate-troubleshooting.md#select-workbooks-with-the-file-browser-control))</span></span>
    * <span data-ttu-id="32789-132">**Сценарий:** Проверка регистрации спикера</span><span class="sxs-lookup"><span data-stu-id="32789-132">**Script**: Validate speaker registration</span></span>

    :::image type="content" source="../../images/cross-reference-flow-2.png" alt-text="Завершенный соедините Excel Online (Бизнес) для второго сценария в Power Automate.":::
1. <span data-ttu-id="32789-134">В этом примере Outlook как клиент электронной почты.</span><span class="sxs-lookup"><span data-stu-id="32789-134">This sample uses Outlook as the email client.</span></span> <span data-ttu-id="32789-135">Вы можете использовать любые соединители электронной почты Power Automate поддерживает.</span><span class="sxs-lookup"><span data-stu-id="32789-135">You could use any email connector Power Automate supports.</span></span> <span data-ttu-id="32789-136">Добавьте новый **шаг,** использующий **соединителю Office 365 Outlook** и действие Отправка и электронная почта **(V2).**</span><span class="sxs-lookup"><span data-stu-id="32789-136">Add a **New step** that uses the **Office 365 Outlook** connector and the **Send and email (V2)** action.</span></span> <span data-ttu-id="32789-137">Используйте следующие значения для действия:</span><span class="sxs-lookup"><span data-stu-id="32789-137">Use the following values for the action:</span></span>
    * <span data-ttu-id="32789-138">**Чтобы:** ваша тестовая учетная запись электронной почты (или личная электронная почта)</span><span class="sxs-lookup"><span data-stu-id="32789-138">**To**: Your test email account (or personal email)</span></span>
    * <span data-ttu-id="32789-139">**Subject:** Результаты проверки событий</span><span class="sxs-lookup"><span data-stu-id="32789-139">**Subject**: Event validation results</span></span>
    * <span data-ttu-id="32789-140">**Body**: result _(динамическое содержимое из **сценария Run 2)**_</span><span class="sxs-lookup"><span data-stu-id="32789-140">**Body**: result (_dynamic content from **Run script 2**_)</span></span>

    :::image type="content" source="../../images/cross-reference-flow-3.png" alt-text="Завершенный соедините Office 365 Outlook в Power Automate.":::
1. <span data-ttu-id="32789-142">Сохраните поток, а затем **выберите Тест,** чтобы попробовать его. Вы должны получить сообщение электронной почты с сообщением "Обнаружено несоответствие.</span><span class="sxs-lookup"><span data-stu-id="32789-142">Save the flow, then select **Test** to try it out. You should receive an email saying "Mismatch found.</span></span> <span data-ttu-id="32789-143">Данные требуют проверки".</span><span class="sxs-lookup"><span data-stu-id="32789-143">Data requires your review."</span></span> <span data-ttu-id="32789-144">Это означает, что между строками вspeaker-registrations.xlsx **и** строками вevent-data.xlsx **.**</span><span class="sxs-lookup"><span data-stu-id="32789-144">This indicates there are differences between rows in **speaker-registrations.xlsx** and rows in **event-data.xlsx**.</span></span> <span data-ttu-id="32789-145">Откройте **speaker-registrations.xlsx,** чтобы увидеть несколько выделенных ячеек, где возможны проблемы с перечислениями регистрации динамиков.</span><span class="sxs-lookup"><span data-stu-id="32789-145">Open **speaker-registrations.xlsx** to see several highlighted cells where there are potential problems with the speaker registration listings.</span></span>
