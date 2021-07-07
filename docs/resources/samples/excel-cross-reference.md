---
title: Перекрестные Excel файлы с Power Automate
description: Узнайте, как использовать Office и Power Automate для перекрестной ссылки и формата Excel файла.
ms.date: 06/29/2021
localization_priority: Normal
ms.openlocfilehash: 0776ce49cacecfa15339cc7c0cd4866daad789ff
ms.sourcegitcommit: 211c157ca746e266eeb079f5fa1925a1e35ab702
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/07/2021
ms.locfileid: "53313962"
---
# <a name="cross-reference-excel-files-with-power-automate"></a><span data-ttu-id="878b8-103">Перекрестные Excel файлы с Power Automate</span><span class="sxs-lookup"><span data-stu-id="878b8-103">Cross-reference Excel files with Power Automate</span></span>

<span data-ttu-id="878b8-104">В этом решении показано, как сравнить данные между двумя Excel файлами, чтобы найти несоответствия.</span><span class="sxs-lookup"><span data-stu-id="878b8-104">This solution shows how to compare data across two Excel files to find discrepancies.</span></span> <span data-ttu-id="878b8-105">Он использует Office скрипты для анализа данных и Power Automate для связи между книгами.</span><span class="sxs-lookup"><span data-stu-id="878b8-105">It uses Office Scripts to analyze data and Power Automate to communicate between the workbooks.</span></span>

## <a name="example-scenario"></a><span data-ttu-id="878b8-106">Пример сценария</span><span class="sxs-lookup"><span data-stu-id="878b8-106">Example scenario</span></span>

<span data-ttu-id="878b8-107">Вы координатор событий, который составляет расписание докладчиков для предстоящих конференций.</span><span class="sxs-lookup"><span data-stu-id="878b8-107">You're an event coordinator who is scheduling speakers for upcoming conferences.</span></span> <span data-ttu-id="878b8-108">Данные событий будут храниться в одной таблице, а регистры динамиков - в другой.</span><span class="sxs-lookup"><span data-stu-id="878b8-108">You keep the event data in one spreadsheet and the speaker registrations in another.</span></span> <span data-ttu-id="878b8-109">Чтобы обеспечить синхронизацию двух книг, для выделения потенциальных проблем используется поток с Office скриптами.</span><span class="sxs-lookup"><span data-stu-id="878b8-109">To ensure the two workbooks are kept in sync, you use a flow with Office Scripts to highlight any potential problems.</span></span>

## <a name="sample-excel-files"></a><span data-ttu-id="878b8-110">Пример Excel файлов</span><span class="sxs-lookup"><span data-stu-id="878b8-110">Sample Excel files</span></span>

<span data-ttu-id="878b8-111">Скачайте следующие файлы, чтобы получить готовые к использованию книги для примера.</span><span class="sxs-lookup"><span data-stu-id="878b8-111">Download the following files to get ready-to-use workbooks for the sample.</span></span>

1. <span data-ttu-id="878b8-112"><a href="event-data.xlsx">event-data.xlsx</a></span><span class="sxs-lookup"><span data-stu-id="878b8-112"><a href="event-data.xlsx">event-data.xlsx</a></span></span>
1. <span data-ttu-id="878b8-113"><a href="speaker-registrations.xlsx">speaker-registrations.xlsx</a></span><span class="sxs-lookup"><span data-stu-id="878b8-113"><a href="speaker-registrations.xlsx">speaker-registrations.xlsx</a></span></span>

<span data-ttu-id="878b8-114">Добавьте следующие скрипты, чтобы попробовать пример самостоятельно!</span><span class="sxs-lookup"><span data-stu-id="878b8-114">Add the following scripts to try the sample yourself!</span></span>

## <a name="sample-code-get-event-data"></a><span data-ttu-id="878b8-115">Пример кода: получить данные событий</span><span class="sxs-lookup"><span data-stu-id="878b8-115">Sample code: Get event data</span></span>

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

## <a name="sample-code-validate-speaker-registrations"></a><span data-ttu-id="878b8-116">Пример кода: Проверка регистрации спикеров</span><span class="sxs-lookup"><span data-stu-id="878b8-116">Sample code: Validate speaker registrations</span></span>

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

## <a name="power-automate-flow-check-for-inconsistencies-across-the-workbooks"></a><span data-ttu-id="878b8-117">Power Automate: проверка несоответствий в книгах</span><span class="sxs-lookup"><span data-stu-id="878b8-117">Power Automate flow: Check for inconsistencies across the workbooks</span></span>

<span data-ttu-id="878b8-118">Этот поток извлекает сведения о событиях из первой книги и использует эти данные для проверки второй книги.</span><span class="sxs-lookup"><span data-stu-id="878b8-118">This flow extracts the event information from the first workbook and uses that data to validate the second workbook.</span></span>

1. <span data-ttu-id="878b8-119">Вопишите [Power Automate](https://flow.microsoft.com) и создайте новый поток **мгновенных облаков.**</span><span class="sxs-lookup"><span data-stu-id="878b8-119">Sign into [Power Automate](https://flow.microsoft.com) and create a new **Instant cloud flow**.</span></span>
1. <span data-ttu-id="878b8-120">Выберите **вручную вызвать поток и** выберите **Создать**.</span><span class="sxs-lookup"><span data-stu-id="878b8-120">Choose **Manually trigger a flow** and select **Create**.</span></span>
1. <span data-ttu-id="878b8-121">Добавьте новый **шаг,** использующий **соединителю Excel Online (Бизнес)** с действием **сценария Run.**</span><span class="sxs-lookup"><span data-stu-id="878b8-121">Add a **New step** that uses the **Excel Online (Business)** connector with the **Run script** action.</span></span> <span data-ttu-id="878b8-122">Используйте следующие значения для действия:</span><span class="sxs-lookup"><span data-stu-id="878b8-122">Use the following values for the action:</span></span>
    * <span data-ttu-id="878b8-123">**Расположение**: OneDrive для бизнеса</span><span class="sxs-lookup"><span data-stu-id="878b8-123">**Location**: OneDrive for Business</span></span>
    * <span data-ttu-id="878b8-124">**Библиотека документов**: OneDrive</span><span class="sxs-lookup"><span data-stu-id="878b8-124">**Document Library**: OneDrive</span></span>
    * <span data-ttu-id="878b8-125">**Файл**: event-data.xlsx [(выбранный с помощью выбора файла)](../../testing/power-automate-troubleshooting.md#select-workbooks-with-the-file-browser-control)</span><span class="sxs-lookup"><span data-stu-id="878b8-125">**File**: event-data.xlsx ([selected with the file chooser](../../testing/power-automate-troubleshooting.md#select-workbooks-with-the-file-browser-control))</span></span>
    * <span data-ttu-id="878b8-126">**Сценарий:** Получить данные событий</span><span class="sxs-lookup"><span data-stu-id="878b8-126">**Script**: Get event data</span></span>

    :::image type="content" source="../../images/cross-reference-flow-1.png" alt-text="Завершенный соедините Excel Online (Бизнес) для первого сценария в Power Automate.":::

1. <span data-ttu-id="878b8-128">Добавьте второй **новый** шаг, использующий **соединителю Excel Online (Бизнес)** с действием **сценария Run.**</span><span class="sxs-lookup"><span data-stu-id="878b8-128">Add a second **New step** that uses the **Excel Online (Business)** connector with the **Run script** action.</span></span> <span data-ttu-id="878b8-129">Используйте следующие значения для действия:</span><span class="sxs-lookup"><span data-stu-id="878b8-129">Use the following values for the action:</span></span>
    * <span data-ttu-id="878b8-130">**Расположение**: OneDrive для бизнеса</span><span class="sxs-lookup"><span data-stu-id="878b8-130">**Location**: OneDrive for Business</span></span>
    * <span data-ttu-id="878b8-131">**Библиотека документов**: OneDrive</span><span class="sxs-lookup"><span data-stu-id="878b8-131">**Document Library**: OneDrive</span></span>
    * <span data-ttu-id="878b8-132">**Файл**: speaker-registration.xlsx [(выбранный с помощью выбора файла)](../../testing/power-automate-troubleshooting.md#select-workbooks-with-the-file-browser-control)</span><span class="sxs-lookup"><span data-stu-id="878b8-132">**File**: speaker-registration.xlsx ([selected with the file chooser](../../testing/power-automate-troubleshooting.md#select-workbooks-with-the-file-browser-control))</span></span>
    * <span data-ttu-id="878b8-133">**Сценарий:** Проверка регистрации спикера</span><span class="sxs-lookup"><span data-stu-id="878b8-133">**Script**: Validate speaker registration</span></span>

    :::image type="content" source="../../images/cross-reference-flow-2.png" alt-text="Завершенный соедините Excel Online (Бизнес) для второго сценария в Power Automate.":::
1. <span data-ttu-id="878b8-135">В этом примере Outlook как клиент электронной почты.</span><span class="sxs-lookup"><span data-stu-id="878b8-135">This sample uses Outlook as the email client.</span></span> <span data-ttu-id="878b8-136">Вы можете использовать любые соединители электронной почты Power Automate поддерживает.</span><span class="sxs-lookup"><span data-stu-id="878b8-136">You could use any email connector Power Automate supports.</span></span> <span data-ttu-id="878b8-137">Добавьте новый **шаг,** использующий **соединителю Office 365 Outlook** и действие Отправка и электронная почта **(V2).**</span><span class="sxs-lookup"><span data-stu-id="878b8-137">Add a **New step** that uses the **Office 365 Outlook** connector and the **Send and email (V2)** action.</span></span> <span data-ttu-id="878b8-138">Используйте следующие значения для действия:</span><span class="sxs-lookup"><span data-stu-id="878b8-138">Use the following values for the action:</span></span>
    * <span data-ttu-id="878b8-139">**Чтобы:** ваша тестовая учетная запись электронной почты (или личная электронная почта)</span><span class="sxs-lookup"><span data-stu-id="878b8-139">**To**: Your test email account (or personal email)</span></span>
    * <span data-ttu-id="878b8-140">**Subject:** Результаты проверки событий</span><span class="sxs-lookup"><span data-stu-id="878b8-140">**Subject**: Event validation results</span></span>
    * <span data-ttu-id="878b8-141">**Body**: result _(динамическое содержимое из **сценария Run 2)**_</span><span class="sxs-lookup"><span data-stu-id="878b8-141">**Body**: result (_dynamic content from **Run script 2**_)</span></span>

    :::image type="content" source="../../images/cross-reference-flow-3.png" alt-text="Завершенный соедините Office 365 Outlook в Power Automate.":::
1. <span data-ttu-id="878b8-143">Сохраните поток.</span><span class="sxs-lookup"><span data-stu-id="878b8-143">Save the flow.</span></span> <span data-ttu-id="878b8-144">Используйте **кнопку Test** на странице редактора потока или запустите поток через вкладку **Мои потоки.** Не забудьте разрешить доступ при запросе.</span><span class="sxs-lookup"><span data-stu-id="878b8-144">Use the **Test** button on the flow editor page or run the flow through your **My flows** tab. Be sure to allow access when prompted.</span></span>
1. <span data-ttu-id="878b8-145">Вы должны получить сообщение электронной почты с сообщением "Обнаружено несоответствие.</span><span class="sxs-lookup"><span data-stu-id="878b8-145">You should receive an email saying "Mismatch found.</span></span> <span data-ttu-id="878b8-146">Данные требуют проверки".</span><span class="sxs-lookup"><span data-stu-id="878b8-146">Data requires your review."</span></span> <span data-ttu-id="878b8-147">Это означает, что между строками вspeaker-registrations.xlsx **и** строками вevent-data.xlsx **.**</span><span class="sxs-lookup"><span data-stu-id="878b8-147">This indicates there are differences between rows in **speaker-registrations.xlsx** and rows in **event-data.xlsx**.</span></span> <span data-ttu-id="878b8-148">Откройте **speaker-registrations.xlsx,** чтобы увидеть несколько выделенных ячеек, где возможны проблемы с перечислениями регистрации динамиков.</span><span class="sxs-lookup"><span data-stu-id="878b8-148">Open **speaker-registrations.xlsx** to see several highlighted cells where there are potential problems with the speaker registration listings.</span></span>
