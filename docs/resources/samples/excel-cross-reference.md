---
title: Перекрестные Excel файлы с Power Automate
description: Узнайте, как использовать Office и Power Automate для перекрестной ссылки и формата Excel файла.
ms.date: 06/29/2021
ms.localizationpriority: medium
ms.openlocfilehash: adeb84140cb9884309c9f37854a29fc4d59b17ed
ms.sourcegitcommit: d3ed4bdeeba805d97c930394e172e8306a0cf484
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/15/2021
ms.locfileid: "59332980"
---
# <a name="cross-reference-excel-files-with-power-automate"></a>Перекрестные Excel файлы с Power Automate

В этом решении показано, как сравнить данные между двумя Excel файлами, чтобы найти несоответствия. Он использует Office скрипты для анализа данных и Power Automate для связи между книгами.

## <a name="example-scenario"></a>Пример сценария

Вы координатор событий, который составляет расписание докладчиков для предстоящих конференций. Данные событий будут храниться в одной таблице, а регистры динамиков - в другой. Чтобы обеспечить синхронизацию двух книг, для выделения потенциальных проблем используется поток с Office скриптами.

## <a name="sample-excel-files"></a>Пример Excel файлов

Скачайте следующие файлы, чтобы получить готовые к использованию книги для примера.

1. <a href="event-data.xlsx">event-data.xlsx</a>
1. <a href="speaker-registrations.xlsx">speaker-registrations.xlsx</a>

Добавьте следующие скрипты, чтобы попробовать пример самостоятельно!

## <a name="sample-code-get-event-data"></a>Пример кода: получить данные событий

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

## <a name="sample-code-validate-speaker-registrations"></a>Пример кода: Проверка регистрации спикеров

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

## <a name="power-automate-flow-check-for-inconsistencies-across-the-workbooks"></a>Power Automate: проверка несоответствий в книгах

Этот поток извлекает сведения о событиях из первой книги и использует эти данные для проверки второй книги.

1. Вопишите [Power Automate](https://flow.microsoft.com) и создайте новый поток **мгновенных облаков.**
1. Выберите **вручную вызвать поток и** выберите **Создать**.
1. Добавьте новый **шаг,** использующий **соединителю Excel Online (Бизнес)** с действием **сценария Run.** Используйте следующие значения для действия.
    * **Расположение**: OneDrive для бизнеса
    * **Библиотека документов**: OneDrive
    * **Файл**: event-data.xlsx [(выбранный с помощью выбора файла)](../../testing/power-automate-troubleshooting.md#select-workbooks-with-the-file-browser-control)
    * **Сценарий:** Получить данные событий

    :::image type="content" source="../../images/cross-reference-flow-1.png" alt-text="Завершенный соедините Excel Online (Бизнес) для первого сценария в Power Automate.":::

1. Добавьте второй **новый** шаг, использующий **соединителю Excel Online (Бизнес)** с действием **сценария Run.** Используйте следующие значения для действия.
    * **Расположение**: OneDrive для бизнеса
    * **Библиотека документов**: OneDrive
    * **Файл**: speaker-registration.xlsx [(выбранный с помощью выбора файла)](../../testing/power-automate-troubleshooting.md#select-workbooks-with-the-file-browser-control)
    * **Сценарий:** Проверка регистрации спикера

    :::image type="content" source="../../images/cross-reference-flow-2.png" alt-text="Завершенный соедините Excel Online (Бизнес) для второго сценария в Power Automate.":::
1. В этом примере Outlook как клиент электронной почты. Вы можете использовать любые соединители электронной почты Power Automate поддерживает. Добавьте новый **шаг,** использующий **соединителю Office 365 Outlook** и действие Отправка и электронная почта **(V2).** Используйте следующие значения для действия.
    * **Чтобы:** ваша тестовая учетная запись электронной почты (или личная электронная почта)
    * **Subject:** Результаты проверки событий
    * **Body**: result _(динамическое содержимое из **сценария Run 2)**_

    :::image type="content" source="../../images/cross-reference-flow-3.png" alt-text="Завершенный соедините Office 365 Outlook в Power Automate.":::
1. Сохраните поток. Используйте **кнопку Test** на странице редактора потока или запустите поток через вкладку **Мои потоки.** Не забудьте разрешить доступ при запросе.
1. Вы должны получить сообщение электронной почты с сообщением "Обнаружено несоответствие. Данные требуют проверки". Это означает, что между строками вspeaker-registrations.xlsx **и** строками вevent-data.xlsx **.** Откройте **speaker-registrations.xlsx,** чтобы увидеть несколько выделенных ячеек, где возможны проблемы с перечислениями регистрации динамиков.
