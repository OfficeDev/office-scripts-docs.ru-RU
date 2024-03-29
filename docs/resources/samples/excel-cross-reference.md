---
title: Перекрестная ссылка на файлы Excel с помощью Power Automate
description: Узнайте, как использовать сценарии Office и Power Automate для перекрестной ссылки и форматирования файла Excel.
ms.date: 06/06/2022
ms.localizationpriority: medium
ms.openlocfilehash: b32249dc7cb1e8c1b841a4db6caaff3b4d2998ec
ms.sourcegitcommit: a6504f8b0d6b717457c6e0b5306c35ad3900914e
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/02/2022
ms.locfileid: "67572677"
---
# <a name="cross-reference-excel-files-with-power-automate"></a>Перекрестная ссылка на файлы Excel с помощью Power Automate

В этом решении показано, как сравнить данные в двух файлах Excel для поиска несоответствий. Он использует сценарии Office для анализа данных и Power Automate для обмена данными между книгами.

Этот пример передает данные между книгами с помощью объектов [JSON](https://www.w3schools.com/whatis/whatis_json.asp) . Дополнительные сведения о работе с JSON см. в статье "Использование JSON для передачи данных в скрипты [Office и из них"](../../develop/use-json.md).

## <a name="example-scenario"></a>Пример сценария

Вы являетесь координатором событий, который запланирует докладчиков для предстоящих конференций. Данные события будут храниться в одной электронной таблице, а регистрация говорящего — в другой. Чтобы обеспечить синхронизацию двух книг, используйте поток со скриптами Office, чтобы выделить возможные проблемы.

## <a name="sample-excel-files"></a>Примеры файлов Excel

Скачайте следующие файлы, чтобы получить готовые к использованию книги для примера.

1. [event-data.xlsx](event-data.xlsx)
1. [speaker-registrations.xlsx](speaker-registrations.xlsx)

Добавьте следующие скрипты, чтобы попробовать пример самостоятельно!

## <a name="sample-code-get-event-data"></a>Пример кода: получение данных о событиях

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

## <a name="sample-code-validate-speaker-registrations"></a>Пример кода: проверка регистрации говорящего

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

## <a name="power-automate-flow-check-for-inconsistencies-across-the-workbooks"></a>Поток Power Automate: проверка несогласованности между книгами

Этот поток извлекает сведения о событии из первой книги и использует эти данные для проверки второй книги.

1. Войдите [в Power Automate](https://flow.microsoft.com) и создайте новый **поток мгновенного облака**.
1. Выберите **"Вручную активировать поток" и** нажмите кнопку **"Создать"**.
1. Добавьте новый **шаг, использующий** соединитель **Excel Online (business)** с действием **запуска скрипта** . Используйте следующие значения для действия.
    * **Расположение**: OneDrive для бизнеса
    * **Библиотека документов**: OneDrive
    * **Файл**: event-data.xlsx ([выбранный с помощью выбора файла](../../testing/power-automate-troubleshooting.md#select-workbooks-with-the-file-browser-control))
    * **Скрипт**: получение данных о событиях

    :::image type="content" source="../../images/cross-reference-flow-1.png" alt-text="Завершенный соединитель Excel Online (business) для первого скрипта в Power Automate.":::

1. Добавьте второй новый **шаг,** использующий соединитель **Excel Online (business)** с действием **запуска скрипта** . При этом в качестве входных данных для скрипта проверки данных событий используются возвращаемые значения из скрипта получения **данных событий.**  Используйте следующие значения для действия.
    * **Расположение**: OneDrive для бизнеса
    * **Библиотека документов**: OneDrive
    * **Файл**: speaker-registration.xlsx ([выбранный с помощью выбора файла](../../testing/power-automate-troubleshooting.md#select-workbooks-with-the-file-browser-control))
    * **Сценарий**: проверка регистрации говорящего
    * **keys**: result (_динамическое содержимое из **скрипта запуска**_)

    :::image type="content" source="../../images/cross-reference-flow-2.png" alt-text="Завершенный соединитель Excel Online (business) для второго скрипта в Power Automate.":::
1. В этом примере в качестве почтового клиента используется Outlook. Вы можете использовать любой соединитель электронной почты, поддерживаемый Power Automate. Добавьте новый **шаг,** использующий **соединитель Office 365 Outlook** и действие **отправки и** отправки электронной почты (V2). При этом в качестве основного содержимого сообщения электронной почты используются возвращаемые значения из скрипта проверки регистрации говорящего. Используйте следующие значения для действия.
    * **To**: Ваша тестовая учетная запись электронной почты (или личная электронная почта)
    * **Тема**: результаты проверки событий
    * **Текст**: результат (_динамическое содержимое из **скрипта выполнения 2**_)

    :::image type="content" source="../../images/cross-reference-flow-3.png" alt-text="Завершенный Office 365 Outlook в Power Automate.":::
1. Сохраните поток. Нажмите **кнопку "** Тест" на странице редактора потоков или запустите поток на **вкладке "Мои потоки** ". Не забудьте разрешить доступ при появлении запроса.
1. Вы должны получить сообщение электронной почты с сообщением "Несоответствие найдено. Данные требуют проверки". Это означает, что между строками в **speaker-registrations.xlsxи** строками в **event-data.xlsx.** Откройте **speaker-registrations.xlsx** , чтобы увидеть несколько выделенных ячеек, в которых могут быть проблемы с регистрацией говорящего.
