---
title: Рекомендации по сценариям Office
description: Как предотвратить общие проблемы и написать надежные Office, которые могут обрабатывать неожиданные входные данные или данные.
ms.date: 05/10/2021
localization_priority: Normal
ms.openlocfilehash: 0697e6fd1fa8f437a4a585d938254deb5a05f20c
ms.sourcegitcommit: 4687693f02fc90a57ba30c461f35046e02e6f5fb
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 05/19/2021
ms.locfileid: "52546033"
---
# <a name="best-practices-in-office-scripts"></a>Рекомендации по сценариям Office

Эти шаблоны и практики разработаны, чтобы помочь вашим сценариям успешно работать каждый раз. Используйте их, чтобы избежать распространенных ловушек при запуске автоматизации рабочего Excel процесса.

## <a name="verify-an-object-is-present"></a>Проверка присутствуют на объекте

Сценарии часто полагаются на определенный лист или таблицу, присутствуют в рабочей книге. Тем не менее, они могут быть переименованы или удалены между запускается сценарий. Проверяя, существуют ли эти таблицы или листы перед вызовом методов на них, вы можете убедиться, что сценарий не заканчивается внезапно.

Следующий пример кода проверяет, присутствует ли лист «Индекс» в рабочей книге. Если лист присутствует, скрипт получает диапазон и продолжается. Если его нет, скрипт регистрирует пользовательское сообщение об ошибке.

```TypeScript
// Make sure the "Index" worksheet exists before using it.
let indexSheet = workbook.getWorksheet('Index');
if (indexSheet) {
  let range = indexSheet.getRange("A1");
  // Continue using the range...
} else {
  console.log("Index sheet not found.");
}
```

Оператор TypeScript `?` проверяет, существует ли объект перед вызовом метода. Это может сделать ваш код более упорядоченным, если вам не нужно делать ничего особенного, когда объект не существует.

```TypeScript
// The ? ensures that the delete() API is only called if the object exists.
workbook.getWorksheet('Index')?.delete();
```

## <a name="validate-data-and-workbook-state-first"></a>Проверка данных и состояние рабочей книги в первую очередь

Убедитесь, что все ваши листы, таблицы, формы и другие объекты присутствуют перед работой над данными. Используя предыдущую схему, проверьте, все ли в рабочей книге и соответствует вашим ожиданиям. Это делается до того, как будут написаны какие-либо данные, и ваш скрипт не оставит трудовую книжку в частичном состоянии.

Следующий скрипт требует, чтобы присутствовали две таблицы под названием "Table1" и "Table2". Скрипт сначала проверяет, присутствуют ли таблицы, а затем заканчивается `return` выпиской и соответствующим сообщением, если это не так.

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // These tables must be in the workbook for the script.
  const TargetTableName = 'Table1';
  const SourceTableName = 'Table2';

  // Get the table objects.
  let targetTable = workbook.getTable(TargetTableName);
  let sourceTable = workbook.getTable(SourceTableName);

  // Check if the tables are there.
  if (!targetTable || !sourceTable) {
    console.log(`Required tables missing - Check that both the source (${TargetTableName}) and target (${SourceTableName}) tables are present before running the script.`);
    return;
  }

  // Continue....
}
```

Если проверка происходит в отдельной функции, вы все равно должны закончить сценарий, выдав `return` выписку из `main` функции. Возвращение из подфункции не заканчивается сценарием.

Следующий скрипт имеет такое же поведение, как и предыдущий. Разница в том, что `main` функция вызывает `inputPresent` функцию, чтобы проверить все. `inputPresent` возвращает boolean (или `true` ) для того чтобы `false` указать присутствуют ли все необходимые входы. Функция `main` использует этот boolean, чтобы принять решение о продолжении или прекращении сценария.

```TypeScript
function main(workbook: ExcelScript.Workbook) {

  // Get the table objects.
  if (!inputPresent(workbook)) {
    return;
  }

  // Continue....
}

function inputPresent( workbook: ExcelScript.Workbook): boolean {
  // These tables must be in the workbook for the script.
  const TargetTableName = 'Table1';
  const SourceTableName = 'Table2';

  // Get the table objects.
  let targetTable = workbook.getTable(TargetTableName);
  let sourceTable = workbook.getTable(SourceTableName);

  // Check if the tables are there.
  if (!targetTable || !sourceTable) {
    console.log(`Required tables missing - Check that both the source (${TargetTableName}) and target (${SourceTableName}) tables are present before running the script.`);
    return false;
  }

  return true;
}
```

## <a name="when-to-use-a-throw-statement"></a>Когда использовать `throw` выписку

В [`throw`](https://developer.mozilla.org/docs/web/javascript/reference/statements/throw) заявлении указывается, что произошла неожиданная ошибка. Он немедленно завершает код. По большей части, вам не нужно из `throw` вашего сценария. Как правило, скрипт автоматически информирует пользователя о том, что скрипт не был запущен из-за проблемы. В большинстве случаев достаточно закончить сценарий сообщением об ошибке и `return` выпиской из `main` функции.

Однако, если скрипт работает как часть Power Automate потока, вы можете остановить поток от продолжения. Заявление `throw` останавливает сценарий и говорит поток, чтобы остановить, а также.

В следующем скрипте показано, как использовать `throw` выписку в примере проверки таблицы.

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // These tables must be in the workbook for the script.
  const TargetTableName = 'Table1';
  const SourceTableName = 'Table2';

  // Get the table objects.
  let targetTable = workbook.getTable(TargetTableName);
  let sourceTable = workbook.getTable(SourceTableName);

  // Check if the tables are there.
  if (!targetTable || !sourceTable) {
    // Immediately end the script with an error.
    throw `Required tables missing - Check that both the source (${TargetTableName}) and target (${SourceTableName}) tables are present before running the script.`;
  }
  
```

## <a name="when-to-use-a-trycatch-statement"></a>Когда использовать `try...catch` выписку

Заявление [`try...catch`](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Statements/try...catch) — это способ определить, не удается ли вызов API, и продолжить запуск скрипта.

Рассмотрим следующий фрагмент, который выполняет большое обновление данных на диапазоне.

```TypeScript
range.setValues(someLargeValues);
```

Если `someLargeValues` больше, чем Excel для интернета может обрабатывать, вызов `setValues()` не удается. Скрипт затем также не удается с [ошибкой времени выполнения](../testing/troubleshooting.md#runtime-errors). Заявление `try...catch` позволяет скрипту распознать это условие, без немедленной прекращения сценария и отображения ошибки по умолчанию.

Один из подходов к предоставлению пользователю скрипта лучшего опыта заключается в том, чтобы представить им пользовательское сообщение об ошибке. Следующий фрагмент показывает заявление `try...catch` регистрации больше информации об ошибках, чтобы лучше помочь читателю.

```TypeScript
try {
    range.setValues(someLargeValues);
} catch (error) {
    console.log(`The script failed to update the values at location ${range.getAddress()}. Please inspect and run again.`);
    console.log(error);
    return; // End the script (assuming this is in the main function).
}
```

Другой подход к работе с ошибками заключается в том, чтобы иметь обратное поведение, которое обрабатывает случай ошибки. Следующий фрагмент использует `catch` блок, чтобы попробовать альтернативный метод разбить обновление на более мелкие части и избежать ошибки.

> [!TIP]
> Полный пример обновления большого диапазона можно найти в большом [наборе данных.](../resources/samples/write-large-dataset.md)

```TypeScript
try {
    range.setValues(someLargeValues);
} catch (error) {
    console.log(`The script failed to update the values at location ${range.getAddress()}. Trying a different approach.`);
    handleUpdatesInSmallerBatches(someLargeValues);
}

// Continue...
}
```

> [!NOTE]
> Использование `try...catch` внутри или вокруг цикла замедляет работу скрипта. Для получения дополнительной информации о производительности [см. `try...catch` ](web-client-performance.md#avoid-using-trycatch-blocks-in-or-surrounding-loops)

## <a name="see-also"></a>См. также

- [Устранение неполадок в сценариях Office](../testing/troubleshooting.md)
- [Информация о устранении неполадок для Power Automate с помощью Office скриптов](../testing/power-automate-troubleshooting.md)
- [Ограничения платформы с Office скриптами](../testing/platform-limits.md)
- [Улучшение производительности ваших Office скриптов](web-client-performance.md)
