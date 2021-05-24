---
title: Рекомендации по сценариям Office
description: Предотвращение распространенных проблем и написание надежных Office скриптов, которые могут обрабатывать неожиданные входные данные или данные.
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

Эти шаблоны и методы предназначены для успешного запуска скриптов каждый раз. Используйте их, чтобы избежать распространенных ошибок, когда вы начинаете Excel рабочий процесс.

## <a name="verify-an-object-is-present"></a>Проверка на подлинность объекта

Сценарии часто зависят от определенного таблицы или таблицы, присутствуют в книге. Однако между запусками скриптов они могут быть переименованы или удалены. Проверив, существуют ли эти таблицы или таблицы, прежде чем вызывать на них методы, можно убедиться, что сценарий не заканчивается внезапно.

Следующий пример кода проверяет, присутствует ли в книге таблица "Index". Если таблица присутствует, скрипт получает диапазон и продолжает. Если его нет, скрипт регистрит пользовательское сообщение об ошибке.

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

Оператор TypeScript `?` проверяет, существует ли объект перед вызовом метода. Это может сделать код более упорядоченным, если вам не нужно делать что-либо особенное, если объект не существует.

```TypeScript
// The ? ensures that the delete() API is only called if the object exists.
workbook.getWorksheet('Index')?.delete();
```

## <a name="validate-data-and-workbook-state-first"></a>Сначала проверка состояния данных и книг

Убедитесь, что все ваши таблицы, таблицы, фигуры и другие объекты присутствуют перед работой над данными. С помощью предыдущего шаблона убедитесь, что все находится в книге и соответствует вашим ожиданиям. При этом перед написанием каких-либо данных сценарий не оставляет книгу в частичном состоянии.

В следующем сценарии необходимо иметь две таблицы с именами "Table1" и "Table2". Сценарий сначала проверяет, присутствуют ли таблицы, а затем заканчивается заявлением и соответствующим сообщением, если `return` они нет.

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

Если проверка происходит в отдельной функции, необходимо закончить сценарий, выпустив `return` заявление из `main` функции. Возвращение из субфункции не заканчивает сценарий.

Следующий сценарий имеет такое же поведение, как и предыдущий. Разница в том, что `main` функция вызывает `inputPresent` функцию для проверки всего. `inputPresent` возвращает boolean (или) для того, чтобы указать, присутствуют ли все `true` `false` необходимые входные данные. Функция `main` использует этот boolean для решения о продолжении или завершении сценария.

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

## <a name="when-to-use-a-throw-statement"></a>Когда использовать `throw` заявление

В [`throw`](https://developer.mozilla.org/docs/web/javascript/reference/statements/throw) заявлении указывается, что произошла неожиданная ошибка. Он немедленно завершает код. По большей части, вам не нужно из `throw` сценария. Обычно скрипт автоматически информирует пользователя о том, что сценарий не удалось выполнить из-за проблемы. В большинстве случаев достаточно закончить сценарий сообщением об ошибке и `return` заявлением из `main` функции.

Однако, если сценарий работает в Power Automate потока, может потребоваться остановить его продолжение. Заявление `throw` останавливает сценарий и сообщает потоку, чтобы остановить, а также.

В следующем сценарии показано, как использовать `throw` заявление в примере проверки таблицы.

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

## <a name="when-to-use-a-trycatch-statement"></a>Когда использовать `try...catch` заявление

Это утверждение является способом обнаружения сбой вызова API и продолжения [`try...catch`](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Statements/try...catch) запуска сценария.

Рассмотрим следующий фрагмент, который выполняет большое обновление данных в диапазоне.

```TypeScript
range.setValues(someLargeValues);
```

Если `someLargeValues` размер Excel для веб-службы, вызов `setValues()` не удается. Затем скрипт также сбой с ошибкой [времени запуска](../testing/troubleshooting.md#runtime-errors). Это утверждение позволяет скрипту распознавать это условие, не завершая сценарий немедленно `try...catch` и не показывая ошибку по умолчанию.

Один из способов предоставления пользователю скрипта более удобного интерфейса — это предоставление им настраиваемой ошибки. В следующем фрагменте показана информация об ошибках, которая поможет читателю в журнале дополнительных сведений об `try...catch` ошибках.

```TypeScript
try {
    range.setValues(someLargeValues);
} catch (error) {
    console.log(`The script failed to update the values at location ${range.getAddress()}. Please inspect and run again.`);
    console.log(error);
    return; // End the script (assuming this is in the main function).
}
```

Другой подход к работе с ошибками заключается в том, чтобы иметь поведение отката, которое обрабатывает случае ошибки. Следующий фрагмент использует блок, чтобы попробовать альтернативный метод разбить обновление на мелкие части и `catch` избежать ошибки.

> [!TIP]
> Полный пример обновления большого диапазона см. в статью [Write a large dataset.](../resources/samples/write-large-dataset.md)

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
> Использование `try...catch` внутри или вокруг цикла замедляет сценарий. Дополнительные сведения о производительности см. в [см. в избегайте использования `try...catch` блоков.](web-client-performance.md#avoid-using-trycatch-blocks-in-or-surrounding-loops)

## <a name="see-also"></a>См. также

- [Устранение неполадок в сценариях Office](../testing/troubleshooting.md)
- [Сведения об устранении неполадок для Power Automate с Office скриптами](../testing/power-automate-troubleshooting.md)
- [Ограничения платформы с Office скриптами](../testing/platform-limits.md)
- [Повышение производительности Office скриптов](web-client-performance.md)
