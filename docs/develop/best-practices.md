---
title: Рекомендации по сценариям Office
description: Предотвращение распространенных проблем и написание надежных Office скриптов, которые могут обрабатывать неожиданные входные данные или данные.
ms.date: 12/29/2021
ms.localizationpriority: medium
ms.openlocfilehash: 689196e1a0ca70c999ec8048de64190cbfe75581
ms.sourcegitcommit: 7023b9e23499806901a5ecf8ebc460b76887cca6
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/31/2022
ms.locfileid: "64585767"
---
# <a name="best-practices-in-office-scripts"></a>Рекомендации по сценариям Office

Эти шаблоны и методы предназначены для успешного запуска скриптов каждый раз. Используйте их, чтобы избежать распространенных ошибок, когда вы начинаете Excel рабочий процесс.

## <a name="use-the-action-recorder-to-learn-new-features"></a>Использование регистратора действий, чтобы узнать новые функции

Excel многое делает. Большинство из них можно сценарии. Регистратор действий записи Excel действия и преобразует их в код. Это самый простой способ узнать, как различные функции работают с Office скриптами. Если вам нужен код для определенного действия, переключите на регистратор действий, выполните действия, выберите **скопируйте** в виде кода и включите в скрипт результативный код.

:::image type="content" source="../images/action-recorder-copy-code.png" alt-text="Области задач средства записи действий с выделенной кнопкой &quot;Копировать как код&quot;.":::

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

В следующем сценарии необходимо иметь две таблицы с именами "Table1" и "Table2". Сценарий сначала проверяет, присутствуют ли таблицы, `return` а затем заканчивается заявлением и соответствующим сообщением, если они нет.

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

  // Continue...
}
```

Если проверка происходит в отдельной функции, необходимо закончить сценарий, `return` выпустив заявление из функции `main` . Возвращение из субфункции не заканчивает сценарий.

Следующий сценарий имеет такое же поведение, как и предыдущий. Разница в том, что функция `main` вызывает функцию `inputPresent` для проверки всего. `inputPresent` возвращает boolean (`true` или `false`) для того, чтобы указать, присутствуют ли все необходимые входные данные. Функция `main` использует этот boolean для решения о продолжении или завершении сценария.

```TypeScript
function main(workbook: ExcelScript.Workbook) {

  // Get the table objects.
  if (!inputPresent(workbook)) {
    return;
  }

  // Continue...
}

function inputPresent(workbook: ExcelScript.Workbook): boolean {
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

## <a name="when-to-use-a-throw-statement"></a>Когда использовать заявление `throw`

В [`throw`](https://developer.mozilla.org/docs/web/javascript/reference/statements/throw) заявлении указывается, что произошла неожиданная ошибка. Он немедленно завершает код. По большей части, вам не нужно из `throw` сценария. Обычно скрипт автоматически информирует пользователя о том, что сценарий не удалось выполнить из-за проблемы. В большинстве случаев `return` достаточно закончить сценарий сообщением об ошибке и заявлением из функции `main` .

Однако, если сценарий работает в Power Automate потока, может потребоваться остановить его продолжение. Заявление `throw` останавливает сценарий и сообщает потоку, чтобы остановить, а также.

В следующем сценарии показано, как использовать заявление `throw` в примере проверки таблицы.

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

## <a name="when-to-use-a-trycatch-statement"></a>Когда использовать заявление `try...catch`

Это [`try...catch`](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Statements/try...catch) утверждение является способом обнаружения сбой вызова API и продолжения запуска сценария.

Рассмотрим следующий фрагмент, который выполняет большое обновление данных в диапазоне.

```TypeScript
range.setValues(someLargeValues);
```

Если `someLargeValues` размер превышает Excel для Интернета, вызов не `setValues()` удается. Затем скрипт также сбой с ошибкой [времени запуска](../testing/troubleshooting.md#runtime-errors). Это `try...catch` утверждение позволяет скрипту распознавать это условие, не завершая сценарий немедленно и не показывая ошибку по умолчанию.

Один из способов предоставления пользователю скрипта более удобного интерфейса — это предоставление им настраиваемой ошибки. В следующем фрагменте показана `try...catch` информация об ошибках, которая поможет читателю в журнале дополнительных сведений об ошибках.

```TypeScript
try {
    range.setValues(someLargeValues);
} catch (error) {
    console.log(`The script failed to update the values at location ${range.getAddress()}. Please inspect and run again.`);
    console.log(error);
    return; // End the script (assuming this is in the main function).
}
```

Другой подход к работе с ошибками заключается в том, чтобы иметь поведение отката, которое обрабатывает случае ошибки. Следующий фрагмент использует блок `catch` , чтобы попробовать альтернативный метод разбить обновление на мелкие части и избежать ошибки.

> [!TIP]
> Полный пример обновления большого диапазона см. в статью [Write a large dataset](../resources/samples/write-large-dataset.md).

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
> Использование `try...catch` внутри или вокруг цикла замедляет сценарий. Дополнительные сведения о производительности см. в [дополнительных сведениях об избегайте использования `try...catch` блоков](web-client-performance.md#avoid-using-trycatch-blocks-in-or-surrounding-loops).

## <a name="see-also"></a>См. также

- [Устранение неполадок в сценариях Office](../testing/troubleshooting.md)
- [Сведения об устранении неполадок для Power Automate с Office скриптами](../testing/power-automate-troubleshooting.md)
- [Ограничения платформы с Office скриптами](../testing/platform-limits.md)
- [Повышение производительности Office скриптов](web-client-performance.md)
