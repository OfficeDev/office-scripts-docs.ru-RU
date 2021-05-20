---
title: Улучшение производительности ваших Office скриптов
description: Создавайте более быстрые сценарии, понимая связь между Excel и вашим скриптом.
ms.date: 05/17/2021
localization_priority: Normal
ms.openlocfilehash: 512e2108cb81cf9ac8ae98980951d5d01b3d2de9
ms.sourcegitcommit: 4687693f02fc90a57ba30c461f35046e02e6f5fb
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 05/19/2021
ms.locfileid: "52544993"
---
# <a name="improve-the-performance-of-your-office-scripts"></a>Улучшение производительности ваших Office скриптов

Целью этих Office является автоматизация обычно выполняемых ряд задач, чтобы сэкономить время. Медленный сценарий может чувствовать, что он не ускоряет ваш рабочий процесс. Большую часть времени, ваш сценарий будет прекрасно и работать, как ожидалось. Тем не менее, есть несколько, предотвратимых сценариев, которые могут повлиять на производительность.

Наиболее распространенной причиной медленного сценария является чрезмерное общение с рабочей книгой. Скрипт работает на локальной машине, в то время как рабочая книга существует в облаке. В определенное время скрипт синхронизирует свои локальные данные с данными рабочей книги. Это означает, что любые операции записи `workbook.addWorksheet()` (например), применяются к рабочей книге только тогда, когда происходит эта закулисная синхронизация. Аналогичным образом, любые операции чтения `myRange.getValues()` (например), получают данные из рабочей книги только для сценария в то время. В любом случае скрипт получает информацию, прежде чем он действует на данные. Например, следующий код точно залогит количество строк в используемом диапазоне.

```TypeScript
let usedRange = workbook.getActiveWorksheet().getUsedRange();
let rowCount = usedRange.getRowCount();
// The script will read the range and row count from
// the workbook before logging the information.
console.log(rowCount);
```

Office API-файлы скриптов гарантируют, что любые данные в рабочей книге или скрипте являются точными и точными, когда это необходимо. Вам не нужно беспокоиться об этих синхронизациях для правильного запуска скрипта. Тем не менее, осведомленность об этой связи между скриптами и облаками может помочь вам избежать ненужных сетевых звонков.

## <a name="performance-optimizations"></a>Оптимизация производительности

Вы можете применить простые методы, чтобы помочь уменьшить связь в облаке. Следующие шаблоны помогают ускорить ваши скрипты.

- Прочитайте данные рабочей книги один раз, а не повторно в цикле.
- Удалите ненужные `console.log` операторы.
- Избегайте использования попробовать / поймать блоков.

### <a name="read-workbook-data-outside-of-a-loop"></a>Читать данные о работе вне цикла

Любой метод, который получает данные из рабочей книги, может вызвать сетевой звонок. Вместо того, чтобы неоднократно делать один и тот же вызов, вы должны сохранить данные локально, когда это возможно. Это особенно верно при работе с петлями.

Рассмотрим сценарий, чтобы получить подсчет отрицательных чисел в используемом диапазоне листа. Скрипт должен итерировать над каждой ячейкой в используемом диапазоне. Для этого ему нужен диапазон, количество строк и количество столбцов. Вы должны хранить их в качестве локальных переменных перед началом цикла. В противном случае каждая итерация цикла заставит вернуться к рабочей книге.

```TypeScript
/**
 * This script provides the count of negative numbers that are present
 * in the used range of the current worksheet.
 */
function main(workbook: ExcelScript.Workbook) {
  // Get the working range.
  let usedRange = workbook.getActiveWorksheet().getUsedRange();

  // Save the values locally to avoid repeatedly asking the workbook.
  let usedRangeValues = usedRange.getValues();

  // Start the negative number counter.
  let negativeCount = 0;

  // Iterate over the entire range looking for negative numbers.
  for (let i = 0; i < usedRangeValues.length; i++) {
    for (let j = 0; j < usedRangeValues[i].length; j++) {
      if (usedRangeValues[i][j] < 0) {
        negativeCount++;
      }
    }
  }

  // Log the negative number count to the console.
  console.log(negativeCount);
}
```

> [!NOTE]
> В качестве эксперимента попробуйте `usedRangeValues` заменить в цикле `usedRange.getValues()` с . Вы можете заметить, что запуск скрипта занимает значительно больше времени при работе с большими диапазонами.

### <a name="avoid-using-trycatch-blocks-in-or-surrounding-loops"></a>Избегайте использования `try...catch` блоков в или окружающих петель

Мы не рекомендуем использовать операторы [`try...catch`](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Statements/try...catch) в петлях или окружающих петлях. Это по той же причине, по которой следует избегать чтения данных в цикле: каждая итерация заставляет скрипт синхронизироваться с рабочей книгой, чтобы убедиться, что ошибка не была брошена. Большинство ошибок можно избежать, проверяя объекты, возвращенные из рабочей книги. Например, следующий скрипт проверяет, что таблица, возвращенная рабочей книгой, существует, прежде чем пытаться добавить строку.

```TypeScript
/**
 * This script adds a row to "MyTable", if that table is present.
 */
function main(workbook: ExcelScript.Workbook) {
  let table = workbook.getTable("MyTable");

  // Check if the table exists.
  if (table) {
    // Add the row.
    table.addRow(-1, ["2012", "Yes", "Maybe"]);
  } else {
    // Report the missing table.
    console.log("MyTable not found.");
  }
}
```

### <a name="remove-unnecessary-consolelog-statements"></a>Удалить ненужные `console.log` операторы

Консольный журнал является жизненно важным инструментом [для отладки скриптов.](../testing/troubleshooting.md) Тем не менее, это требует синхронизации скрипта с рабочей книгой, чтобы убедиться, что зарегистрированная информация является в курсе. Перед просмотром сценария следует удалить ненужные операторы журналов (например, те, которые используются для тестирования). Обычно это не вызывает заметной проблемы с производительностью, если только `console.log()` заявление не находится в цикле.

## <a name="case-by-case-help"></a>Помощь в каждом конкретном случае

По мере Office платформы Power Automate, [адаптивных карт](/adaptive-cards) [и](https://flow.microsoft.com/)других кросс-продуктов, детали общения скрипта и рабочей книги становятся все более запутанными. Если вам нужна помощь в том, чтобы сделать ваш скрипт работать быстрее, пожалуйста, пройдите [через корпорацию Майкрософт&A.](/answers/topics/office-scripts-dev.html) Не забудьте отметить ваш вопрос с "офис-скрипты-dev", чтобы эксперты могли найти его и помочь.

## <a name="see-also"></a>См. также

- [Основные сведения о сценариях Office в Excel в Интернете](scripting-fundamentals.md)
- [Веб-документы MDN: петли и итерация](https://developer.mozilla.org/docs/Web/JavaScript/Guide/Loops_and_iteration)
