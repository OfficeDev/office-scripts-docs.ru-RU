---
title: Повышение производительности Office скриптов
description: Создайте более быстрые сценарии, понимая связь между Excel книгой и скриптом.
ms.date: 05/17/2021
ms.localizationpriority: medium
ms.openlocfilehash: 2deb417d41c4be663efaf83735459eab26146410
ms.sourcegitcommit: 7023b9e23499806901a5ecf8ebc460b76887cca6
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/31/2022
ms.locfileid: "64585634"
---
# <a name="improve-the-performance-of-your-office-scripts"></a>Повышение производительности Office скриптов

Цель Office скриптов — автоматизировать обычно выполняемые серии задач, чтобы сэкономить время. При медленном сценарии может быть ощущение, что он не ускоряет рабочий процесс. В большинстве своем сценарий будет работать в отличном порядке и работать, как и ожидалось. Однако существует несколько сценариев, которые могут повлиять на производительность.

Наиболее частой причиной медленного сценария является чрезмерная связь с книгой. Сценарий выполняется на локальном компьютере, а книга существует в облаке. В определенное время сценарий синхронизирует локальные данные с данными книги. Это означает, что любые операции записи ( `workbook.addWorksheet()`например) применяются к книге только тогда, когда происходит эта закулисье синхронизация. Кроме того, любые операции чтения ( `myRange.getValues()`например) получают данные из книги для скрипта в это время. В любом случае сценарий извлекает сведения, прежде чем он будет действовать на данных. Например, в следующем коде будет точно входить число строк в используемом диапазоне.

```TypeScript
let usedRange = workbook.getActiveWorksheet().getUsedRange();
let rowCount = usedRange.getRowCount();
// The script will read the range and row count from
// the workbook before logging the information.
console.log(rowCount);
```

Office API скриптов гарантируют, что любые данные в книге или скрипте точны и в случае необходимости устарели. Вам не нужно беспокоиться об этих синхронизациях для правильного запуска скрипта. Однако осведомленность об этом сообщении от скрипта к облаку поможет избежать нежелательных сетевых вызовов.

## <a name="performance-optimizations"></a>Оптимизация производительности

Вы можете применить простые методы, чтобы уменьшить сообщение с облаком. Следующие шаблоны помогают ускорить скрипты.

- Чтение данных книг один раз, а не несколько раз в цикле.
- Удаление ненужных `console.log` заявлений.
- Избегайте использования блоков try/catch.

### <a name="read-workbook-data-outside-of-a-loop"></a>Чтение данных книг за пределами цикла

Любой метод, который получает данные из книги, может вызвать сетевой вызов. Вместо того, чтобы повторять один и тот же вызов, необходимо сохранять данные локально по мере возможности. Это особенно актуально при работе с циклами.

Рассмотрим сценарий, чтобы получить количество отрицательных чисел в используемом диапазоне таблицы. Сценарию необходимо итерировать каждую ячейку используемого диапазона. Для этого ему необходимы диапазон, количество строк и количество столбцов. Перед запуском цикла следует хранить эти параметры в качестве локальных переменных. В противном случае каждая итерация цикла заставит вернуться к книге.

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
> В качестве эксперимента попробуйте заменить в `usedRangeValues` цикле .`usedRange.getValues()` При работе с большими диапазонами скрипт может работать значительно дольше.

### <a name="avoid-using-trycatch-blocks-in-or-surrounding-loops"></a>Избегайте использования `try...catch` блоков в или окружающих циклах

Мы не рекомендуем использовать заявления [`try...catch`](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Statements/try...catch) ни в циклах, ни в окружающих циклах. По той же причине следует избегать чтения данных в цикле: каждая итерация заставляет скрипт синхронизироваться с книгой, чтобы убедиться, что ошибка не была брошена. Большинство ошибок можно избежать, проверяя объекты, возвращенные из книги. Например, следующий сценарий проверяет, что таблица, возвращаемая книгой, существует перед попыткой добавить строку.

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

### <a name="remove-unnecessary-consolelog-statements"></a>Удаление ненужных заявлений `console.log`

Ведение журнала консоли — это жизненно важный инструмент для [отладки скриптов](../testing/troubleshooting.md). Однако этот сценарий должен синхронизироваться с книгой, чтобы убедиться в том, что зарегистрированные сведения устарели. Перед совместным использованием скрипта следует удалить ненужные отчеты о журнале (например, используемые для тестирования). Обычно это не вызывает заметной проблемы с производительностью, `console.log()` если заявление не находится в цикле.

## <a name="case-by-case-help"></a>Помощь в разных случаях

По мере расширения платформы Office скриптов для работы с [Power Automate](https://flow.microsoft.com/), [адаптивными](/adaptive-cards) картами и другими функциями кросс-продуктов, сведения о связи скрипта и книги становятся более сложными. Если вам нужна помощь, чтобы сделать сценарий более быстрым, обратитесь к [Microsoft Q&A](/answers/topics/office-scripts-excel-dev.html). Обязательно пометите свой вопрос с помощью "office-scripts-dev", чтобы эксперты могли найти его и помочь.

## <a name="see-also"></a>См. также

- [Основные сведения о сценариях Office в Excel для Интернета](scripting-fundamentals.md)
- [Веб-документы MDN: циклы и итерация](https://developer.mozilla.org/docs/Web/JavaScript/Guide/Loops_and_iteration)
