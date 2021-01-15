---
title: Повышение производительности сценариев Office
description: Создавайте более быстрые сценарии, понимая связь между книгой Excel и сценарием.
ms.date: 06/15/2020
localization_priority: Normal
ms.openlocfilehash: ce50a6fd7ad02ddcd2dd304be8b4dd8fa3d0acf3
ms.sourcegitcommit: 7580dcb8f2f97974c2a9cce25ea30d6526730e28
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 01/14/2021
ms.locfileid: "49867872"
---
# <a name="improve-the-performance-of-your-office-scripts"></a>Повышение производительности сценариев Office

Целью сценариев Office является автоматизация часто выполняемых рядов задач, чтобы сэкономить время. Медленный сценарий может выглядеть так, будто он не ускоряет рабочий процесс. В большинстве своем сценарий будет работать безошибок. Однако существует несколько сценариев, которые можно избежать, которые могут повлиять на производительность.

Наиболее распространенная причина медленного сценария — чрезмерное взаимодействие с книгой. Сценарий выполняется на локальном компьютере, а книга существует в облаке. В определенное время сценарий синхронизирует локальные данные с данными книги. Это означает, что любые операции записи (например,) применяются к книге только при такой синхронизации за `workbook.addWorksheet()` кадром. Аналогично, в такие моменты любые операции чтения (например,) получают данные только из книги для `myRange.getValues()` сценария. В любом случае сценарий получает сведения, прежде чем он будет действовать с данными. Например, в следующем коде точно занося в журнал количество строк в используемом диапазоне.

```TypeScript
let usedRange = workbook.getActiveWorksheet().getUsedRange();
let rowCount = usedRange.getRowCount();
// The script will read the range and row count from
// the workbook before logging the information.
console.log(rowCount);
```

API сценариев Office обеспечивают точность и правильность любых данных в книге или сценарии при необходимости. Вам не нужно беспокоиться об этих синхронизациях для правильного запуска скрипта. Тем не менее, понимание этого сценария для облачного взаимодействия поможет избежать нежелательных сетевых вызовов.

## <a name="performance-optimizations"></a>Оптимизация производительности

Вы можете применять простые методы, чтобы сократить объем взаимодействия с облаком. Следующие шаблоны помогают ускорить ваши сценарии.

- Чтение данных книги один раз, а не несколько раз в цикле.
- Удалите `console.log` ненужные утверждения.
- Избегайте использования блоков try/catch.

### <a name="read-workbook-data-outside-of-a-loop"></a>Чтение данных книги вне цикла

Любой метод, который получает данные из книги, может вызвать сетевой вызов. Вместо того чтобы повторять один и тот же вызов, по возможности следует сохранять данные локально. Это особенно актуально при работе с циклами.

Рассмотрим сценарий, чтобы получить количество отрицательных чисел в используемом диапазоне таблицы. Сценарию необходимо итерировать каждую ячейку в используемом диапазоне. Для этого ему требуется диапазон, количество строк и число столбцов. Перед запуском цикла их следует сохранить в качестве локальных переменных. В противном случае каждая итерация цикла принудительно возвращает книгу.

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
> В качестве эксперимента попробуйте заменить `usedRangeValues` в цикле на `usedRange.getValues()` . При работе с большими диапазонами может потребоваться значительно больше времени.

### <a name="remove-unnecessary-consolelog-statements"></a>Удаление ненужных `console.log` заявлений

Ведение журнала консоли — это важный инструмент для отладки [сценариев.](../testing/troubleshooting.md) Однако он принудительно синхронизирует сценарий с книгой, чтобы убедиться, что зарегистрированные сведения имеют последние данные. Перед совместным использованием скрипта можно удалить ненужные утверждения ведения журнала (например, используемые для тестирования). Как правило, это не вызывает заметной проблемы с производительностью, если только данный отчет не `console.log()` находится в цикле.

### <a name="avoid-using-trycatch-blocks"></a>Избегайте использования блоков try/catch

Мы не рекомендуем использовать [ `try` / `catch` блоки](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Statements/try...catch) в рамках ожидаемого потока управления сценария. Большинство ошибок можно избежать, проверяя объекты, возвращенные из книги. Например, следующий сценарий проверяет, существует ли таблица, возвращенная книгой, перед попыткой добавления строки.

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

## <a name="case-by-case-help"></a>Справка по делу

По мере расширения платформы сценариев Office для работы с [Power Automate,](https://flow.microsoft.com/) [адаптивными](/adaptive-cards)карточками и другими функциями для разных продуктов подробности взаимодействия между скриптами и книгой становятся более сложными. Если вам нужна помощь в ускорении запуска скрипта, свяжитесь с [помощью Stack Overflow.](https://stackoverflow.com/questions/tagged/office-scripts) Не забудьте пометить свой вопрос с помощью "office-scripts", чтобы эксперты могли найти его и помочь.

## <a name="see-also"></a>См. также

- [Основные сведения о сценариях Office в Excel в Интернете](scripting-fundamentals.md)
- [Веб-документы MDN: циклы и итерация](https://developer.mozilla.org/docs/Web/JavaScript/Guide/Loops_and_iteration)