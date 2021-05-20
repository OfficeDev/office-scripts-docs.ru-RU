---
title: Использование встроенных объектов JavaScript в сценариях Office
description: Как вызвать встроенные API JavaScript из программного Office в Excel в Интернете.
ms.date: 05/17/2021
localization_priority: Normal
ms.openlocfilehash: 680dd326e357bd06e2fc66cba5bd6745bbd33c24
ms.sourcegitcommit: 4687693f02fc90a57ba30c461f35046e02e6f5fb
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 05/19/2021
ms.locfileid: "52545049"
---
# <a name="use-built-in-javascript-objects-in-office-scripts"></a>Использование встроенных объектов JavaScript в Office скриптах

JavaScript предоставляет несколько встроенных объектов, которые можно использовать в скриптах Office, независимо от того, делаете ли вы сценарий в JavaScript [или TypeScript](../overview/code-editor-environment.md) (суперсет JavaScript). В этой статье описывается, как можно использовать некоторые встроенные объекты JavaScript в Office для Excel в Интернете.

> [!NOTE]
> Полный список всех встроенных объектов JavaScript можно посмотреть в статье Mozilla [«Стандартные встроенные объекты».](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects)

## <a name="array"></a>Массив

Объект [Array](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Array) предоставляет стандартизированный способ работы с массивами в скрипте. Хотя массивы являются стандартными конструкциями JavaScript, они связаны Office скриптами двумя основными способами: диапазонами и коллекциями.

### <a name="work-with-ranges"></a>Работа с диапазонами

Диапазоны содержат несколько двумерных массивов, которые непосредственно карт на ячейки в этом диапазоне. Эти массивы содержат конкретную информацию о каждой ячейке в этом диапазоне. Например, `Range.getValues` возвращает все значения в этих ячейках (с рядами и столбцами двумерного отображения массива к строкам и столбцам этого подраздела листа). `Range.getFormulas` и `Range.getNumberFormats` другие часто используемые методы, которые возвращают массивы, такие как `Range.getValues` .

Следующий скрипт ищет диапазон **A1:D4** для любого формата числа, содержащего "$". Скрипт устанавливает цвет заполнения в этих ячейках на "желтый".

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Get the range From A1 to D4.
  let range = workbook.getActiveWorksheet().getRange("A1:D4");

  // Get the number formats for each cell in the range.
  let rangeNumberFormats = range.getNumberFormats();
  // Iterate through the arrays of rows and columns corresponding to those in the range.
  rangeNumberFormats.forEach((rowItem, rowIndex) => {
    rangeNumberFormats[rowIndex].forEach((columnItem, columnIndex) => {
      // Treat the numberFormat as a string so we can do text comparisons.
      let columnItemText = columnItem as string;
      if (columnItemText.indexOf("$") >= 0) {
        // Set the cell's fill to yellow.
        range.getCell(rowIndex, columnIndex).getFormat().getFill().setColor("yellow");
      }
    });
  });
}
```

### <a name="work-with-collections"></a>Работа с коллекциями

Многие Excel объекты содержатся в коллекции. Коллекция управляется API Office и выставлена в качестве массива. Например, все [фигуры](/javascript/api/office-scripts/excelscript/excelscript.shape) в листе содержатся в `Shape[]` возвращенной `Worksheet.getShapes` методом форме. Этот массив можно использовать для чтения значений из коллекции или получить доступ к определенным объектам из методов родительского `get*` объекта.

> [!NOTE]
> Не добавляйте и не удаляйте объекты из этих массивов коллекции вручную. Используйте `add` методы на родительских объектах и `delete` методы на объектах типа сбора. Например, добавьте [таблицу](/javascript/api/office-scripts/excelscript/excelscript.table) [в лист с](/javascript/api/office-scripts/excelscript/excelscript.worksheet) помощью метода и удалите `Worksheet.addTable` `Table` `Table.delete` использование.

Следующий скрипт регистрирует тип каждой формы в текущем листе.

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Get the current worksheet.
  let selectedSheet = workbook.getActiveWorksheet();

  // Get the shapes in this worksheet.
  let shapes = selectedSheet.getShapes();

  // Log the type of every shape in the collection.
  shapes.forEach((shape) => {
    console.log(shape.getType());
  });
}
```

Следующий скрипт удаляет старейшую форму в текущем листе.

```Typescript
function main(workbook: ExcelScript.Workbook) {
  // Get the current worksheet.
  let selectedSheet = workbook.getActiveWorksheet();

  // Get the first (oldest) shape in the worksheet.
  // Note that this script will thrown an error if there are no shapes.
  let shape = selectedSheet.getShapes()[0];

  // Remove the shape from the worksheet.
  shape.delete();
}
```

## <a name="date"></a>Дата

Объект [Дата](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Date) предоставляет стандартизированный способ работы с датами в скрипте. `Date.now()` генерирует объект с текущей датой и временем, что полезно при добавлении метки времени к вводу данных скрипта.

Следующий скрипт добавляет текущую дату в лист. Обратите внимание, что `toLocaleDateString` с помощью Excel, он распознает значение как дату и автоматически изменяет формат номера ячейки.

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Get the range for cell A1.
  let range = workbook.getActiveWorksheet().getRange("A1");

  // Get the current date and time.
  let date = new Date(Date.now());

  // Set the value at A1 to the current date, using a localized string.
  range.setValue(date.toLocaleDateString());
}
```

В [разделе «Работа с датами»](../resources/samples/excel-samples.md#dates) образцов имеется больше скриптов, связанных с датой.

## <a name="math"></a>математика;

Объект [Math предоставляет](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Math) методы и константы для общих математических операций. Они обеспечивают много функций, Excel также доступны в других местах, без необходимости использования двигателя расчета рабочей книги. Это избавляет ваш скрипт от необходимости запрашивать трудовую книжку, что повышает производительность.

Следующий скрипт используется `Math.min` для поиск и регистрации малейшего числа в **диапазоне A1:D4.** Обратите внимание, что этот пример предполагает, что весь диапазон содержит только числа, а не строки.

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Get the range from A1 to D4.
  let comparisonRange = workbook.getActiveWorksheet().getRange("A1:D4");

  // Load the range's values.
  let comparisonRangeValues = comparisonRange.getValues();

  // Set the minimum values as the first value.
  let minimum = comparisonRangeValues[0][0];

  // Iterate over each row looking for the smallest value.
  comparisonRangeValues.forEach((rowItem, rowIndex) => {
    // Iterate over each column looking for the smallest value.
    comparisonRangeValues[rowIndex].forEach((columnItem) => {
      // Use `Math.min` to set the smallest value as either the current cell's value or the previous minimum.
      minimum = Math.min(minimum, columnItem);
    });
  });

  console.log(minimum);
}

```

## <a name="use-of-external-javascript-libraries-is-not-supported"></a>Использование внешних библиотек JavaScript не поддерживается

Office Скрипты не поддерживают использование внешних, сторонних библиотек. Скрипт может использовать только встроенные объекты JavaScript и api Office скриптов.

## <a name="see-also"></a>См. также

- [Стандартные встроенные объекты](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects)
- [Office Среда редактора кода скриптов](../overview/code-editor-environment.md)
