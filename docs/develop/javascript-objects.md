---
title: Использование встроенных объектов JavaScript в сценариях Office
description: Вызов встроенных API JavaScript из сценария Office в Excel в Интернете.
ms.date: 05/17/2021
ms.localizationpriority: medium
ms.openlocfilehash: bf12a405814bb626a72c1de4f4c75462ce0018ec
ms.sourcegitcommit: d3ed4bdeeba805d97c930394e172e8306a0cf484
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/15/2021
ms.locfileid: "59327691"
---
# <a name="use-built-in-javascript-objects-in-office-scripts"></a>Использование встроенных объектов JavaScript в Office скриптах

JavaScript предоставляет несколько встроенных объектов, которые можно использовать в Office скриптах, независимо от того, создаете ли вы сценарии в JavaScript или [TypeScript](../overview/code-editor-environment.md) (суперсети JavaScript). В этой статье описывается, как можно использовать некоторые встроенные объекты JavaScript в Office скриптов для Excel в Интернете.

> [!NOTE]
> Полный список всех встроенных объектов JavaScript см. в статье Mozilla's [Standard built-in objects.](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects)

## <a name="array"></a>Массив

Объект [Array](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Array) предоставляет стандартный способ работы с массивами в скрипте. Хотя массивы являются стандартными конструкциями JavaScript, они относятся к Office скриптам двумя основными способами: диапазонами и коллекциями.

### <a name="work-with-ranges"></a>Работа с диапазонами

Диапазоны содержат несколько двухмерных массивов, которые непосредственно соеряду с ячейками в этом диапазоне. Эти массивы содержат конкретные сведения о каждой ячейке в этом диапазоне. Например, возвращает все значения в этих ячейках (с строками и столбцами сопоставления двухмерных массивов в строки и столбцы этого `Range.getValues` подсети). `Range.getFormulas` и `Range.getNumberFormats` являются другими часто используемыми методами, возвращая массивы, такие как `Range.getValues` .

В следующем скрипте выполняется поиск диапазона **A1:D4** для любого формата номеров, содержащего "$". Сценарий задает цвет заполнения в этих ячейках на "желтый".

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

Многие Excel содержатся в коллекции. Коллекция управляется API Office скриптов и выставляется в качестве массива. Например, все [фигуры](/javascript/api/office-scripts/excelscript/excelscript.shape) в таблице содержатся в возвращаемом `Shape[]` `Worksheet.getShapes` методом методе. Этот массив можно использовать для чтения значений из коллекции или для доступа к определенным объектам из методов родительского `get*` объекта.

> [!NOTE]
> Не добавляйте или удаляйте объекты из этих массивов коллекции вручную. Используйте методы для родительских объектов и `add` методы для объектов типа `delete` коллекции. Например, добавьте [таблицу](/javascript/api/office-scripts/excelscript/excelscript.table) в [таблицу](/javascript/api/office-scripts/excelscript/excelscript.worksheet) с `Worksheet.addTable` методом и удалите `Table` использование `Table.delete` .

В следующем скрипте региструется тип каждой фигуры текущего таблицы.

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

Следующий скрипт удаляет старейшую фигуру текущего таблицы.

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

Объект [Date](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Date) предоставляет стандартный способ работы с датами в скрипте. `Date.now()` создает объект с текущей датой и временем, что полезно при добавлении в запись данных скрипта.

Следующий сценарий добавляет текущую дату в таблицу. Обратите внимание, что с помощью метода Excel распознает значение как дату и автоматически меняет формат `toLocaleDateString` номера ячейки.

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

В [разделе Работа с датами](../resources/samples/excel-samples.md#dates) в примерах больше сценариев, связанных с датами.

## <a name="math"></a>математика;

Объект [Math](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Math) предоставляет методы и константы для общих математических операций. Они предоставляют множество функций, Excel в Excel, без необходимости использования двигателя вычислений книги. Это позволяет сохранить скрипт от необходимости запрашивать книгу, что повышает производительность.

Следующий скрипт использует `Math.min` для поиска и входа наименьшее число в **диапазоне A1:D4.** Обратите внимание, что в этом примере предполагается, что весь диапазон содержит только числа, а не строки.

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

Office Скрипты не поддерживают использование внешних сторонних библиотек. В скрипте можно использовать только встроенные объекты JavaScript и API Office скриптов.

## <a name="see-also"></a>См. также

- [Стандартные встроенные объекты](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects)
- [Office Среда редактора кода скриптов](../overview/code-editor-environment.md)
