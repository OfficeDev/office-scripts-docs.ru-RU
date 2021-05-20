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
# <a name="use-built-in-javascript-objects-in-office-scripts"></a><span data-ttu-id="faede-103">Использование встроенных объектов JavaScript в Office скриптах</span><span class="sxs-lookup"><span data-stu-id="faede-103">Use built-in JavaScript objects in Office Scripts</span></span>

<span data-ttu-id="faede-104">JavaScript предоставляет несколько встроенных объектов, которые можно использовать в скриптах Office, независимо от того, делаете ли вы сценарий в JavaScript [или TypeScript](../overview/code-editor-environment.md) (суперсет JavaScript).</span><span class="sxs-lookup"><span data-stu-id="faede-104">JavaScript provides several built-in objects that you can use in your Office Scripts, regardless of whether you're scripting in JavaScript or [TypeScript](../overview/code-editor-environment.md) (a superset of JavaScript).</span></span> <span data-ttu-id="faede-105">В этой статье описывается, как можно использовать некоторые встроенные объекты JavaScript в Office для Excel в Интернете.</span><span class="sxs-lookup"><span data-stu-id="faede-105">This article describes how you can use some of the built-in JavaScript objects in Office Scripts for Excel on the web.</span></span>

> [!NOTE]
> <span data-ttu-id="faede-106">Полный список всех встроенных объектов JavaScript можно посмотреть в статье Mozilla [«Стандартные встроенные объекты».](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects)</span><span class="sxs-lookup"><span data-stu-id="faede-106">For a complete list of all built-in JavaScript objects, see Mozilla's [Standard built-in objects](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects) article.</span></span>

## <a name="array"></a><span data-ttu-id="faede-107">Массив</span><span class="sxs-lookup"><span data-stu-id="faede-107">Array</span></span>

<span data-ttu-id="faede-108">Объект [Array](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Array) предоставляет стандартизированный способ работы с массивами в скрипте.</span><span class="sxs-lookup"><span data-stu-id="faede-108">The [Array](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Array) object provides a standardized way to work with arrays in your script.</span></span> <span data-ttu-id="faede-109">Хотя массивы являются стандартными конструкциями JavaScript, они связаны Office скриптами двумя основными способами: диапазонами и коллекциями.</span><span class="sxs-lookup"><span data-stu-id="faede-109">While arrays are standard JavaScript constructs, they relate to Office Scripts in two major ways: ranges and collections.</span></span>

### <a name="work-with-ranges"></a><span data-ttu-id="faede-110">Работа с диапазонами</span><span class="sxs-lookup"><span data-stu-id="faede-110">Work with ranges</span></span>

<span data-ttu-id="faede-111">Диапазоны содержат несколько двумерных массивов, которые непосредственно карт на ячейки в этом диапазоне.</span><span class="sxs-lookup"><span data-stu-id="faede-111">Ranges contain several two-dimensional arrays that directly map to the cells in that range.</span></span> <span data-ttu-id="faede-112">Эти массивы содержат конкретную информацию о каждой ячейке в этом диапазоне.</span><span class="sxs-lookup"><span data-stu-id="faede-112">These arrays contain specific information about each cell in that range.</span></span> <span data-ttu-id="faede-113">Например, `Range.getValues` возвращает все значения в этих ячейках (с рядами и столбцами двумерного отображения массива к строкам и столбцам этого подраздела листа).</span><span class="sxs-lookup"><span data-stu-id="faede-113">For example, `Range.getValues` returns all the values in those cells (with the rows and columns of the two-dimensional array mapping to the rows and columns of that worksheet subsection).</span></span> <span data-ttu-id="faede-114">`Range.getFormulas` и `Range.getNumberFormats` другие часто используемые методы, которые возвращают массивы, такие как `Range.getValues` .</span><span class="sxs-lookup"><span data-stu-id="faede-114">`Range.getFormulas` and `Range.getNumberFormats` are other frequently used methods that return arrays like `Range.getValues`.</span></span>

<span data-ttu-id="faede-115">Следующий скрипт ищет диапазон **A1:D4** для любого формата числа, содержащего "$".</span><span class="sxs-lookup"><span data-stu-id="faede-115">The following script searches the **A1:D4** range for any number format containing a "$".</span></span> <span data-ttu-id="faede-116">Скрипт устанавливает цвет заполнения в этих ячейках на "желтый".</span><span class="sxs-lookup"><span data-stu-id="faede-116">The script sets the fill color in those cells to "yellow".</span></span>

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

### <a name="work-with-collections"></a><span data-ttu-id="faede-117">Работа с коллекциями</span><span class="sxs-lookup"><span data-stu-id="faede-117">Work with collections</span></span>

<span data-ttu-id="faede-118">Многие Excel объекты содержатся в коллекции.</span><span class="sxs-lookup"><span data-stu-id="faede-118">Many Excel objects are contained in a collection.</span></span> <span data-ttu-id="faede-119">Коллекция управляется API Office и выставлена в качестве массива.</span><span class="sxs-lookup"><span data-stu-id="faede-119">The collection is managed by the Office Scripts API and exposed as an array.</span></span> <span data-ttu-id="faede-120">Например, все [фигуры](/javascript/api/office-scripts/excelscript/excelscript.shape) в листе содержатся в `Shape[]` возвращенной `Worksheet.getShapes` методом форме.</span><span class="sxs-lookup"><span data-stu-id="faede-120">For example, all [Shapes](/javascript/api/office-scripts/excelscript/excelscript.shape) in a worksheet are contained in a `Shape[]` that is returned by the `Worksheet.getShapes` method.</span></span> <span data-ttu-id="faede-121">Этот массив можно использовать для чтения значений из коллекции или получить доступ к определенным объектам из методов родительского `get*` объекта.</span><span class="sxs-lookup"><span data-stu-id="faede-121">You can use this array to read values from the collection, or you can access specific objects from the parent object's `get*` methods.</span></span>

> [!NOTE]
> <span data-ttu-id="faede-122">Не добавляйте и не удаляйте объекты из этих массивов коллекции вручную.</span><span class="sxs-lookup"><span data-stu-id="faede-122">Do not manually add or remove objects from these collection arrays.</span></span> <span data-ttu-id="faede-123">Используйте `add` методы на родительских объектах и `delete` методы на объектах типа сбора.</span><span class="sxs-lookup"><span data-stu-id="faede-123">Use the `add` methods on the parent objects and the `delete` methods on the collection-type objects.</span></span> <span data-ttu-id="faede-124">Например, добавьте [таблицу](/javascript/api/office-scripts/excelscript/excelscript.table) [в лист с](/javascript/api/office-scripts/excelscript/excelscript.worksheet) помощью метода и удалите `Worksheet.addTable` `Table` `Table.delete` использование.</span><span class="sxs-lookup"><span data-stu-id="faede-124">For example, add a [Table](/javascript/api/office-scripts/excelscript/excelscript.table) to a [Worksheet](/javascript/api/office-scripts/excelscript/excelscript.worksheet) with the `Worksheet.addTable` method and remove the `Table` using `Table.delete`.</span></span>

<span data-ttu-id="faede-125">Следующий скрипт регистрирует тип каждой формы в текущем листе.</span><span class="sxs-lookup"><span data-stu-id="faede-125">The following script logs the type of every shape in the current worksheet.</span></span>

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

<span data-ttu-id="faede-126">Следующий скрипт удаляет старейшую форму в текущем листе.</span><span class="sxs-lookup"><span data-stu-id="faede-126">The following script deletes the oldest shape in the current worksheet.</span></span>

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

## <a name="date"></a><span data-ttu-id="faede-127">Дата</span><span class="sxs-lookup"><span data-stu-id="faede-127">Date</span></span>

<span data-ttu-id="faede-128">Объект [Дата](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Date) предоставляет стандартизированный способ работы с датами в скрипте.</span><span class="sxs-lookup"><span data-stu-id="faede-128">The [Date](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Date) object provides a standardized way to work with dates in your script.</span></span> <span data-ttu-id="faede-129">`Date.now()` генерирует объект с текущей датой и временем, что полезно при добавлении метки времени к вводу данных скрипта.</span><span class="sxs-lookup"><span data-stu-id="faede-129">`Date.now()` generates an object with the current date and time, which is useful when adding timestamps to your script's data entry.</span></span>

<span data-ttu-id="faede-130">Следующий скрипт добавляет текущую дату в лист.</span><span class="sxs-lookup"><span data-stu-id="faede-130">The following script adds the current date to the worksheet.</span></span> <span data-ttu-id="faede-131">Обратите внимание, что `toLocaleDateString` с помощью Excel, он распознает значение как дату и автоматически изменяет формат номера ячейки.</span><span class="sxs-lookup"><span data-stu-id="faede-131">Note that by using the `toLocaleDateString` method, Excel recognizes the value as a date and changes the number format of the cell automatically.</span></span>

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

<span data-ttu-id="faede-132">В [разделе «Работа с датами»](../resources/samples/excel-samples.md#dates) образцов имеется больше скриптов, связанных с датой.</span><span class="sxs-lookup"><span data-stu-id="faede-132">The [Work with dates](../resources/samples/excel-samples.md#dates) section of the samples has more date-related scripts.</span></span>

## <a name="math"></a><span data-ttu-id="faede-133">математика;</span><span class="sxs-lookup"><span data-stu-id="faede-133">Math</span></span>

<span data-ttu-id="faede-134">Объект [Math предоставляет](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Math) методы и константы для общих математических операций.</span><span class="sxs-lookup"><span data-stu-id="faede-134">The [Math](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Math) object provides methods and constants for common mathematical operations.</span></span> <span data-ttu-id="faede-135">Они обеспечивают много функций, Excel также доступны в других местах, без необходимости использования двигателя расчета рабочей книги.</span><span class="sxs-lookup"><span data-stu-id="faede-135">These provide many functions also available in Excel, without the need to use the workbook's calculation engine.</span></span> <span data-ttu-id="faede-136">Это избавляет ваш скрипт от необходимости запрашивать трудовую книжку, что повышает производительность.</span><span class="sxs-lookup"><span data-stu-id="faede-136">This saves your script from having to query the workbook, which improves performance.</span></span>

<span data-ttu-id="faede-137">Следующий скрипт используется `Math.min` для поиск и регистрации малейшего числа в **диапазоне A1:D4.**</span><span class="sxs-lookup"><span data-stu-id="faede-137">The following script uses `Math.min` to find and log the smallest number in the **A1:D4** range.</span></span> <span data-ttu-id="faede-138">Обратите внимание, что этот пример предполагает, что весь диапазон содержит только числа, а не строки.</span><span class="sxs-lookup"><span data-stu-id="faede-138">Note that this sample assumes the entire range contains only numbers, not strings.</span></span>

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

## <a name="use-of-external-javascript-libraries-is-not-supported"></a><span data-ttu-id="faede-139">Использование внешних библиотек JavaScript не поддерживается</span><span class="sxs-lookup"><span data-stu-id="faede-139">Use of external JavaScript libraries is not supported</span></span>

<span data-ttu-id="faede-140">Office Скрипты не поддерживают использование внешних, сторонних библиотек.</span><span class="sxs-lookup"><span data-stu-id="faede-140">Office Scripts don't support the use of external, third-party libraries.</span></span> <span data-ttu-id="faede-141">Скрипт может использовать только встроенные объекты JavaScript и api Office скриптов.</span><span class="sxs-lookup"><span data-stu-id="faede-141">Your script can only use the built-in JavaScript objects and the Office Scripts APIs.</span></span>

## <a name="see-also"></a><span data-ttu-id="faede-142">См. также</span><span class="sxs-lookup"><span data-stu-id="faede-142">See also</span></span>

- [<span data-ttu-id="faede-143">Стандартные встроенные объекты</span><span class="sxs-lookup"><span data-stu-id="faede-143">Standard built-in objects</span></span>](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects)
- [<span data-ttu-id="faede-144">Office Среда редактора кода скриптов</span><span class="sxs-lookup"><span data-stu-id="faede-144">Office Scripts Code Editor environment</span></span>](../overview/code-editor-environment.md)
