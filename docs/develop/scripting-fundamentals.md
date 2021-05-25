---
title: Основные сведения о сценариях Office в Excel для Интернета
description: Информация об объектной модели и другие основы для изучения перед написанием сценариев Office.
ms.date: 05/24/2021
localization_priority: Priority
ms.openlocfilehash: 629e816ea988d6b8ffe5264c701e3a1eba6c6feb
ms.sourcegitcommit: 90ca8cdf30f2065f63938f6bb6780d024c128467
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 05/25/2021
ms.locfileid: "52639896"
---
# <a name="scripting-fundamentals-for-office-scripts-in-excel-on-the-web"></a>Основные сведения о сценариях Office в Excel для Интернета

Эта статья познакомит вас с техническими аспектами сценариев Office. Вы узнаете, как объекты Excel работают вместе и как редактор кода синхронизируется с книгой.

## <a name="typescript-the-language-of-office-scripts"></a>TypeScript: язык сценариев Office

Сценарии Office написаны на языке [TypeScript](https://www.typescriptlang.org/docs/home.html), который является супермножеством [JavaScript](https://developer.mozilla.org/docs/Web/JavaScript). Если вы знакомы с JavaScript, ваши знания пригодятся, так как большая часть кода одинакова в обоих языках. Перед началом написания кода сценариев Office рекомендуется получить опыт программирования на начальном уровне. Следующие ресурсы помогут вам понять код сценариев Office.

[!INCLUDE [Recommended coding resources](../includes/coding-basics-references.md)]

## <a name="main-function-the-scripts-starting-point"></a>Функция `main`: начальная точка сценария

Каждый сценарий должен содержать функцию `main` с типом `ExcelScript.Workbook` в качестве первого параметра. При выполнении функции приложение Excel вызывает функцию `main`, предоставляя книгу в качестве ее первого параметра. Параметр `ExcelScript.Workbook` всегда должен быть первым параметром.

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Your code goes here
}
```

Код внутри `main` функции запускается при запуске скрипта. `main` может вызывать другие функции в вашем скрипте, но код, который не содержится в функции, не будет работать. Сценарии не могут вызывать другие сценарии Office.

[Power Automate](https://flow.microsoft.com) позволяет подключать сценарии в потоках. Данные передаются между сценариями и потоком через параметры и возвращаемые результаты метода `main`. Способ интеграции сценариев Office с Power Automate подробно описан в статье [Запуск сценариев Office с помощью Power Automate](power-automate-integration.md).

## <a name="object-model-overview"></a>Обзор объектной модели

Чтобы написать сценарий, необходимо знать, как устроены API сценариев Office. Компоненты книги определенным образом взаимосвязаны друг с другом. Эти взаимосвязи во многом схожи с пользовательским интерфейсом Excel.

- **Рабочая книга** содержит одну или несколько **рабочих листов**.
- **Рабочий лист** предоставляет доступ к ячейкам через объекты **Range**.
- **Range** представляет группу смежных клеток.
- **Диапазоны** используются для создания и размещения **таблиц**, **диаграмм**, **фигур** и других объектов визуализации данных или организации.
- **Рабочий лист** содержит коллекции тех объектов данных, которые присутствуют на отдельном листе.
- **Рабочие книги** содержат коллекции некоторых из этих объектов данных (таких как **таблицы**) для всей **рабочей книги**.

## <a name="workbook"></a>Книга

Для каждого сценария предоставляется объект `workbook` типа `Workbook`, он предоставляется функцией `main`. Это объект верхнего уровня, через который сценарий взаимодействует с книгой Excel.

Следующий сценарий получает активный лист из книги и записывает его имя.

```TypeScript
function main(workbook: ExcelScript.Workbook) {
    // Get the active worksheet.
    let sheet = workbook.getActiveWorksheet();

    // Display the current worksheet's name.
    console.log(sheet.getName());
}
```

## <a name="ranges"></a>Диапазоны

Диапазон - это группа непрерывных ячеек в рабочей книге. В сценариях обычно используется нотация в стиле A1 (например, **B3** для отдельной ячейки в столбце **B** и строке **3** или **C2:F4** для ячеек из столбцов с **C** по **F** и строк с **2** по **4**) для определения диапазонов.

У диапазонов три основных свойства: значения, формулы и формат. Эти свойства получают или устанавливают значения ячеек, формулы для вычисления и визуальное форматирование ячеек. Для доступа к ним используются `getValues`, `getFormulas` и `getFormat`. Значения и формулы можно изменять с помощью `setValues` и `setFormulas`, а формат является объектом `RangeFormat`, который состоит из нескольких меньших объектов, задаваемых по отдельности.

Диапазоны используют двухмерные массивы для управления информацией. Дополнительные сведения об обработке массивов в инфраструктуре сценариев Office см. в статье [Работа с диапазонами](javascript-objects.md#work-with-ranges).

### <a name="range-sample"></a>Образец диапазона

В следующем примере показано, как создавать записи продаж. В этом сценарии используются объекты `Range` для установки значений, формул и частей формата.

```TypeScript
function main(workbook: ExcelScript.Workbook) {
    // Get the active worksheet.
    let sheet = workbook.getActiveWorksheet();

    // Create the headers and format them to stand out.
    let headers = [["Product", "Quantity", "Unit Price", "Totals"]];
    let headerRange = sheet.getRange("B2:E2");
    headerRange.setValues(headers);
    headerRange.getFormat().getFill().setColor("#4472C4");
    headerRange.getFormat().getFont().setColor("white");

    // Create the product data rows.
    let productData = [
        ["Almonds", 6, 7.5],
        ["Coffee", 20, 34.5],
        ["Chocolate", 10, 9.54],
    ];
    let dataRange = sheet.getRange("B3:D5");
    dataRange.setValues(productData);

    // Create the formulas to total the amounts sold.
    let totalFormulas = [
        ["=C3 * D3"],
        ["=C4 * D4"],
        ["=C5 * D5"],
        ["=SUM(E3:E5)"],
    ];
    let totalRange = sheet.getRange("E3:E6");
    totalRange.setFormulas(totalFormulas);
    totalRange.getFormat().getFont().setBold(true);

    // Display the totals as US dollar amounts.
    totalRange.setNumberFormat("$0.00");
}
```

Выполнение этого скрипта создает следующие данные в текущей рабочей таблице:

:::image type="content" source="../images/range-sample.png" alt-text="Лист с записями о продажах, содержащий строки значений, столбец формулы и отформатированные заголовки":::

### <a name="the-types-of-range-values"></a>Типы значений диапазона

Каждая ячейка содержит значение. Это значение является базовым значением, введенным в ячейку, которое может отличаться от текста, отображаемого в Excel. Например, в ячейке может отображаться дата 02.05.2021, но действительное значение — 44318. Это отображаемое значение можно изменить с использованием числового формата, но действительное значение и тип в ячейке изменяются только при настройке нового значения.

При использовании значения ячейки важно сообщить TypeScript, какое значение вы ожидаете получить из ячейки или диапазона. Ячейка содержит один из следующих типов: `string`, `number`или `boolean`. Чтобы сценарий обрабатывал возвращенные значения как один из этих типов, необходимо объявить тип.

Следующий сценарий получает среднюю цену из таблицы в предыдущем примере. Обратите внимание на код `priceRange.getValues() as number[][]`. Это [утверждает](https://www.typescriptlang.org/docs/handbook/2/everyday-types.html#type-assertions) `number[][]` в качестве типа значений диапазона. После этого все значения в этом массиве могут рассматриваться как числа в сценарии.

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Get the active worksheet.
  let sheet = workbook.getActiveWorksheet();

  // Get the "Unit Price" column. 
  // The result of calling getValues is declared to be a number[][] so that we can perform arithmetic operations.
  let priceRange = sheet.getRange("D3:D5");
  let prices = priceRange.getValues() as number[][];

  // Get the average price.
  let totalPrices = 0;
  prices.forEach((price) => totalPrices += price[0]);
  let averagePrice = totalPrices / prices.length;
  console.log(averagePrice);
}
```

## <a name="charts-tables-and-other-data-objects"></a>Диаграммы, таблицы и другие объекты данных

Скрипты могут создавать и управлять структурами данных и визуализациями в Excel. Таблицы и диаграммы являются двумя наиболее часто используемыми объектами, но API поддерживают сводные таблицы, фигуры, изображения и многое другое. Они сохраняются в коллекциях, которые рассматриваются далее в этой статье.

### <a name="create-a-table"></a>Создание таблицы

Создайте таблицы с помощью диапазонов данных. Форматирование и элементы управления таблицами (например, фильтры) автоматически применяются к диапазону.

Следующий скрипт создает таблицу с использованием диапазонов из предыдущего примера.

```TypeScript
function main(workbook: ExcelScript.Workbook) {
    // Get the active worksheet.
    let sheet = workbook.getActiveWorksheet();

    // Add a table that has headers using the data from B2:E5.
    sheet.addTable("B2:E5", true);
}
```

Выполнение этого сценария на листе с предыдущими данными создает следующую таблицу:

:::image type="content" source="../images/table-sample.png" alt-text="Лист, содержащий таблицу, созданную из предыдущей записи о продажах":::

### <a name="create-a-chart"></a>Создание диаграммы

Создайте диаграммы для визуализации данных в диапазоне. Сценарии позволяют создавать десятки разновидностей диаграмм, каждая из которых может быть настроена в соответствии с вашими потребностями.

Следующий скрипт создает простую столбчатую диаграмму для трех элементов и размещает ее на 100 пикселей ниже верхней части листа.

```TypeScript
function main(workbook: ExcelScript.Workbook) {
    // Get the active worksheet.
    let sheet = workbook.getActiveWorksheet();

    // Create a column chart using the data from B3:C5.
    let chart = sheet.addChart(
        ExcelScript.ChartType.columnStacked,
        sheet.getRange("B3:C5")
    );

    // Set the margin of the chart to be 100 pixels from the top of the screen.
    chart.setTop(100);
}
```

Запуск этого скрипта на листе с предыдущей таблицей создает следующую диаграмму:

:::image type="content" source="../images/chart-sample.png" alt-text="Гистограмма, показывающая количество для трех элементов из предыдущей записи о продажах":::

## <a name="collections"></a>Коллекции

Если объект Excel содержит коллекцию из одного или нескольких объектов одного типа, он сохраняет их в массиве. Например, объект `Workbook` содержит `Worksheet[]`. Доступ к этому массиву обеспечивается методом `Workbook.getWorksheets()`. Множественные методы `get`, например `Worksheet.getCharts()`, возвращают всю коллекцию объектов в качестве массива. Вы увидите этот шаблон во всех API сценариев Office: объект `Worksheet` использует метод `getTables()`, возвращающий `Table[]`, объект `Table` использует метод `getColumns()`, возвращающий `TableColumn[]`, и т. д.

Возвращаемый массив является обычным массивом, поэтому все обычные операции массивов доступны для вашего сценария. Также можно получить доступ к отдельным объектам внутри коллекции с помощью значения индекса массива. Например, `workbook.getTables()[0]` возвращает первую таблицу в коллекции. Дополнительные сведения об использовании встроенных функций массива в структуре сценариев Office см. в статье [Работа с коллекциями](javascript-objects.md#work-with-collections). 

Отдельные объекты также доступны из коллекции с помощью метода `get`. Одиночные методы `get`, например `Worksheet.getTable(name)`, возвращают один объект и требуют идентификатор или имя конкретного объекта. Этот идентификатор или имя обычно задается сценарием или с помощью пользовательского интерфейса Excel.

Следующий сценарий возвращает все таблицы в книге. При этом отображаются заголовки, видны кнопки фильтров, а для таблицы устанавливается стиль "TableStyleLight1".

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Get the table collection.
  let tables = workbook.getTables();

  // Set the table formatting properties for every table.
  tables.forEach(table => {
    table.setShowHeaders(true);
    table.setShowFilterButton(true);
    table.setPredefinedTableStyle("TableStyleLight1");
  })
}
```

## <a name="add-excel-objects-with-a-script"></a>Добавление объектов Excel с помощью сценария

Можно программным образом добавлять объекты документов, например таблицы или диаграммы, путем вызова соответствующего метода `add`, доступного для родительского объекта.

> [!IMPORTANT]
> Не следует вручную добавлять объекты в массивы коллекций. Используйте методы `add` для родительских объектов. Например, можно добавить `Table` к `Worksheet` методом `Worksheet.addTable`.

Следующий сценарий создает таблицу в Excel на первом листе книги. Обратите внимание, что метод `addTable` возвращает созданную таблицу.

```TypeScript
function main(workbook: ExcelScript.Workbook) {
    // Get the first worksheet.
    let sheet = workbook.getWorksheets()[0];

    // Add a table that uses the data in A1:G10.
    let table = sheet.addTable(
      "A1:G10",
       true /* True because the table has headers. */
    );
    
    // Give the table a name for easy reference in other scripts.
    table.setName("MyTable");
}
```

> [!TIP]
> Большинство объектов Excel используют метод `setName`. Это позволяет легко получить доступ к объектам Excel позже в сценарии или в других сценариях для той же книги.

### <a name="verify-an-object-exists-in-the-collection"></a>Проверка существования объекта в коллекции

Перед продолжением сценариям часто требуется проверить, существует ли таблица или похожий объект. Используйте имена, заданные сценариями или с помощью пользовательского интерфейса Excel, чтобы определить необходимые объекты и действовать соответствующим образом. Методы `get` возвращают `undefined`, когда запрашиваемый объект отсутствует в коллекции.

Следующий сценарий запрашивает таблицу MyTable и использует оператор `if...else`, чтобы проверить, найдена ли таблица.

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Get the table named "MyTable".
  let myTable = workbook.getTable("MyTable");

  // If the table is in the workbook, myTable will have a value.
  // Otherwise, the variable will be undefined and go to the else clause.
  if (myTable) {
    let worksheetName = myTable.getWorksheet().getName();
    console.log(`MyTable is on the ${worksheetName} worksheet`);
  } else {
    console.log(`MyTable is not in the workbook.`);
  }
}
```

Распространенный шаблон в сценариях Office — воссоздание таблицы, диаграммы или другого объекта при каждом запуске сценария. Если старые данные не нужны, рекомендуется удалить старый объект перед созданием нового. Это позволяет избежать конфликтов имен или других различий, которые могли быть добавлены другими пользователями.

Следующий сценарий удаляет таблицу MyTable, если она существует, а затем добавляет новую таблицу с таким же именем.

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Get the table named "MyTable" from the first worksheet.
  let sheet = workbook.getWorksheets()[0];
  let tableName = "MyTable";
  let oldTable = sheet.getTable(tableName);

  // If the table exists, remove it.
  if (oldTable) {
    oldTable.delete();
  }

  // Add a new table with the same name.
  let newTable = sheet.addTable("A1:G10", true);
  newTable.setName(tableName);
}
```

## <a name="remove-excel-objects-with-a-script"></a>Удаление объектов Excel с помощью сценария

Чтобы удалить объект, вызовите метод `delete` этого объекта.

> [!NOTE]
> Как и в случае добавления объектов, не следует вручную удалять объекты из массивов коллекций. Используйте методы `delete` для объектов типа коллекции. Например, для удаления `Table` из `Worksheet` используйте `Table.delete`.

Следующий сценарий удаляет первый лист в книге.

```TypeScript
function main(workbook: ExcelScript.Workbook) {
    // Get first worksheet.
    let sheet = workbook.getWorksheets()[0];

    // Remove that worksheet from the workbook.
    sheet.delete();
}
```

## <a name="further-reading-on-the-object-model"></a>Дальнейшее чтение по объектной модели

[Справочная документация по API сценариев Office](/javascript/api/office-scripts/overview) представляет собой полный список объектов, используемых в сценариях Office. Там вы можете использовать оглавление, чтобы перейти к любому классу, о котором вы хотите узнать больше. Ниже приведены несколько часто просматриваемых страниц.

- [Chart](/javascript/api/office-scripts/excelscript/excelscript.chart)
- [Comment](/javascript/api/office-scripts/excelscript/excelscript.comment)
- [PivotTable](/javascript/api/office-scripts/excelscript/excelscript.pivottable)
- [Range](/javascript/api/office-scripts/excelscript/excelscript.range)
- [RangeFormat](/javascript/api/office-scripts/excelscript/excelscript.rangeformat)
- [Shape](/javascript/api/office-scripts/excelscript/excelscript.shape)
- [Table](/javascript/api/office-scripts/excelscript/excelscript.table)
- [Workbook](/javascript/api/office-scripts/excelscript/excelscript.workbook)
- [Worksheet](/javascript/api/office-scripts/excelscript/excelscript.worksheet)

## <a name="see-also"></a>См. также

- [Запись, редактирование и создание сценариев Office в Excel в Интернете](../tutorials/excel-tutorial.md)
- [Чтение данных рабочей книги с помощью сценариев Office в Excel в Интернете](../tutorials/excel-read-tutorial.md)
- [Справочник API для сценариев Office](/javascript/api/office-scripts/overview)
- [Использование встроенных объектов JavaScript в сценариях Office](javascript-objects.md)
- [Рекомендации по сценариям Office](best-practices.md)
