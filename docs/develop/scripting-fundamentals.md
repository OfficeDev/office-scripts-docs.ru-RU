---
title: Основы сценариев для сценариев Office в Excel в Интернете
description: Информация об объектной модели и другие основы для изучения перед написанием сценариев Office.
ms.date: 07/08/2020
localization_priority: Priority
ms.openlocfilehash: 4609dce617d18f7a7c9d2868a24792c51fb048c6
ms.sourcegitcommit: 5d24e77df70aa2c1c982275d53213c2a9323ff86
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 04/02/2021
ms.locfileid: "51570236"
---
# <a name="scripting-fundamentals-for-office-scripts-in-excel-on-the-web-preview"></a>Основы сценариев для сценариев Office в Excel в Интернете (предварительная версия)

Эта статья познакомит вас с техническими аспектами сценариев Office. Вы узнаете, как объекты Excel работают вместе и как редактор кода синхронизируется с книгой.

[!INCLUDE [Preview note](../includes/preview-note.md)]

## <a name="main-function"></a>Функция `main`

Каждый сценарий Office должен содержать функцию `main` с типом `ExcelScript.Workbook` в качестве первого параметра. При выполнении `main` приложение Excel вызывает эту функцию, предоставляя книгу в качестве ее первого параметра. Поэтому важно не изменять базовую подпись функции `main` после записи сценария или создания нового сценария в редакторе кода.

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Your code goes here
}
```

Код внутри `main` функции запускается при запуске скрипта. `main` может вызывать другие функции в вашем скрипте, но код, который не содержится в функции, не будет работать.

> [!CAUTION]
> Если ваша функция `main` выглядит как `async function main(context: Excel.RequestContext)`, то сценарий использует устаревшую асинхронную модель API. Дополнительные сведения, включая сведения о преобразовании устаревших сценариев в текущую модель API, см. в статье [Использование асинхронных API сценариев Office для поддержки устаревших сценариев](excel-async-model.md).

## <a name="object-model"></a>Объектная модель

Чтобы написать сценарий, необходимо знать, как устроены API Office Script. Компоненты книги определенным образом взаимосвязаны друг с другом. Эти взаимосвязи во многом схожи с пользовательским интерфейсом Excel.

- **Рабочая книга** содержит одну или несколько **рабочих листов**.
- **Рабочий лист** предоставляет доступ к ячейкам через объекты **Range**.
- **Range** представляет группу смежных клеток.
- **Диапазоны** используются для создания и размещения **таблиц**, **диаграмм**, **фигур** и других объектов визуализации данных или организации.
- **Рабочий лист** содержит коллекции тех объектов данных, которые присутствуют на отдельном листе.
- **Рабочие книги** содержат коллекции некоторых из этих объектов данных (таких как **таблицы**) для всей **рабочей книги**.

### <a name="workbook"></a>Книга

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

### <a name="ranges"></a>Диапазоны

Диапазон - это группа непрерывных ячеек в рабочей книге. В сценариях обычно используется нотация в стиле A1 (например, **B3** для отдельной ячейки в столбце **B** и строке **3** или **C2:F4** для ячеек из столбцов с **C** по **F** и строк с **2** по **4**) для определения диапазонов.

У диапазонов три основных свойства: значения, формулы и формат. Эти свойства получают или устанавливают значения ячеек, формулы для вычисления и визуальное форматирование ячеек. Для доступа к ним используются `getValues`, `getFormulas` и `getFormat`. Значения и формулы можно изменять с помощью `setValues` и `setFormulas`, а формат является объектом `RangeFormat`, который состоит из нескольких меньших объектов, задаваемых по отдельности.

Диапазоны используют двухмерные массивы для управления информацией. Дополнительные сведения об обработке этих массивов на платформе сценариев Office см. в разделе ["Работа с диапазонами" статьи "Использование встроенных объектов JavaScript в сценариях Office"](javascript-objects.md#working-with-ranges).

#### <a name="range-sample"></a>Образец диапазона

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
        ["Chocolate", 10, 9.56],
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

![Запись о продажах, показывающая строки значений, столбец формулы и отформатированные заголовки.](../images/range-sample.png)

### <a name="charts-tables-and-other-data-objects"></a>Диаграммы, таблицы и другие объекты данных

Скрипты могут создавать и управлять структурами данных и визуализациями в Excel. Таблицы и диаграммы являются двумя наиболее часто используемыми объектами, но API поддерживают сводные таблицы, фигуры, изображения и многое другое. Они сохраняются в коллекциях, которые рассматриваются далее в этой статье.

#### <a name="creating-a-table"></a>Создание таблицы

Создавайте таблицы, используя заполненные данными диапазоны. Элементы управления форматированием и таблицами (например, фильтры) автоматически применяются к диапазону.

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

![Таблица сделана из предыдущего рекорда продаж.](../images/table-sample.png)

#### <a name="creating-a-chart"></a>Создание диаграммы

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

![Столбчатая диаграмма, показывающая количество трех предметов из предыдущей записи о продажах.](../images/chart-sample.png)

### <a name="collections-and-other-object-relations"></a>Коллекции и другие отношения объектов

Доступ к любому дочернему объекту осуществляется через его родительский объект. Например, можно прочесть `Worksheets` из объекта `Workbook`. Будет доступен связанный метод `get` родительского класса (например, `Workbook.getWorksheets()` или `Workbook.getWorksheet(name)`). Одиночные методы `get` возвращают один объект, им требуется идентификатор или имя конкретного объекта (например, имя листа). Множественные методы `get` возвращают всю коллекцию объектов в качестве массива. Если коллекция пуста, возвращается пустой массив (`[]`).

После получения коллекции можно использовать обычные операции с массивами, такие как получение его `length` или использование циклов `for`, `for..of`, `while` для итерации. Также можно использовать методы массивов TypeScript, такие как `map`, `forEach`. Также можно получить доступ к отдельным объектам внутри коллекции с помощью значения индекса массива. Например, `workbook.getTables()[0]` возвращает первую таблицу в коллекции. Дополнительные сведения об использовании встроенной функциональности массивов платформы сценариев Office см. в разделе ["Работа с коллекциями" статьи "Использование встроенных объектов JavaScript в сценариях Office"](javascript-objects.md#working-with-collections).

Следующий сценарий возвращает все таблицы в книге. При этом отображаются заголовки, видны кнопки фильтров, а для таблицы устанавливается стиль "TableStyleLight1".

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  /* Get table collection */
  const tables = workbook.getTables();
  /* Set table formatting properties */
  tables.forEach(table => {
    table.setShowHeaders(true);
    table.setShowFilterButton(true);
    table.setPredefinedTableStyle("TableStyleLight1");
  })
}
```

#### <a name="adding-excel-objects-with-a-script"></a>Добавление объектов Excel с помощью сценария

Можно программным образом добавлять объекты документов, например таблицы или диаграммы, путем вызова соответствующего метода `add`, доступного для родительского объекта.

> [!NOTE]
> Не следует вручную добавлять объекты в массивы коллекций. Используйте методы `add` для родительских объектов. Например, можно добавить `Table` к `Worksheet` методом `Worksheet.addTable`.

Следующий сценарий создает таблицу в Excel на первом листе книги. Обратите внимание, что метод `addTable` возвращает созданную таблицу.

```TypeScript
function main(workbook: ExcelScript.Workbook) {
    // Get the first worksheet.
    let sheet = workbook.getWorksheets()[0];

    // Add a table that uses the data in C3:G10.
    let table = sheet.addTable(
      "C3:G10",
       true /* True because the table has headers. */
    );
}
```

## <a name="removing-excel-objects-with-a-script"></a>Удаление объектов Excel с помощью сценария

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

### <a name="further-reading-on-the-object-model"></a>Дальнейшее чтение по объектной модели

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
