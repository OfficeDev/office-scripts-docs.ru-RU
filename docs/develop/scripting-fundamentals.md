---
title: Основные сведения о сценариях Office в Excel для Интернета
description: Информация об объектной модели и другие основы для изучения перед написанием сценариев Office.
ms.date: 05/24/2021
localization_priority: Priority
ms.openlocfilehash: 9c3c10e283e40f1e719e73106bcdacfcff44dbc9
ms.sourcegitcommit: 4693c8f79428ec74695328275703af0ba1bfea8f
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 06/23/2021
ms.locfileid: "53074510"
---
# <a name="scripting-fundamentals-for-office-scripts-in-excel-on-the-web"></a><span data-ttu-id="1e012-103">Основные сведения о сценариях Office в Excel для Интернета</span><span class="sxs-lookup"><span data-stu-id="1e012-103">Scripting fundamentals for Office Scripts in Excel on the web</span></span>

<span data-ttu-id="1e012-104">Эта статья познакомит вас с техническими аспектами сценариев Office.</span><span class="sxs-lookup"><span data-stu-id="1e012-104">This article will introduce you to the technical aspects of Office Scripts.</span></span> <span data-ttu-id="1e012-105">Вы узнаете, как объекты Excel работают вместе и как редактор кода синхронизируется с книгой.</span><span class="sxs-lookup"><span data-stu-id="1e012-105">You'll learn how the Excel objects work together and how the Code Editor synchronizes with a workbook.</span></span>

## <a name="typescript-the-language-of-office-scripts"></a><span data-ttu-id="1e012-106">TypeScript: язык сценариев Office</span><span class="sxs-lookup"><span data-stu-id="1e012-106">TypeScript: The language of Office Scripts</span></span>

<span data-ttu-id="1e012-107">Сценарии Office написаны на языке [TypeScript](https://www.typescriptlang.org/docs/home.html), который является супермножеством [JavaScript](https://developer.mozilla.org/docs/Web/JavaScript).</span><span class="sxs-lookup"><span data-stu-id="1e012-107">Office Scripts are written in [TypeScript](https://www.typescriptlang.org/docs/home.html), which is a superset of [JavaScript](https://developer.mozilla.org/docs/Web/JavaScript).</span></span> <span data-ttu-id="1e012-108">Если вы знакомы с JavaScript, ваши знания пригодятся, так как большая часть кода одинакова в обоих языках.</span><span class="sxs-lookup"><span data-stu-id="1e012-108">If you're familiar with JavaScript, your knowledge will carry over because much of the code is the same in both languages.</span></span> <span data-ttu-id="1e012-109">Перед началом написания кода сценариев Office рекомендуется получить опыт программирования на начальном уровне.</span><span class="sxs-lookup"><span data-stu-id="1e012-109">We recommend you have some beginner-level programming knowledge before starting your Office Scripts coding journey.</span></span> <span data-ttu-id="1e012-110">Следующие ресурсы помогут вам понять код сценариев Office.</span><span class="sxs-lookup"><span data-stu-id="1e012-110">The following resources can help you understand the coding side of Office Scripts.</span></span>

[!INCLUDE [Recommended coding resources](../includes/coding-basics-references.md)]

## <a name="main-function-the-scripts-starting-point"></a><span data-ttu-id="1e012-111">Функция `main`: начальная точка сценария</span><span class="sxs-lookup"><span data-stu-id="1e012-111">`main` function: The script's starting point</span></span>

<span data-ttu-id="1e012-112">Каждый сценарий должен содержать функцию `main` с типом `ExcelScript.Workbook` в качестве первого параметра.</span><span class="sxs-lookup"><span data-stu-id="1e012-112">Each script must contain a `main` function with the `ExcelScript.Workbook` type as its first parameter.</span></span> <span data-ttu-id="1e012-113">При выполнении функции приложение Excel вызывает функцию `main`, предоставляя книгу в качестве ее первого параметра.</span><span class="sxs-lookup"><span data-stu-id="1e012-113">When the function runs, the Excel application invokes the `main` function by providing the workbook as its first parameter.</span></span> <span data-ttu-id="1e012-114">Параметр `ExcelScript.Workbook` всегда должен быть первым параметром.</span><span class="sxs-lookup"><span data-stu-id="1e012-114">An `ExcelScript.Workbook` should always be the first parameter.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Your code goes here
}
```

<span data-ttu-id="1e012-115">Код внутри `main` функции запускается при запуске скрипта.</span><span class="sxs-lookup"><span data-stu-id="1e012-115">The code inside the `main` function runs when the script is run.</span></span> <span data-ttu-id="1e012-116">`main` может вызывать другие функции в вашем скрипте, но код, который не содержится в функции, не будет работать.</span><span class="sxs-lookup"><span data-stu-id="1e012-116">`main` can call other functions in your script, but code that's not contained in a function will not run.</span></span> <span data-ttu-id="1e012-117">Сценарии не могут вызывать другие сценарии Office.</span><span class="sxs-lookup"><span data-stu-id="1e012-117">Scripts cannot invoke or call other Office Scripts.</span></span>

<span data-ttu-id="1e012-118">[Power Automate](https://flow.microsoft.com) позволяет подключать сценарии в потоках.</span><span class="sxs-lookup"><span data-stu-id="1e012-118">[Power Automate](https://flow.microsoft.com) allows you to connect scripts in flows.</span></span> <span data-ttu-id="1e012-119">Данные передаются между сценариями и потоком через параметры и возвращаемые результаты метода `main`.</span><span class="sxs-lookup"><span data-stu-id="1e012-119">Data is passed between the scripts and the flow through the parameters and returns of the`main` method.</span></span> <span data-ttu-id="1e012-120">Способ интеграции сценариев Office с Power Automate подробно описан в статье [Запуск сценариев Office с помощью Power Automate](power-automate-integration.md).</span><span class="sxs-lookup"><span data-stu-id="1e012-120">How to integrate Office Scripts with Power Automate is covered in detail in [Run Office Scripts with Power Automate](power-automate-integration.md).</span></span>

## <a name="object-model-overview"></a><span data-ttu-id="1e012-121">Обзор объектной модели</span><span class="sxs-lookup"><span data-stu-id="1e012-121">Object model overview</span></span>

<span data-ttu-id="1e012-122">Чтобы написать сценарий, необходимо знать, как устроены API сценариев Office.</span><span class="sxs-lookup"><span data-stu-id="1e012-122">To write a script, you need to understand how the Office Scripts APIs fit together.</span></span> <span data-ttu-id="1e012-123">Компоненты книги определенным образом взаимосвязаны друг с другом.</span><span class="sxs-lookup"><span data-stu-id="1e012-123">The components of a workbook have specific relations to one another.</span></span> <span data-ttu-id="1e012-124">Эти взаимосвязи во многом схожи с пользовательским интерфейсом Excel.</span><span class="sxs-lookup"><span data-stu-id="1e012-124">In many ways, these relations match those of the Excel UI.</span></span>

- <span data-ttu-id="1e012-125">**Рабочая книга** содержит одну или несколько **рабочих листов**.</span><span class="sxs-lookup"><span data-stu-id="1e012-125">A **Workbook** contains one or more **Worksheets**.</span></span>
- <span data-ttu-id="1e012-126">**Рабочий лист** предоставляет доступ к ячейкам через объекты **Range**.</span><span class="sxs-lookup"><span data-stu-id="1e012-126">A **Worksheet** gives access to cells through **Range** objects.</span></span>
- <span data-ttu-id="1e012-127">**Range** представляет группу смежных клеток.</span><span class="sxs-lookup"><span data-stu-id="1e012-127">A **Range** represents a group of contiguous cells.</span></span>
- <span data-ttu-id="1e012-128">**Диапазоны** используются для создания и размещения **таблиц**, **диаграмм**, **фигур** и других объектов визуализации данных или организации.</span><span class="sxs-lookup"><span data-stu-id="1e012-128">**Ranges** are used to create and place **Tables**, **Charts**, **Shapes**, and other data visualization or organization objects.</span></span>
- <span data-ttu-id="1e012-129">**Рабочий лист** содержит коллекции тех объектов данных, которые присутствуют на отдельном листе.</span><span class="sxs-lookup"><span data-stu-id="1e012-129">A **Worksheet** contains collections of those data objects that are present in the individual sheet.</span></span>
- <span data-ttu-id="1e012-130">**Рабочие книги** содержат коллекции некоторых из этих объектов данных (таких как **таблицы**) для всей **рабочей книги**.</span><span class="sxs-lookup"><span data-stu-id="1e012-130">**Workbooks** contain collections of some of those data objects (such as **Tables**) for the entire **Workbook**.</span></span>

## <a name="workbook"></a><span data-ttu-id="1e012-131">Книга</span><span class="sxs-lookup"><span data-stu-id="1e012-131">Workbook</span></span>

<span data-ttu-id="1e012-132">Для каждого сценария предоставляется объект `workbook` типа `Workbook`, он предоставляется функцией `main`.</span><span class="sxs-lookup"><span data-stu-id="1e012-132">Every script is provided a `workbook` object of type `Workbook` by the `main` function.</span></span> <span data-ttu-id="1e012-133">Это объект верхнего уровня, через который сценарий взаимодействует с книгой Excel.</span><span class="sxs-lookup"><span data-stu-id="1e012-133">This represents the top level object through which your script interacts with the Excel workbook.</span></span>

<span data-ttu-id="1e012-134">Следующий сценарий получает активный лист из книги и записывает его имя.</span><span class="sxs-lookup"><span data-stu-id="1e012-134">The following script gets the active worksheet from the workbook and logs its name.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
    // Get the active worksheet.
    let sheet = workbook.getActiveWorksheet();

    // Display the current worksheet's name.
    console.log(sheet.getName());
}
```

## <a name="ranges"></a><span data-ttu-id="1e012-135">Диапазоны</span><span class="sxs-lookup"><span data-stu-id="1e012-135">Ranges</span></span>

<span data-ttu-id="1e012-136">Диапазон - это группа непрерывных ячеек в рабочей книге.</span><span class="sxs-lookup"><span data-stu-id="1e012-136">A range is a group of contiguous cells in the workbook.</span></span> <span data-ttu-id="1e012-137">В сценариях обычно используется нотация в стиле A1 (например, **B3** для отдельной ячейки в столбце **B** и строке **3** или **C2:F4** для ячеек из столбцов с **C** по **F** и строк с **2** по **4**) для определения диапазонов.</span><span class="sxs-lookup"><span data-stu-id="1e012-137">Scripts typically use A1-style notation (e.g., **B3** for the single cell in column **B** and row **3** or **C2:F4** for the cells from columns **C** through **F** and rows **2** through **4**) to define ranges.</span></span>

<span data-ttu-id="1e012-138">У диапазонов три основных свойства: значения, формулы и формат.</span><span class="sxs-lookup"><span data-stu-id="1e012-138">Ranges have three core properties: values, formulas, and format.</span></span> <span data-ttu-id="1e012-139">Эти свойства получают или устанавливают значения ячеек, формулы для вычисления и визуальное форматирование ячеек.</span><span class="sxs-lookup"><span data-stu-id="1e012-139">These properties get or set the cell values, formulas to be evaluated, and the visual formatting of the cells.</span></span> <span data-ttu-id="1e012-140">Для доступа к ним используются `getValues`, `getFormulas` и `getFormat`.</span><span class="sxs-lookup"><span data-stu-id="1e012-140">They are accessed through `getValues`, `getFormulas`, and `getFormat`.</span></span> <span data-ttu-id="1e012-141">Значения и формулы можно изменять с помощью `setValues` и `setFormulas`, а формат является объектом `RangeFormat`, который состоит из нескольких меньших объектов, задаваемых по отдельности.</span><span class="sxs-lookup"><span data-stu-id="1e012-141">Values and formulas can be changed with `setValues` and `setFormulas`, while the format is a `RangeFormat` object comprised of several smaller objects that are individually set.</span></span>

<span data-ttu-id="1e012-142">Диапазоны используют двухмерные массивы для управления информацией.</span><span class="sxs-lookup"><span data-stu-id="1e012-142">Ranges use two-dimensional arrays to manage information.</span></span> <span data-ttu-id="1e012-143">Дополнительные сведения об обработке массивов в инфраструктуре сценариев Office см. в статье [Работа с диапазонами](javascript-objects.md#work-with-ranges).</span><span class="sxs-lookup"><span data-stu-id="1e012-143">For more information on handling arrays in the Office Scripts framework, see [Work with ranges](javascript-objects.md#work-with-ranges).</span></span>

### <a name="range-sample"></a><span data-ttu-id="1e012-144">Образец диапазона</span><span class="sxs-lookup"><span data-stu-id="1e012-144">Range sample</span></span>

<span data-ttu-id="1e012-145">В следующем примере показано, как создавать записи продаж.</span><span class="sxs-lookup"><span data-stu-id="1e012-145">The following sample shows how to create sales records.</span></span> <span data-ttu-id="1e012-146">В этом сценарии используются объекты `Range` для установки значений, формул и частей формата.</span><span class="sxs-lookup"><span data-stu-id="1e012-146">This script uses `Range` objects to set the values, formulas, and parts of the format.</span></span>

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

<span data-ttu-id="1e012-147">Выполнение этого скрипта создает следующие данные в текущей рабочей таблице:</span><span class="sxs-lookup"><span data-stu-id="1e012-147">Running this script creates the following data in the current worksheet:</span></span>

:::image type="content" source="../images/range-sample.png" alt-text="Лист с записями о продажах, содержащий строки значений, столбец формулы и отформатированные заголовки.":::

### <a name="the-types-of-range-values"></a><span data-ttu-id="1e012-149">Типы значений диапазона</span><span class="sxs-lookup"><span data-stu-id="1e012-149">The types of Range values</span></span>

<span data-ttu-id="1e012-150">Каждая ячейка содержит значение.</span><span class="sxs-lookup"><span data-stu-id="1e012-150">Each cell has value.</span></span> <span data-ttu-id="1e012-151">Это значение является базовым значением, введенным в ячейку, которое может отличаться от текста, отображаемого в Excel.</span><span class="sxs-lookup"><span data-stu-id="1e012-151">This value is the underlying value entered into the cell, which may be different from the text displayed in Excel.</span></span> <span data-ttu-id="1e012-152">Например, в ячейке может отображаться дата 02.05.2021, но действительное значение — 44318.</span><span class="sxs-lookup"><span data-stu-id="1e012-152">For example, you might see "5/2/2021" displayed in the cell as a date, but the actual value is 44318.</span></span> <span data-ttu-id="1e012-153">Это отображаемое значение можно изменить с использованием числового формата, но действительное значение и тип в ячейке изменяются только при настройке нового значения.</span><span class="sxs-lookup"><span data-stu-id="1e012-153">This display can be changed with the number format, but the actual value and type in the cell only changes when a new value is set.</span></span>

<span data-ttu-id="1e012-154">При использовании значения ячейки важно сообщить TypeScript, какое значение вы ожидаете получить из ячейки или диапазона.</span><span class="sxs-lookup"><span data-stu-id="1e012-154">When you are using the cell value, it's important to tell TypeScript what value you are expecting to get from a cell or range.</span></span> <span data-ttu-id="1e012-155">Ячейка содержит один из следующих типов: `string`, `number`или `boolean`.</span><span class="sxs-lookup"><span data-stu-id="1e012-155">A cell contains one of the following types: `string`, `number`, or `boolean`.</span></span> <span data-ttu-id="1e012-156">Чтобы сценарий обрабатывал возвращенные значения как один из этих типов, необходимо объявить тип.</span><span class="sxs-lookup"><span data-stu-id="1e012-156">In order for your script to treat the returned values as one of those types, you must declare the type.</span></span>

<span data-ttu-id="1e012-157">Следующий сценарий получает среднюю цену из таблицы в предыдущем примере.</span><span class="sxs-lookup"><span data-stu-id="1e012-157">The following script gets the average price from the table in the previous sample.</span></span> <span data-ttu-id="1e012-158">Обратите внимание на код `priceRange.getValues() as number[][]`.</span><span class="sxs-lookup"><span data-stu-id="1e012-158">Note the code `priceRange.getValues() as number[][]`.</span></span> <span data-ttu-id="1e012-159">Это [утверждает](https://www.typescriptlang.org/docs/handbook/2/everyday-types.html#type-assertions) `number[][]` в качестве типа значений диапазона.</span><span class="sxs-lookup"><span data-stu-id="1e012-159">This [asserts](https://www.typescriptlang.org/docs/handbook/2/everyday-types.html#type-assertions) the type of the range values to be a `number[][]`.</span></span> <span data-ttu-id="1e012-160">После этого все значения в этом массиве могут рассматриваться как числа в сценарии.</span><span class="sxs-lookup"><span data-stu-id="1e012-160">All the values in that array can then be treated as numbers in the script.</span></span>

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

## <a name="charts-tables-and-other-data-objects"></a><span data-ttu-id="1e012-161">Диаграммы, таблицы и другие объекты данных</span><span class="sxs-lookup"><span data-stu-id="1e012-161">Charts, tables, and other data objects</span></span>

<span data-ttu-id="1e012-162">Скрипты могут создавать и управлять структурами данных и визуализациями в Excel.</span><span class="sxs-lookup"><span data-stu-id="1e012-162">Scripts can create and manipulate the data structures and visualizations within Excel.</span></span> <span data-ttu-id="1e012-163">Таблицы и диаграммы являются двумя наиболее часто используемыми объектами, но API поддерживают сводные таблицы, фигуры, изображения и многое другое.</span><span class="sxs-lookup"><span data-stu-id="1e012-163">Tables and charts are two of the more commonly used objects, but the APIs support PivotTables, shapes, images, and more.</span></span> <span data-ttu-id="1e012-164">Они сохраняются в коллекциях, которые рассматриваются далее в этой статье.</span><span class="sxs-lookup"><span data-stu-id="1e012-164">These are stored in collections, which will be discussed later in this article.</span></span>

### <a name="create-a-table"></a><span data-ttu-id="1e012-165">Создание таблицы</span><span class="sxs-lookup"><span data-stu-id="1e012-165">Create a table</span></span>

<span data-ttu-id="1e012-p116">Создайте таблицы с помощью диапазонов данных. Форматирование и элементы управления таблицами (например, фильтры) автоматически применяются к диапазону.</span><span class="sxs-lookup"><span data-stu-id="1e012-p116">Create tables by using data-filled ranges. Formatting and table controls (such as filters) are automatically applied to the range.</span></span>

<span data-ttu-id="1e012-168">Следующий скрипт создает таблицу с использованием диапазонов из предыдущего примера.</span><span class="sxs-lookup"><span data-stu-id="1e012-168">The following script creates a table using the ranges from the previous sample.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
    // Get the active worksheet.
    let sheet = workbook.getActiveWorksheet();

    // Add a table that has headers using the data from B2:E5.
    sheet.addTable("B2:E5", true);
}
```

<span data-ttu-id="1e012-169">Выполнение этого сценария на листе с предыдущими данными создает следующую таблицу:</span><span class="sxs-lookup"><span data-stu-id="1e012-169">Running this script on the worksheet with the previous data creates the following table:</span></span>

:::image type="content" source="../images/table-sample.png" alt-text="Лист, содержащий таблицу, созданную из предыдущей записи о продажах.":::

### <a name="create-a-chart"></a><span data-ttu-id="1e012-171">Создание диаграммы</span><span class="sxs-lookup"><span data-stu-id="1e012-171">Create a chart</span></span>

<span data-ttu-id="1e012-172">Создайте диаграммы для визуализации данных в диапазоне.</span><span class="sxs-lookup"><span data-stu-id="1e012-172">Create charts to visualize the data in a range.</span></span> <span data-ttu-id="1e012-173">Сценарии позволяют создавать десятки разновидностей диаграмм, каждая из которых может быть настроена в соответствии с вашими потребностями.</span><span class="sxs-lookup"><span data-stu-id="1e012-173">Scripts allow for dozens of chart varieties, each of which can be customized to suit your needs.</span></span>

<span data-ttu-id="1e012-174">Следующий скрипт создает простую столбчатую диаграмму для трех элементов и размещает ее на 100 пикселей ниже верхней части листа.</span><span class="sxs-lookup"><span data-stu-id="1e012-174">The following script creates a simple column chart for three items and places it 100 pixels below the top of the worksheet.</span></span>

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

<span data-ttu-id="1e012-175">Запуск этого скрипта на листе с предыдущей таблицей создает следующую диаграмму:</span><span class="sxs-lookup"><span data-stu-id="1e012-175">Running this script on the worksheet with the previous table creates the following chart:</span></span>

:::image type="content" source="../images/chart-sample.png" alt-text="Гистограмма, показывающая количества трех элементов из предыдущей записи о продажах.":::

## <a name="collections"></a><span data-ttu-id="1e012-177">Коллекции</span><span class="sxs-lookup"><span data-stu-id="1e012-177">Collections</span></span>

<span data-ttu-id="1e012-178">Если объект Excel содержит коллекцию из одного или нескольких объектов одного типа, он сохраняет их в массиве.</span><span class="sxs-lookup"><span data-stu-id="1e012-178">When an Excel object has a collection of one or more objects of the same type, it stores them in an array.</span></span> <span data-ttu-id="1e012-179">Например, объект `Workbook` содержит `Worksheet[]`.</span><span class="sxs-lookup"><span data-stu-id="1e012-179">For example, a `Workbook` object contains a `Worksheet[]`.</span></span> <span data-ttu-id="1e012-180">Доступ к этому массиву обеспечивается методом `Workbook.getWorksheets()`.</span><span class="sxs-lookup"><span data-stu-id="1e012-180">This array is accessed by the `Workbook.getWorksheets()` method.</span></span> <span data-ttu-id="1e012-181">Множественные методы `get`, например `Worksheet.getCharts()`, возвращают всю коллекцию объектов в качестве массива.</span><span class="sxs-lookup"><span data-stu-id="1e012-181">`get` methods that are plural, such as `Worksheet.getCharts()`, return the entire object collection as an array.</span></span> <span data-ttu-id="1e012-182">Вы увидите этот шаблон во всех API сценариев Office: объект `Worksheet` использует метод `getTables()`, возвращающий `Table[]`, объект `Table` использует метод `getColumns()`, возвращающий `TableColumn[]`, и т. д.</span><span class="sxs-lookup"><span data-stu-id="1e012-182">You'll see this pattern throughout the Office Scripts APIs: the `Worksheet` object has a `getTables()` method that returns a `Table[]`, the `Table` object has a `getColumns()` method that returns a `TableColumn[]`, as so on.</span></span>

<span data-ttu-id="1e012-183">Возвращаемый массив является обычным массивом, поэтому все обычные операции массивов доступны для вашего сценария.</span><span class="sxs-lookup"><span data-stu-id="1e012-183">The returned array is a normal array, so all the regular array operations are available for your script.</span></span> <span data-ttu-id="1e012-184">Также можно получить доступ к отдельным объектам внутри коллекции с помощью значения индекса массива.</span><span class="sxs-lookup"><span data-stu-id="1e012-184">You can also access individual objects within the collection using the array index value.</span></span> <span data-ttu-id="1e012-185">Например, `workbook.getTables()[0]` возвращает первую таблицу в коллекции.</span><span class="sxs-lookup"><span data-stu-id="1e012-185">For example, `workbook.getTables()[0]` returns the first table in the collection.</span></span> <span data-ttu-id="1e012-186">Дополнительные сведения об использовании встроенных функций массива в структуре сценариев Office см. в статье [Работа с коллекциями](javascript-objects.md#work-with-collections).</span><span class="sxs-lookup"><span data-stu-id="1e012-186">For more information on using the built-in array functionality with the Office Scripts framework, see [Work with collections](javascript-objects.md#work-with-collections).</span></span> 

<span data-ttu-id="1e012-187">Отдельные объекты также доступны из коллекции с помощью метода `get`.</span><span class="sxs-lookup"><span data-stu-id="1e012-187">Individual objects are also accessed from the collection through a `get` method.</span></span> <span data-ttu-id="1e012-188">Одиночные методы `get`, например `Worksheet.getTable(name)`, возвращают один объект и требуют идентификатор или имя конкретного объекта.</span><span class="sxs-lookup"><span data-stu-id="1e012-188">`get` methods that are singular, such as `Worksheet.getTable(name)`, return a single object and require an ID or name for the specific object.</span></span> <span data-ttu-id="1e012-189">Этот идентификатор или имя обычно задается сценарием или с помощью пользовательского интерфейса Excel.</span><span class="sxs-lookup"><span data-stu-id="1e012-189">This ID or name is usually set by the script or through the Excel UI.</span></span>

<span data-ttu-id="1e012-p121">Следующий сценарий возвращает все таблицы в книге. При этом отображаются заголовки, видны кнопки фильтров, а для таблицы устанавливается стиль "TableStyleLight1".</span><span class="sxs-lookup"><span data-stu-id="1e012-p121">The following script gets all tables in the workbook. It then ensures the headers are displays, the filter buttons are visible, and the table style is set to "TableStyleLight1".</span></span>

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

## <a name="add-excel-objects-with-a-script"></a><span data-ttu-id="1e012-192">Добавление объектов Excel с помощью сценария</span><span class="sxs-lookup"><span data-stu-id="1e012-192">Add Excel objects with a script</span></span>

<span data-ttu-id="1e012-193">Можно программным образом добавлять объекты документов, например таблицы или диаграммы, путем вызова соответствующего метода `add`, доступного для родительского объекта.</span><span class="sxs-lookup"><span data-stu-id="1e012-193">You can programmatically add document objects, such as tables or charts, by calling the corresponding `add` method available on the parent object.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="1e012-194">Не следует вручную добавлять объекты в массивы коллекций.</span><span class="sxs-lookup"><span data-stu-id="1e012-194">Do not manually add objects to collection arrays.</span></span> <span data-ttu-id="1e012-195">Используйте методы `add` для родительских объектов. Например, можно добавить `Table` к `Worksheet` методом `Worksheet.addTable`.</span><span class="sxs-lookup"><span data-stu-id="1e012-195">Use the `add` methods on the parent objects For example, add a `Table` to a `Worksheet` with the `Worksheet.addTable` method.</span></span>

<span data-ttu-id="1e012-196">Следующий сценарий создает таблицу в Excel на первом листе книги.</span><span class="sxs-lookup"><span data-stu-id="1e012-196">The following script creates a table in Excel on the first worksheet in the workbook.</span></span> <span data-ttu-id="1e012-197">Обратите внимание, что метод `addTable` возвращает созданную таблицу.</span><span class="sxs-lookup"><span data-stu-id="1e012-197">Note that the created table is returned by the `addTable` method.</span></span>

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
> <span data-ttu-id="1e012-198">Большинство объектов Excel используют метод `setName`.</span><span class="sxs-lookup"><span data-stu-id="1e012-198">Most Excel objects have a `setName` method.</span></span> <span data-ttu-id="1e012-199">Это позволяет легко получить доступ к объектам Excel позже в сценарии или в других сценариях для той же книги.</span><span class="sxs-lookup"><span data-stu-id="1e012-199">This gives you an easy way to access Excel objects later in the script or in other scripts for the same workbook.</span></span>

### <a name="verify-an-object-exists-in-the-collection"></a><span data-ttu-id="1e012-200">Проверка существования объекта в коллекции</span><span class="sxs-lookup"><span data-stu-id="1e012-200">Verify an object exists in the collection</span></span>

<span data-ttu-id="1e012-201">Перед продолжением сценариям часто требуется проверить, существует ли таблица или похожий объект.</span><span class="sxs-lookup"><span data-stu-id="1e012-201">Scripts often need to check if a table or similar object exists before continuing.</span></span> <span data-ttu-id="1e012-202">Используйте имена, заданные сценариями или с помощью пользовательского интерфейса Excel, чтобы определить необходимые объекты и действовать соответствующим образом.</span><span class="sxs-lookup"><span data-stu-id="1e012-202">Use the names given by scripts or through the Excel UI to identify necessary objects and act accordingly.</span></span> <span data-ttu-id="1e012-203">Методы `get` возвращают `undefined`, когда запрашиваемый объект отсутствует в коллекции.</span><span class="sxs-lookup"><span data-stu-id="1e012-203">`get` methods return `undefined` when the requested object is not in the collection.</span></span>

<span data-ttu-id="1e012-204">Следующий сценарий запрашивает таблицу MyTable и использует оператор `if...else`, чтобы проверить, найдена ли таблица.</span><span class="sxs-lookup"><span data-stu-id="1e012-204">The following script requests a table named "MyTable" and uses an `if...else` statement to check if the table was found.</span></span>

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

<span data-ttu-id="1e012-205">Распространенный шаблон в сценариях Office — воссоздание таблицы, диаграммы или другого объекта при каждом запуске сценария.</span><span class="sxs-lookup"><span data-stu-id="1e012-205">A common pattern in Office Scripts is to recreate a table, chart, or other object every time the script is run.</span></span> <span data-ttu-id="1e012-206">Если старые данные не нужны, рекомендуется удалить старый объект перед созданием нового.</span><span class="sxs-lookup"><span data-stu-id="1e012-206">If you don't need the old data, it's best to delete the old object before creating the new one.</span></span> <span data-ttu-id="1e012-207">Это позволяет избежать конфликтов имен или других различий, которые могли быть добавлены другими пользователями.</span><span class="sxs-lookup"><span data-stu-id="1e012-207">This avoids name conflicts or other differences that may have been introduced by other users.</span></span>

<span data-ttu-id="1e012-208">Следующий сценарий удаляет таблицу MyTable, если она существует, а затем добавляет новую таблицу с таким же именем.</span><span class="sxs-lookup"><span data-stu-id="1e012-208">The following script removes the table named "MyTable", if it is present, then adds a new table with the same name.</span></span>

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

## <a name="remove-excel-objects-with-a-script"></a><span data-ttu-id="1e012-209">Удаление объектов Excel с помощью сценария</span><span class="sxs-lookup"><span data-stu-id="1e012-209">Remove Excel objects with a script</span></span>

<span data-ttu-id="1e012-210">Чтобы удалить объект, вызовите метод `delete` этого объекта.</span><span class="sxs-lookup"><span data-stu-id="1e012-210">To delete an object, call the object's `delete` method.</span></span>

> [!NOTE]
> <span data-ttu-id="1e012-211">Как и в случае добавления объектов, не следует вручную удалять объекты из массивов коллекций.</span><span class="sxs-lookup"><span data-stu-id="1e012-211">As with adding objects, do not manually remove objects from collection arrays.</span></span> <span data-ttu-id="1e012-212">Используйте методы `delete` для объектов типа коллекции.</span><span class="sxs-lookup"><span data-stu-id="1e012-212">Use the `delete` methods on the collection-type objects.</span></span> <span data-ttu-id="1e012-213">Например, для удаления `Table` из `Worksheet` используйте `Table.delete`.</span><span class="sxs-lookup"><span data-stu-id="1e012-213">For example, remove a `Table` from a `Worksheet` using `Table.delete`.</span></span>

<span data-ttu-id="1e012-214">Следующий сценарий удаляет первый лист в книге.</span><span class="sxs-lookup"><span data-stu-id="1e012-214">The following script removes the first worksheet in the workbook.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
    // Get first worksheet.
    let sheet = workbook.getWorksheets()[0];

    // Remove that worksheet from the workbook.
    sheet.delete();
}
```

## <a name="further-reading-on-the-object-model"></a><span data-ttu-id="1e012-215">Дальнейшее чтение по объектной модели</span><span class="sxs-lookup"><span data-stu-id="1e012-215">Further reading on the object model</span></span>

<span data-ttu-id="1e012-216">[Справочная документация по API сценариев Office](/javascript/api/office-scripts/overview) представляет собой полный список объектов, используемых в сценариях Office.</span><span class="sxs-lookup"><span data-stu-id="1e012-216">The [Office Scripts API reference documentation](/javascript/api/office-scripts/overview) is a comprehensive listing of the objects used in Office Scripts.</span></span> <span data-ttu-id="1e012-217">Там вы можете использовать оглавление, чтобы перейти к любому классу, о котором вы хотите узнать больше.</span><span class="sxs-lookup"><span data-stu-id="1e012-217">There, you can use the table of contents to navigate to any class you'd like to learn more about.</span></span> <span data-ttu-id="1e012-218">Ниже приведены несколько часто просматриваемых страниц.</span><span class="sxs-lookup"><span data-stu-id="1e012-218">The following are several commonly viewed pages.</span></span>

- [<span data-ttu-id="1e012-219">Chart</span><span class="sxs-lookup"><span data-stu-id="1e012-219">Chart</span></span>](/javascript/api/office-scripts/excelscript/excelscript.chart)
- [<span data-ttu-id="1e012-220">Comment</span><span class="sxs-lookup"><span data-stu-id="1e012-220">Comment</span></span>](/javascript/api/office-scripts/excelscript/excelscript.comment)
- [<span data-ttu-id="1e012-221">PivotTable</span><span class="sxs-lookup"><span data-stu-id="1e012-221">PivotTable</span></span>](/javascript/api/office-scripts/excelscript/excelscript.pivottable)
- [<span data-ttu-id="1e012-222">Range</span><span class="sxs-lookup"><span data-stu-id="1e012-222">Range</span></span>](/javascript/api/office-scripts/excelscript/excelscript.range)
- [<span data-ttu-id="1e012-223">RangeFormat</span><span class="sxs-lookup"><span data-stu-id="1e012-223">RangeFormat</span></span>](/javascript/api/office-scripts/excelscript/excelscript.rangeformat)
- [<span data-ttu-id="1e012-224">Shape</span><span class="sxs-lookup"><span data-stu-id="1e012-224">Shape</span></span>](/javascript/api/office-scripts/excelscript/excelscript.shape)
- [<span data-ttu-id="1e012-225">Table</span><span class="sxs-lookup"><span data-stu-id="1e012-225">Table</span></span>](/javascript/api/office-scripts/excelscript/excelscript.table)
- [<span data-ttu-id="1e012-226">Workbook</span><span class="sxs-lookup"><span data-stu-id="1e012-226">Workbook</span></span>](/javascript/api/office-scripts/excelscript/excelscript.workbook)
- [<span data-ttu-id="1e012-227">Worksheet</span><span class="sxs-lookup"><span data-stu-id="1e012-227">Worksheet</span></span>](/javascript/api/office-scripts/excelscript/excelscript.worksheet)

## <a name="see-also"></a><span data-ttu-id="1e012-228">См. также</span><span class="sxs-lookup"><span data-stu-id="1e012-228">See also</span></span>

- [<span data-ttu-id="1e012-229">Запись, редактирование и создание сценариев Office в Excel в Интернете</span><span class="sxs-lookup"><span data-stu-id="1e012-229">Record, edit, and create Office Scripts in Excel on the web</span></span>](../tutorials/excel-tutorial.md)
- [<span data-ttu-id="1e012-230">Чтение данных рабочей книги с помощью сценариев Office в Excel в Интернете</span><span class="sxs-lookup"><span data-stu-id="1e012-230">Read workbook data with Office Scripts in Excel on the web</span></span>](../tutorials/excel-read-tutorial.md)
- [<span data-ttu-id="1e012-231">Справочник API для сценариев Office</span><span class="sxs-lookup"><span data-stu-id="1e012-231">Office Scripts API reference</span></span>](/javascript/api/office-scripts/overview)
- [<span data-ttu-id="1e012-232">Использование встроенных объектов JavaScript в сценариях Office</span><span class="sxs-lookup"><span data-stu-id="1e012-232">Using built-in JavaScript objects in Office Scripts</span></span>](javascript-objects.md)
- [<span data-ttu-id="1e012-233">Рекомендации по сценариям Office</span><span class="sxs-lookup"><span data-stu-id="1e012-233">Best practices in Office Scripts</span></span>](best-practices.md)
