---
title: Поддержка старых сценариев Office, которые используют асимментные API
description: В этой теме вы можете узнать, как использовать шаблон загрузки и синхронизации для более старых сценариев.
ms.date: 02/08/2021
localization_priority: Normal
ms.openlocfilehash: be7847efe59dc6026875b8a8e3b3c93e0eb82e4d
ms.sourcegitcommit: 345f1dd96d80471b246044b199fe11126a192a88
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 02/14/2021
ms.locfileid: "50242027"
---
# <a name="support-older-office-scripts-that-use-the-async-apis"></a><span data-ttu-id="2d97f-103">Поддержка старых сценариев Office, которые используют асимментные API</span><span class="sxs-lookup"><span data-stu-id="2d97f-103">Support older Office Scripts that use the async APIs</span></span>

<span data-ttu-id="2d97f-104">В этой статье вы научите вас поддерживать и обновлять сценарии, которые используют асимнкторы API старой модели.</span><span class="sxs-lookup"><span data-stu-id="2d97f-104">This article will teach you how to maintain and update scripts that use the older model's async APIs.</span></span> <span data-ttu-id="2d97f-105">Эти API имеют те же основные функции, что и стандартные синхронные API сценариев Office, но они требуют от скрипта управления синхронизацией данных между сценарием и книгой.</span><span class="sxs-lookup"><span data-stu-id="2d97f-105">These APIs have the same core functionality as the now-standard, synchronous Office Scripts APIs, but they require your script to control the data synchronization between the script and the workbook.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="2d97f-106">А async model can only be used with scripts created before the implementation of the current [API model.](scripting-fundamentals.md?view=office-scripts&preserve-view=true)</span><span class="sxs-lookup"><span data-stu-id="2d97f-106">The async model can only be used with scripts created before the implementation of the current [API model](scripting-fundamentals.md?view=office-scripts&preserve-view=true).</span></span> <span data-ttu-id="2d97f-107">Скрипты навсегда заблокированы к модели API, которая у них есть при создании.</span><span class="sxs-lookup"><span data-stu-id="2d97f-107">Scripts are permanently locked to the API model they have upon creation.</span></span> <span data-ttu-id="2d97f-108">Это также означает, что если вы хотите преобразовать старый сценарий в новую модель, необходимо создать совершенно новый сценарий.</span><span class="sxs-lookup"><span data-stu-id="2d97f-108">This also means that if you want to convert an old script to the new model, you must create a brand new script.</span></span> <span data-ttu-id="2d97f-109">При внесении изменений рекомендуется обновить старые сценарии до новой модели, так как текущую модель проще использовать.</span><span class="sxs-lookup"><span data-stu-id="2d97f-109">We recommend you update your old scripts to the new model when making changes, since the current model is easier to use.</span></span> <span data-ttu-id="2d97f-110">В [разделе "Преобразование а async сценариев](#converting-async-scripts-to-the-current-model) в текущую модель" есть советы по этому переходу.</span><span class="sxs-lookup"><span data-stu-id="2d97f-110">The [Converting async scripts to the current model](#converting-async-scripts-to-the-current-model) section has advice on how to make this transition.</span></span>

## <a name="main-function"></a><span data-ttu-id="2d97f-111">Функция `main`</span><span class="sxs-lookup"><span data-stu-id="2d97f-111">`main` function</span></span>

<span data-ttu-id="2d97f-112">Сценарии, которые используют асимронные API, имеют другую `main` функцию.</span><span class="sxs-lookup"><span data-stu-id="2d97f-112">Scripts that use the async APIs have a different `main` function.</span></span> <span data-ttu-id="2d97f-113">Это `async` функция, которая имеет в `Excel.RequestContext` качестве первого параметра.</span><span class="sxs-lookup"><span data-stu-id="2d97f-113">It's an `async` function that has an `Excel.RequestContext` as the first parameter.</span></span>

```TypeScript
async function main(context: Excel.RequestContext) {
    // Your async Office Script
}
```

## <a name="context"></a><span data-ttu-id="2d97f-114">Context</span><span class="sxs-lookup"><span data-stu-id="2d97f-114">Context</span></span>

<span data-ttu-id="2d97f-115">Функция `main` принимает `Excel.RequestContext` параметра с именем `context`.</span><span class="sxs-lookup"><span data-stu-id="2d97f-115">The `main` function accepts an `Excel.RequestContext` parameter, named `context`.</span></span> <span data-ttu-id="2d97f-116">Думайте о `context` как о мосте между вашим сценарием и книгой.</span><span class="sxs-lookup"><span data-stu-id="2d97f-116">Think of `context` as the bridge between your script and the workbook.</span></span> <span data-ttu-id="2d97f-117">Ваш сценарий обращается к книге с помощью `context` объекта и использует этот `context` для отправки данных туда и обратно.</span><span class="sxs-lookup"><span data-stu-id="2d97f-117">Your script accesses the workbook with the `context` object and uses that `context` to send data back and forth.</span></span>

<span data-ttu-id="2d97f-118">Объект `context` необходим, потому что скрипт и Excel работают в разных процессах и местах.</span><span class="sxs-lookup"><span data-stu-id="2d97f-118">The `context` object is necessary because the script and Excel are running in different processes and locations.</span></span> <span data-ttu-id="2d97f-119">Сценарий должен будет внести изменения или запросить данные из рабочей книги в облаке.</span><span class="sxs-lookup"><span data-stu-id="2d97f-119">The script will need to make changes to or query data from the workbook in the cloud.</span></span> <span data-ttu-id="2d97f-120">Объект `context` управляет этими транзакциями.</span><span class="sxs-lookup"><span data-stu-id="2d97f-120">The `context` object manages those transactions.</span></span>

## <a name="sync-and-load"></a><span data-ttu-id="2d97f-121">Синхронизация и загрузка</span><span class="sxs-lookup"><span data-stu-id="2d97f-121">Sync and Load</span></span>

<span data-ttu-id="2d97f-122">Поскольку ваш сценарий и рабочая книга работают в разных местах, любая передача данных между ними занимает много времени.</span><span class="sxs-lookup"><span data-stu-id="2d97f-122">Because your script and workbook run in different locations, any data transfer between the two takes time.</span></span> <span data-ttu-id="2d97f-123">В aync API команды находятся в очереди до тех пор, пока сценарий явным образом не вызывает операцию для синхронизации `sync` сценария и книги.</span><span class="sxs-lookup"><span data-stu-id="2d97f-123">In the async API, commands are queued up until the script explicitly calls the `sync` operation to synchronize the script and workbook.</span></span> <span data-ttu-id="2d97f-124">Ваш скрипт может работать независимо, пока он не выполнит одно из следующих действий:</span><span class="sxs-lookup"><span data-stu-id="2d97f-124">Your script can work independently until it needs to do either of the following:</span></span>

- <span data-ttu-id="2d97f-125">Прочитайте данные из рабочей книги (с помощью операции `load` или метода возвращения [ClientResult](/javascript/api/office/officeextension.clientresult?view=excel-js-online&preserve-view=true)).</span><span class="sxs-lookup"><span data-stu-id="2d97f-125">Read data from the workbook (following a `load` operation or method that returns a [ClientResult](/javascript/api/office/officeextension.clientresult?view=excel-js-online&preserve-view=true)).</span></span>
- <span data-ttu-id="2d97f-126">Запишите данные в рабочую книгу (обычно потому, что сценарий завершен).</span><span class="sxs-lookup"><span data-stu-id="2d97f-126">Write data to the workbook (usually because the script has finished).</span></span>

<span data-ttu-id="2d97f-127">На следующем рисунке показан пример потока управления между сценарием и книгой:</span><span class="sxs-lookup"><span data-stu-id="2d97f-127">The following image shows an example control flow between the script and workbook:</span></span>

![Диаграмма, показывающая операции чтения и записи, идущие в рабочую книгу из сценария.](../images/load-sync.png)

### <a name="sync"></a><span data-ttu-id="2d97f-129">Синхронизировать</span><span class="sxs-lookup"><span data-stu-id="2d97f-129">Sync</span></span>

<span data-ttu-id="2d97f-130">Каждый раз, когда вашему астинкингову сценарию необходимо считывать данные из книги или записывать их в нее, вызовите метод, как `RequestContext.sync` показано ниже:</span><span class="sxs-lookup"><span data-stu-id="2d97f-130">Whenever your async script needs to read data from or write data to the workbook, call the `RequestContext.sync` method as shown here:</span></span>

```TypeScript
await context.sync();
```

> [!NOTE]
> <span data-ttu-id="2d97f-131">`context.sync()` неявно вызывается, когда скрипт заканчивается.</span><span class="sxs-lookup"><span data-stu-id="2d97f-131">`context.sync()` is implicitly called when a script ends.</span></span>

<span data-ttu-id="2d97f-132">После завершения операции `sync` книга обновляется, чтобы отразить все операции записи, указанные сценарием.</span><span class="sxs-lookup"><span data-stu-id="2d97f-132">After the `sync` operation completes, the workbook updates to reflect any write operations that script has specified.</span></span> <span data-ttu-id="2d97f-133">Операция записи — это установка любого свойства объекта Excel (например, ) или вызов метода, который изменяет свойство `range.format.fill.color = "red"` (например, `range.format.autoFitColumns()` ).</span><span class="sxs-lookup"><span data-stu-id="2d97f-133">A write operation is setting any property on a Excel object (e.g., `range.format.fill.color = "red"`) or calling a method that changes a property (e.g., `range.format.autoFitColumns()`).</span></span> <span data-ttu-id="2d97f-134">Операция `sync` также считывает любые значения из рабочей книги, запрошенные сценарием с помощью операции `load` или метода возвращения `ClientResult` (как описано в следующих разделах).</span><span class="sxs-lookup"><span data-stu-id="2d97f-134">The `sync` operation also reads any values from the workbook that the script requested by using a `load` operation or a method that returns a `ClientResult` (as discussed in the next sections).</span></span>

<span data-ttu-id="2d97f-135">Синхронизация вашего сценария с книгой может занять некоторое время, в зависимости от вашей сети.</span><span class="sxs-lookup"><span data-stu-id="2d97f-135">Synchronizing your script with the workbook can take time, depending on your network.</span></span> <span data-ttu-id="2d97f-136">Свести к минимуму `sync` количество вызовов для быстрого запуска скрипта.</span><span class="sxs-lookup"><span data-stu-id="2d97f-136">Minimize the number of `sync` calls to help your script run fast.</span></span> <span data-ttu-id="2d97f-137">В противном случае асинхронные API не быстрее стандартных синхронных API.</span><span class="sxs-lookup"><span data-stu-id="2d97f-137">Otherwise, the async APIs are not faster the standard, synchronous APIs.</span></span>

### <a name="load"></a><span data-ttu-id="2d97f-138">Load</span><span class="sxs-lookup"><span data-stu-id="2d97f-138">Load</span></span>

<span data-ttu-id="2d97f-139">Перед чтением а также сценарий должен загрузить данные из книги.</span><span class="sxs-lookup"><span data-stu-id="2d97f-139">An async script must load data from the workbook before reading it.</span></span> <span data-ttu-id="2d97f-140">Однако загрузка данных из всей книги значительно снизит скорость работы сценария.</span><span class="sxs-lookup"><span data-stu-id="2d97f-140">However, loading data from the entire workbook would greatly reduce the script's speed.</span></span> <span data-ttu-id="2d97f-141">Этот метод позволяет сценарию в частности указать данные, которые `load` должны быть извлечены из книги.</span><span class="sxs-lookup"><span data-stu-id="2d97f-141">The `load` method lets your script specifically state what data should be retrieved from the workbook.</span></span>

<span data-ttu-id="2d97f-142">Метод `load` доступен для каждого объекта Excel.</span><span class="sxs-lookup"><span data-stu-id="2d97f-142">The `load` method is available on every Excel object.</span></span> <span data-ttu-id="2d97f-143">Ваш скрипт должен загрузить свойства объекта, прежде чем он сможет их прочитать.</span><span class="sxs-lookup"><span data-stu-id="2d97f-143">Your script must load an object's properties before it can read them.</span></span> <span data-ttu-id="2d97f-144">Если этого не сделать, будет выявить ошибку.</span><span class="sxs-lookup"><span data-stu-id="2d97f-144">Not doing so results in an error.</span></span>

<span data-ttu-id="2d97f-145">В следующих примерах объект `Range` используется для демонстрации трех способов использования метода `load` для загрузки данных.</span><span class="sxs-lookup"><span data-stu-id="2d97f-145">The following examples use a `Range` object to show the three ways the `load` method can be used to load data.</span></span>

|<span data-ttu-id="2d97f-146">Intent</span><span class="sxs-lookup"><span data-stu-id="2d97f-146">Intent</span></span> |<span data-ttu-id="2d97f-147">Пример команды</span><span class="sxs-lookup"><span data-stu-id="2d97f-147">Example Command</span></span> | <span data-ttu-id="2d97f-148">Эффект</span><span class="sxs-lookup"><span data-stu-id="2d97f-148">Effect</span></span> |
|:--|:--|:--|
|<span data-ttu-id="2d97f-149">Загрузить одно свойство</span><span class="sxs-lookup"><span data-stu-id="2d97f-149">Load one property</span></span> |`myRange.load("values");` | <span data-ttu-id="2d97f-150">Загружает одно свойство, в данном случае двумерный массив значений в этом диапазоне.</span><span class="sxs-lookup"><span data-stu-id="2d97f-150">Loads a single property, in this case the two-dimensional array of values in this range.</span></span> |
|<span data-ttu-id="2d97f-151">Загрузить несколько свойств</span><span class="sxs-lookup"><span data-stu-id="2d97f-151">Load multiple properties</span></span> |`myRange.load("values, rowCount, columnCount");`| <span data-ttu-id="2d97f-152">Загружает все свойства из списка, разделенного запятыми, в этом примере значения, количество строк и количество столбцов.</span><span class="sxs-lookup"><span data-stu-id="2d97f-152">Loads all the properties from a comma-delimited list, in this example the values, row count, and column count.</span></span> |
|<span data-ttu-id="2d97f-153">Загрузить все</span><span class="sxs-lookup"><span data-stu-id="2d97f-153">Load everything</span></span> | `myRange.load();`|<span data-ttu-id="2d97f-154">Загружает все свойства в диапазоне.</span><span class="sxs-lookup"><span data-stu-id="2d97f-154">Loads all the properties on the range.</span></span> <span data-ttu-id="2d97f-155">Это не рекомендуемое решение, так как оно замедлит сценарий, получив ненужные данные.</span><span class="sxs-lookup"><span data-stu-id="2d97f-155">This isn't a recommended solution, since it will slow down your script by getting unnecessary data.</span></span> <span data-ttu-id="2d97f-156">Используйте его только при тестировании скрипта или при необходимости каждого свойства из объекта.</span><span class="sxs-lookup"><span data-stu-id="2d97f-156">Only use this while testing your script or if you need every property from the object.</span></span> |

<span data-ttu-id="2d97f-157">Ваш скрипт должен вызывать `context.sync()` перед чтением любых загруженных значений.</span><span class="sxs-lookup"><span data-stu-id="2d97f-157">Your script must call `context.sync()` before reading any loaded values.</span></span>

```TypeScript
/**
 * This script uses the async API to get the row count for a range.
 * It shows how to load a property in the async model.
 */
async function main(context: Excel.RequestContext) {
    let selectedSheet = context.workbook.worksheets.getActiveWorksheet();
    let range = selectedSheet.getRange("A1:B3");

    // Load the property.
    range.load("rowCount");

    // Synchronize with the workbook to get the property.
    await context.sync();

    // Read and log the property value (3).
    console.log(range.rowCount);
}
```

<span data-ttu-id="2d97f-158">Вы также можете загрузить свойства всей коллекции.</span><span class="sxs-lookup"><span data-stu-id="2d97f-158">You can also load properties across an entire collection.</span></span> <span data-ttu-id="2d97f-159">Каждый объект коллекции в а async API имеет свойство, которое является `items` массивом, содержащим объекты в этой коллекции.</span><span class="sxs-lookup"><span data-stu-id="2d97f-159">Every collection object in the async API has an `items` property that is an array containing the objects in that collection.</span></span> <span data-ttu-id="2d97f-160">Использование `items` в качестве начала иерархического вызова (`items\myProperty`) для `load` загружает указанные свойства для каждого из этих элементов.</span><span class="sxs-lookup"><span data-stu-id="2d97f-160">Using `items` as the start of a hierarchical call (`items\myProperty`) to `load` loads the specified properties on each of those items.</span></span> <span data-ttu-id="2d97f-161">В следующем примере загружается свойство `resolved` для каждых `Comment` объектов в `CommentCollection` объекте рабочего листа.</span><span class="sxs-lookup"><span data-stu-id="2d97f-161">The following example loads the `resolved` property on every `Comment` object in the `CommentCollection` object of a worksheet.</span></span>

```TypeScript
/**
 * This script uses the async API to get resolved property on every comment in the worksheet.
 * It shows how to load a property from every object in a collection.
 */
async function main(context: Excel.RequestContext){
    let selectedSheet = context.workbook.worksheets.getActiveWorksheet();
    let comments = selectedSheet.comments;

    // Load the `resolved` property from every comment in this collection.
    comments.load("items/resolved");

    // Synchronize with the workbook to get the properties.
    await context.sync();
}
```

### <a name="clientresult"></a><span data-ttu-id="2d97f-162">ClientResult</span><span class="sxs-lookup"><span data-stu-id="2d97f-162">ClientResult</span></span>

<span data-ttu-id="2d97f-163">Методы в а async API, которые возвращают сведения из книги, имеют шаблон, аналогичный `load` / `sync` парадигме.</span><span class="sxs-lookup"><span data-stu-id="2d97f-163">Methods in the async API that return information from the workbook have a similar pattern to the `load`/`sync` paradigm.</span></span> <span data-ttu-id="2d97f-164">Например, `TableCollection.getCount` получает количество таблиц в коллекции.</span><span class="sxs-lookup"><span data-stu-id="2d97f-164">As an example, `TableCollection.getCount` gets the number of tables in the collection.</span></span> <span data-ttu-id="2d97f-165">`getCount` возвращает `ClientResult<number>` значение, означающее, что свойство `value` в возвращаемом значении [`ClientResult`](/javascript/api/office/officeextension.clientresult?view=excel-js-online&preserve-view=true) является числом.</span><span class="sxs-lookup"><span data-stu-id="2d97f-165">`getCount` returns a `ClientResult<number>`, meaning the `value` property in the returned [`ClientResult`](/javascript/api/office/officeextension.clientresult?view=excel-js-online&preserve-view=true) is a number.</span></span> <span data-ttu-id="2d97f-166">Скрипт не может получить доступ к этому значению, пока не вызовет `context.sync()`.</span><span class="sxs-lookup"><span data-stu-id="2d97f-166">Your script can't access that value until `context.sync()` is called.</span></span> <span data-ttu-id="2d97f-167">По аналогии с загрузкой свойства, `value` — это локальное пустое значение до вызова `sync`.</span><span class="sxs-lookup"><span data-stu-id="2d97f-167">Much like loading a property, the `value` is a local "empty" value until that `sync` call.</span></span>

<span data-ttu-id="2d97f-168">Следующий сценарий получает общее количество таблиц в рабочей книге и записывает его в консоль.</span><span class="sxs-lookup"><span data-stu-id="2d97f-168">The following script gets the total number of tables in the workbook and logs that number to the console.</span></span>

```TypeScript
/**
 * This script uses the async API to get the table count of the workbook.
 * It shows how ClientResult objects return workbook information.
 */
async function main(context: Excel.RequestContext) {
    let tableCount = context.workbook.tables.getCount();

    // This sync call implicitly loads tableCount.value.
    // Any other ClientResult values are loaded too.
    await context.sync();

    // Trying to log the value before calling sync would throw an error.
    console.log(tableCount.value);
}
```

## <a name="converting-async-scripts-to-the-current-model"></a><span data-ttu-id="2d97f-169">Преобразование а async сценариев в текущую модель</span><span class="sxs-lookup"><span data-stu-id="2d97f-169">Converting async scripts to the current model</span></span>

<span data-ttu-id="2d97f-170">Текущая модель API не использует `load` `sync` , или `RequestContext` .</span><span class="sxs-lookup"><span data-stu-id="2d97f-170">The current API model doesn't use `load`, `sync`, or a `RequestContext`.</span></span> <span data-ttu-id="2d97f-171">Это значительно упрощает написание и обслуживание сценариев.</span><span class="sxs-lookup"><span data-stu-id="2d97f-171">This makes the scripts much easier to write and maintain.</span></span> <span data-ttu-id="2d97f-172">Лучшим ресурсом для преобразования старых сценариев является [Stack Overflow.](https://stackoverflow.com/questions/tagged/office-scripts)</span><span class="sxs-lookup"><span data-stu-id="2d97f-172">Your best resource for converting old scripts is [Stack Overflow](https://stackoverflow.com/questions/tagged/office-scripts).</span></span> <span data-ttu-id="2d97f-173">Там вы можете обратиться за помощью к сообществу по определенным сценариям.</span><span class="sxs-lookup"><span data-stu-id="2d97f-173">There, you can ask the community for help with specific scenarios.</span></span> <span data-ttu-id="2d97f-174">Следующие рекомендации помогут вам с общими действиями, которые необходимо предпринять.</span><span class="sxs-lookup"><span data-stu-id="2d97f-174">The following guidance should help outline the general steps you'll need to take.</span></span>

1. <span data-ttu-id="2d97f-175">Создайте новый сценарий и скопируйте в него старый а async-код.</span><span class="sxs-lookup"><span data-stu-id="2d97f-175">Create a new script and copy the old async code into it.</span></span> <span data-ttu-id="2d97f-176">Не включай старую `main` сигнатуру метода, используя вместо этого `function main(workbook: ExcelScript.Workbook)` текущую.</span><span class="sxs-lookup"><span data-stu-id="2d97f-176">Be sure not to include the old `main` method signature, using the current `function main(workbook: ExcelScript.Workbook)` instead.</span></span>

2. <span data-ttu-id="2d97f-177">Удалите все `load` вызовы `sync` и вызовы.</span><span class="sxs-lookup"><span data-stu-id="2d97f-177">Remove all the `load` and `sync` calls.</span></span> <span data-ttu-id="2d97f-178">Они больше не нужны.</span><span class="sxs-lookup"><span data-stu-id="2d97f-178">They are no longer necessary.</span></span>

3. <span data-ttu-id="2d97f-179">Удалены все свойства.</span><span class="sxs-lookup"><span data-stu-id="2d97f-179">All properties have been removed.</span></span> <span data-ttu-id="2d97f-180">Теперь к этим объектам можно получить доступ с помощью методов и методов, поэтому вам потребуется переключить эти ссылки на `get` `set` свойства на вызовы методов.</span><span class="sxs-lookup"><span data-stu-id="2d97f-180">You now access those objects through `get` and `set` methods, so you'll need to switch those property references to method calls.</span></span> <span data-ttu-id="2d97f-181">Например, вместо того чтобы устанавливать цвет заливки ячейки с помощью доступа к свойству, как по этому: `mySheet.getRange("A2:C2").format.fill.color = "blue";` теперь вы будете использовать такие методы: `mySheet.getRange("A2:C2").getFormat().getFill().setColor("blue");`</span><span class="sxs-lookup"><span data-stu-id="2d97f-181">For example, instead of setting a cell's fill color through property access like this: `mySheet.getRange("A2:C2").format.fill.color = "blue";`, you'll now use methods like this: `mySheet.getRange("A2:C2").getFormat().getFill().setColor("blue");`</span></span>

4. <span data-ttu-id="2d97f-182">Классы коллекций заменены массивами.</span><span class="sxs-lookup"><span data-stu-id="2d97f-182">Collection classes have been replaced by arrays.</span></span> <span data-ttu-id="2d97f-183">Методы и методы этих классов коллекции были перемещены в объект, который владеет коллекцией, поэтому ссылки должны обновляться `add` `get` соответствующим образом.</span><span class="sxs-lookup"><span data-stu-id="2d97f-183">The `add` and `get` methods of those collection classes were moved to the object that owned the collection, so your references must be updated accordingly.</span></span> <span data-ttu-id="2d97f-184">Например, чтобы получить диаграмму с именем MyChart с первого таблицы в книге, используйте следующий код: `workbook.getWorksheets()[0].getChart("MyChart");` .</span><span class="sxs-lookup"><span data-stu-id="2d97f-184">For example, to get a chart named "MyChart" from the first worksheet in the workbook, use the following code: `workbook.getWorksheets()[0].getChart("MyChart");`.</span></span> <span data-ttu-id="2d97f-185">Обратите `[0]` внимание, что нужно получить доступ к первому значению `Worksheet[]` возвращаемого `getWorksheets()` .</span><span class="sxs-lookup"><span data-stu-id="2d97f-185">Note the `[0]` to access the first value of the `Worksheet[]` returned by `getWorksheets()`.</span></span>

5. <span data-ttu-id="2d97f-186">Некоторые методы были переименованы для ясности и добавлены для удобства.</span><span class="sxs-lookup"><span data-stu-id="2d97f-186">Some methods have been renamed for clarity and added for convenience.</span></span> <span data-ttu-id="2d97f-187">Дополнительные сведения можно [получить в справочнике по API](/javascript/api/office-scripts/overview?view=office-scripts&preserve-view=true) сценариев Office.</span><span class="sxs-lookup"><span data-stu-id="2d97f-187">Please consult the [Office Scripts API reference](/javascript/api/office-scripts/overview?view=office-scripts&preserve-view=true) for more details.</span></span>

## <a name="office-scripts-async-api-reference-documentation"></a><span data-ttu-id="2d97f-188">Справочная документация по API сценариев Office</span><span class="sxs-lookup"><span data-stu-id="2d97f-188">Office Scripts async API reference documentation</span></span>

<span data-ttu-id="2d97f-189">API-aync эквивалентны API, используемым в надстройки Office. Справочная документация находится в разделе Excel справочника по [API JavaScript](/javascript/api/excel?view=excel-js-online&preserve-view=true)для надстройки Office.</span><span class="sxs-lookup"><span data-stu-id="2d97f-189">The async APIs are equivalent to those used in Office Add-ins. The reference documentation is found in [the Excel section of the Office Add-ins JavaScript API reference](/javascript/api/excel?view=excel-js-online&preserve-view=true).</span></span>
