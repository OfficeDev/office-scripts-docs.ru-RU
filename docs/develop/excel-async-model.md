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
# <a name="support-older-office-scripts-that-use-the-async-apis"></a>Поддержка старых сценариев Office, которые используют асимментные API

В этой статье вы научите вас поддерживать и обновлять сценарии, которые используют асимнкторы API старой модели. Эти API имеют те же основные функции, что и стандартные синхронные API сценариев Office, но они требуют от скрипта управления синхронизацией данных между сценарием и книгой.

> [!IMPORTANT]
> А async model can only be used with scripts created before the implementation of the current [API model.](scripting-fundamentals.md?view=office-scripts&preserve-view=true) Скрипты навсегда заблокированы к модели API, которая у них есть при создании. Это также означает, что если вы хотите преобразовать старый сценарий в новую модель, необходимо создать совершенно новый сценарий. При внесении изменений рекомендуется обновить старые сценарии до новой модели, так как текущую модель проще использовать. В [разделе "Преобразование а async сценариев](#converting-async-scripts-to-the-current-model) в текущую модель" есть советы по этому переходу.

## <a name="main-function"></a>Функция `main`

Сценарии, которые используют асимронные API, имеют другую `main` функцию. Это `async` функция, которая имеет в `Excel.RequestContext` качестве первого параметра.

```TypeScript
async function main(context: Excel.RequestContext) {
    // Your async Office Script
}
```

## <a name="context"></a>Context

Функция `main` принимает `Excel.RequestContext` параметра с именем `context`. Думайте о `context` как о мосте между вашим сценарием и книгой. Ваш сценарий обращается к книге с помощью `context` объекта и использует этот `context` для отправки данных туда и обратно.

Объект `context` необходим, потому что скрипт и Excel работают в разных процессах и местах. Сценарий должен будет внести изменения или запросить данные из рабочей книги в облаке. Объект `context` управляет этими транзакциями.

## <a name="sync-and-load"></a>Синхронизация и загрузка

Поскольку ваш сценарий и рабочая книга работают в разных местах, любая передача данных между ними занимает много времени. В aync API команды находятся в очереди до тех пор, пока сценарий явным образом не вызывает операцию для синхронизации `sync` сценария и книги. Ваш скрипт может работать независимо, пока он не выполнит одно из следующих действий:

- Прочитайте данные из рабочей книги (с помощью операции `load` или метода возвращения [ClientResult](/javascript/api/office/officeextension.clientresult?view=excel-js-online&preserve-view=true)).
- Запишите данные в рабочую книгу (обычно потому, что сценарий завершен).

На следующем рисунке показан пример потока управления между сценарием и книгой:

![Диаграмма, показывающая операции чтения и записи, идущие в рабочую книгу из сценария.](../images/load-sync.png)

### <a name="sync"></a>Синхронизировать

Каждый раз, когда вашему астинкингову сценарию необходимо считывать данные из книги или записывать их в нее, вызовите метод, как `RequestContext.sync` показано ниже:

```TypeScript
await context.sync();
```

> [!NOTE]
> `context.sync()` неявно вызывается, когда скрипт заканчивается.

После завершения операции `sync` книга обновляется, чтобы отразить все операции записи, указанные сценарием. Операция записи — это установка любого свойства объекта Excel (например, ) или вызов метода, который изменяет свойство `range.format.fill.color = "red"` (например, `range.format.autoFitColumns()` ). Операция `sync` также считывает любые значения из рабочей книги, запрошенные сценарием с помощью операции `load` или метода возвращения `ClientResult` (как описано в следующих разделах).

Синхронизация вашего сценария с книгой может занять некоторое время, в зависимости от вашей сети. Свести к минимуму `sync` количество вызовов для быстрого запуска скрипта. В противном случае асинхронные API не быстрее стандартных синхронных API.

### <a name="load"></a>Load

Перед чтением а также сценарий должен загрузить данные из книги. Однако загрузка данных из всей книги значительно снизит скорость работы сценария. Этот метод позволяет сценарию в частности указать данные, которые `load` должны быть извлечены из книги.

Метод `load` доступен для каждого объекта Excel. Ваш скрипт должен загрузить свойства объекта, прежде чем он сможет их прочитать. Если этого не сделать, будет выявить ошибку.

В следующих примерах объект `Range` используется для демонстрации трех способов использования метода `load` для загрузки данных.

|Intent |Пример команды | Эффект |
|:--|:--|:--|
|Загрузить одно свойство |`myRange.load("values");` | Загружает одно свойство, в данном случае двумерный массив значений в этом диапазоне. |
|Загрузить несколько свойств |`myRange.load("values, rowCount, columnCount");`| Загружает все свойства из списка, разделенного запятыми, в этом примере значения, количество строк и количество столбцов. |
|Загрузить все | `myRange.load();`|Загружает все свойства в диапазоне. Это не рекомендуемое решение, так как оно замедлит сценарий, получив ненужные данные. Используйте его только при тестировании скрипта или при необходимости каждого свойства из объекта. |

Ваш скрипт должен вызывать `context.sync()` перед чтением любых загруженных значений.

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

Вы также можете загрузить свойства всей коллекции. Каждый объект коллекции в а async API имеет свойство, которое является `items` массивом, содержащим объекты в этой коллекции. Использование `items` в качестве начала иерархического вызова (`items\myProperty`) для `load` загружает указанные свойства для каждого из этих элементов. В следующем примере загружается свойство `resolved` для каждых `Comment` объектов в `CommentCollection` объекте рабочего листа.

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

### <a name="clientresult"></a>ClientResult

Методы в а async API, которые возвращают сведения из книги, имеют шаблон, аналогичный `load` / `sync` парадигме. Например, `TableCollection.getCount` получает количество таблиц в коллекции. `getCount` возвращает `ClientResult<number>` значение, означающее, что свойство `value` в возвращаемом значении [`ClientResult`](/javascript/api/office/officeextension.clientresult?view=excel-js-online&preserve-view=true) является числом. Скрипт не может получить доступ к этому значению, пока не вызовет `context.sync()`. По аналогии с загрузкой свойства, `value` — это локальное пустое значение до вызова `sync`.

Следующий сценарий получает общее количество таблиц в рабочей книге и записывает его в консоль.

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

## <a name="converting-async-scripts-to-the-current-model"></a>Преобразование а async сценариев в текущую модель

Текущая модель API не использует `load` `sync` , или `RequestContext` . Это значительно упрощает написание и обслуживание сценариев. Лучшим ресурсом для преобразования старых сценариев является [Stack Overflow.](https://stackoverflow.com/questions/tagged/office-scripts) Там вы можете обратиться за помощью к сообществу по определенным сценариям. Следующие рекомендации помогут вам с общими действиями, которые необходимо предпринять.

1. Создайте новый сценарий и скопируйте в него старый а async-код. Не включай старую `main` сигнатуру метода, используя вместо этого `function main(workbook: ExcelScript.Workbook)` текущую.

2. Удалите все `load` вызовы `sync` и вызовы. Они больше не нужны.

3. Удалены все свойства. Теперь к этим объектам можно получить доступ с помощью методов и методов, поэтому вам потребуется переключить эти ссылки на `get` `set` свойства на вызовы методов. Например, вместо того чтобы устанавливать цвет заливки ячейки с помощью доступа к свойству, как по этому: `mySheet.getRange("A2:C2").format.fill.color = "blue";` теперь вы будете использовать такие методы: `mySheet.getRange("A2:C2").getFormat().getFill().setColor("blue");`

4. Классы коллекций заменены массивами. Методы и методы этих классов коллекции были перемещены в объект, который владеет коллекцией, поэтому ссылки должны обновляться `add` `get` соответствующим образом. Например, чтобы получить диаграмму с именем MyChart с первого таблицы в книге, используйте следующий код: `workbook.getWorksheets()[0].getChart("MyChart");` . Обратите `[0]` внимание, что нужно получить доступ к первому значению `Worksheet[]` возвращаемого `getWorksheets()` .

5. Некоторые методы были переименованы для ясности и добавлены для удобства. Дополнительные сведения можно [получить в справочнике по API](/javascript/api/office-scripts/overview?view=office-scripts&preserve-view=true) сценариев Office.

## <a name="office-scripts-async-api-reference-documentation"></a>Справочная документация по API сценариев Office

API-aync эквивалентны API, используемым в надстройки Office. Справочная документация находится в разделе Excel справочника по [API JavaScript](/javascript/api/excel?view=excel-js-online&preserve-view=true)для надстройки Office.
