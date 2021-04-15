---
title: Поддержка старых скриптов Office, которые используют API async
description: Праймер в API API Office Scripts Async и использование шаблона нагрузки и синхронизации для старых скриптов.
ms.date: 02/08/2021
localization_priority: Normal
ms.openlocfilehash: 143f52a7ffefb4f19ee36ba4343fd7c2f1cbdffe
ms.sourcegitcommit: 45ffe3dbd2c834b78592ad35928cf8096f5e80bc
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 04/14/2021
ms.locfileid: "51755079"
---
# <a name="support-older-office-scripts-that-use-the-async-apis"></a>Поддержка старых скриптов Office, которые используют API async

В этой статье будет поучено, как поддерживать и обновлять скрипты, которые используют API async старшей модели. Эти API имеют те же основные функции, что и стандартные, синхронные API office Scripts, но они требуют, чтобы ваш скрипт контролировал синхронизацию данных между сценарием и книгой.

> [!IMPORTANT]
> Модель async можно использовать только со сценариями, созданными до реализации текущей [модели API.](scripting-fundamentals.md) Скрипты навсегда заблокированы в модели API, которая имеется при создании. Это также означает, что если вы хотите преобразовать старый сценарий в новую модель, необходимо создать новый скрипт. При внесении изменений рекомендуется обновить старые скрипты в новую модель, так как текущая модель проще в использовании. Сценарии [преобразования async](#converting-async-scripts-to-the-current-model) в текущий раздел модели имеет рекомендации по этому переходу.

## <a name="main-function"></a>Функция `main`

Скрипты, которые используют API async, имеют другую `main` функцию. Это `async` функция, которая имеет в `Excel.RequestContext` качестве первого параметра.

```TypeScript
async function main(context: Excel.RequestContext) {
    // Your async Office Script
}
```

## <a name="context"></a>Context

Функция `main` принимает `Excel.RequestContext` параметра с именем `context`. Думайте о `context` как о мосте между вашим сценарием и книгой. Ваш сценарий обращается к книге с помощью `context` объекта и использует этот `context` для отправки данных туда и обратно.

Объект `context` необходим, потому что скрипт и Excel работают в разных процессах и местах. Сценарий должен будет внести изменения или запросить данные из рабочей книги в облаке. Объект `context` управляет этими транзакциями.

## <a name="sync-and-load"></a>Синхронизация и загрузка

Поскольку ваш сценарий и рабочая книга работают в разных местах, любая передача данных между ними занимает много времени. В API async команды выстраиваются в очередь до тех пор, пока сценарий явно не вызывает операцию для синхронизации `sync` сценария и книги. Ваш скрипт может работать независимо, пока он не выполнит одно из следующих действий:

- Прочитайте данные из рабочей книги (с помощью операции `load` или метода возвращения [ClientResult](/javascript/api/office/officeextension.clientresult?view=excel-js-online&preserve-view=true)).
- Запишите данные в рабочую книгу (обычно потому, что сценарий завершен).

На следующем рисунке показан пример потока управления между сценарием и книгой:

:::image type="content" source="../images/load-sync.png" alt-text="Диаграмма, показывающая операции чтения и записи, идущие в рабочую книгу из сценария.":::

### <a name="sync"></a>Синхронизировать

Всякий раз, когда сценарию async необходимо читать данные из книги или записывать их в книгу, вызывайте `RequestContext.sync` метод, как показано здесь:

```TypeScript
await context.sync();
```

> [!NOTE]
> `context.sync()` неявно вызывается, когда скрипт заканчивается.

После завершения операции `sync` книга обновляется, чтобы отразить все операции записи, указанные сценарием. Операция записи устанавливает любое свойство объекта Excel (например, ) или вызывает метод, который изменяет свойство `range.format.fill.color = "red"` (например, `range.format.autoFitColumns()` ). Операция `sync` также считывает любые значения из рабочей книги, запрошенные сценарием с помощью операции `load` или метода возвращения `ClientResult` (как описано в следующих разделах).

Синхронизация вашего сценария с книгой может занять некоторое время, в зависимости от вашей сети. Свести к минимуму `sync` количество вызовов для быстрого запуска сценария. В противном случае API async не быстрее стандартных синхронных API.

### <a name="load"></a>Load

Перед чтением скрипт async должен загружать данные из книги. Однако загрузка данных из всей книги значительно снизит скорость скрипта. Этот метод позволяет скрипту конкретно указать, какие данные должны `load` быть извлечены из книги.

Метод `load` доступен для каждого объекта Excel. Ваш скрипт должен загрузить свойства объекта, прежде чем он сможет их прочитать. Это не приводит к ошибке.

В следующих примерах объект `Range` используется для демонстрации трех способов использования метода `load` для загрузки данных.

|Intent |Пример команды | Эффект |
|:--|:--|:--|
|Загрузить одно свойство |`myRange.load("values");` | Загружает одно свойство, в данном случае двумерный массив значений в этом диапазоне. |
|Загрузить несколько свойств |`myRange.load("values, rowCount, columnCount");`| Загружает все свойства из списка, разделенного запятыми, в этом примере значения, количество строк и количество столбцов. |
|Загрузить все | `myRange.load();`|Загружает все свойства в диапазоне. Это не рекомендуемое решение, так как оно замедлит сценарий, получив ненужные данные. Используйте его только при тестировании скрипта или при необходимости каждого свойства объекта. |

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

Вы также можете загрузить свойства всей коллекции. Каждый объект коллекции в API async имеет свойство, которое является `items` массивом, содержащим объекты в этой коллекции. Использование `items` в качестве начала иерархического вызова (`items\myProperty`) для `load` загружает указанные свойства для каждого из этих элементов. В следующем примере загружается свойство `resolved` для каждых `Comment` объектов в `CommentCollection` объекте рабочего листа.

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

Методы в API async, возвращаемой из книги, имеют аналогичный шаблон `load` / `sync` парадигмы. Например, `TableCollection.getCount` получает количество таблиц в коллекции. `getCount` возвращает `ClientResult<number>`. Это означает, что свойство `value` возвращаемого [`ClientResult`](/javascript/api/office/officeextension.clientresult?view=excel-js-online&preserve-view=true) выражено числом. Сценарий не может получить доступ к этому значению, пока не вызовет `context.sync()`. По аналогии с загрузкой свойства, `value` — это локальное пустое значение до вызова `sync`.

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

## <a name="converting-async-scripts-to-the-current-model"></a>Преобразование скриптов async в текущую модель

Текущая модель API не использует `load` `sync` , или `RequestContext` . Это значительно упрощает написание и обслуживание сценариев. Лучшим ресурсом для преобразования старых скриптов является [переполнение стека.](https://stackoverflow.com/questions/tagged/office-scripts) Там вы можете обратиться к сообществу за помощью в определенных сценариях. Следующие рекомендации должны помочь в описании общих действий, которые необходимо предпринять.

1. Создайте новый скрипт и скопируйте в него старый код async. Не следует включать старую подпись `main` метода, используя вместо нее `function main(workbook: ExcelScript.Workbook)` текущий.

2. Удалите `load` все `sync` вызовы и вызовы. Они больше не нужны.

3. Все свойства удалены. Теперь вы получите доступ к этим объектам с помощью и методами, поэтому вам потребуется переключить эти ссылки `get` `set` свойств на вызовы методов. Например, вместо настройки цвета заполнения ячейки с помощью доступа к свойству, как это: теперь вы будете `mySheet.getRange("A2:C2").format.fill.color = "blue";` использовать методы, как это: `mySheet.getRange("A2:C2").getFormat().getFill().setColor("blue");`

4. Классы коллекции заменены массивами. Методы и методы этих классов коллекции были перемещены в объект, который владел коллекцией, поэтому ваши ссылки должны `add` `get` обновляться соответствующим образом. Например, чтобы получить диаграмму с именем "MyChart" из первого таблицы в книге, используйте следующий код: `workbook.getWorksheets()[0].getChart("MyChart");` . Обратите внимание `[0]` на доступ к первому значению `Worksheet[]` возвращаемого `getWorksheets()` .

5. Некоторые методы были переименованы для ясности и добавлены для удобства. Дополнительные сведения можно получить в ссылке [на API](/javascript/api/office-scripts/overview) сценариев Office.

## <a name="office-scripts-async-api-reference-documentation"></a>Справочная документация office Scripts async API

API async эквивалентны API, используемым в надстройки Office. Эталонная документация находится в разделе Excel ссылки [на API JavaScript надстройки Office.](/javascript/api/excel?view=excel-js-online&preserve-view=true)
