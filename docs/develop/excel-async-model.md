---
title: Поддержка старых Office скриптов, которые используют API async
description: Праймер на основе Office Async API и как использовать шаблон нагрузки/синхронизации для старых скриптов.
ms.date: 05/17/2021
localization_priority: Normal
ms.openlocfilehash: 80a1c0dec5393d8882ddb37eea5f81ef23b1ebb1
ms.sourcegitcommit: 4687693f02fc90a57ba30c461f35046e02e6f5fb
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 05/19/2021
ms.locfileid: "52545077"
---
# <a name="support-older-office-scripts-that-use-the-async-apis"></a>Поддержка старых Office скриптов, которые используют API async

Эта статья учит вас, как поддерживать и обновлять скрипты, которые используют API-api старой модели async. Эти API имеют ту же основную функциональность, что и стандартные, синхронные API Office Scripts, но они требуют, чтобы ваш скрипт контролировал синхронизацию данных между скриптом и рабочей книгой.

> [!IMPORTANT]
> Модель async может использоваться только со скриптами, созданными до реализации текущей [модели API.](scripting-fundamentals.md) Скрипты постоянно заблокированы в модели API, которую они имеют при создании. Это также означает, что если вы хотите преобразовать старый скрипт в новую модель, необходимо создать совершенно новый скрипт. Мы рекомендуем вам обновить старые скрипты до новой модели при внесении изменений, так как текущая модель проще в использовании. В [скриптах Converting async в текущий раздел модели есть](#convert-async-scripts-to-the-current-model) советы о том, как сделать этот переход.

## <a name="older-main-function-signature"></a>Старая `main` подпись функции

Сценарии, используя API async, имеют другую `main` функцию. Это функция, `async` которая имеет в качестве первого `Excel.RequestContext` параметра.

```TypeScript
async function main(context: Excel.RequestContext) {
    // Your async Office Script
}
```

## <a name="context"></a>Context

Функция `main` принимает `Excel.RequestContext` параметра с именем `context`. Думайте о `context` как о мосте между вашим сценарием и книгой. Ваш сценарий обращается к книге с помощью `context` объекта и использует этот `context` для отправки данных туда и обратно.

Объект `context` необходим, потому что скрипт и Excel работают в разных процессах и местах. Сценарий должен будет внести изменения или запросить данные из рабочей книги в облаке. Объект `context` управляет этими транзакциями.

## <a name="sync-and-load"></a>Синхронизация и загрузка

Поскольку ваш сценарий и рабочая книга работают в разных местах, любая передача данных между ними занимает много времени. В API async команды выстраиваются в очередь до тех пор, пока скрипт явно не вызывает `sync` операцию для синхронизации сценария и рабочей книги. Ваш скрипт может работать независимо, пока он не выполнит одно из следующих действий:

- Прочитайте данные из рабочей книги (с помощью операции `load` или метода возвращения [ClientResult](/javascript/api/office/officeextension.clientresult?view=excel-js-online&preserve-view=true)).
- Запишите данные в рабочую книгу (обычно потому, что сценарий завершен).

На следующем рисунке показан пример потока управления между сценарием и книгой:

:::image type="content" source="../images/load-sync.png" alt-text="Диаграмма, показывающая чтение и написание операций, иных в трудовую книжку из сценария":::

### <a name="sync"></a>Синхронизировать

Всякий раз, когда ваш скрипт async должен читать данные или писать данные в трудовую книжку, позвоните `RequestContext.sync` методу, как показано в следующем фрагменте кода:

```TypeScript
await context.sync();
```

> [!NOTE]
> `context.sync()` неявно вызывается, когда скрипт заканчивается.

После завершения операции `sync` книга обновляется, чтобы отразить все операции записи, указанные сценарием. Операция записи устанавливает любое свойство на объекте Excel (например,) `range.format.fill.color = "red"` или называет метод, который изменяет свойство (например, `range.format.autoFitColumns()` ). Операция `sync` также считывает любые значения из рабочей книги, запрошенные сценарием с помощью операции `load` или метода возвращения `ClientResult` (как описано в следующих разделах).

Синхронизация вашего сценария с книгой может занять некоторое время, в зависимости от вашей сети. Свести к минимуму количество `sync` вызовов, чтобы помочь скрипту работать быстро. В противном случае API async не являются более быстрыми стандартными, синхронными API.

### <a name="load"></a>Load

Скрипт async должен загрузить данные из рабочей книги перед чтением. Однако загрузка данных из всей рабочей книги значительно снизит скорость работы скрипта. Метод `load` позволяет скрипту конкретно узнать, какие данные следует извлечь из рабочей книги.

Метод `load` доступен для каждого объекта Excel. Ваш скрипт должен загрузить свойства объекта, прежде чем он сможет их прочитать. Не делать этого приводит к ошибке.

В следующих примерах объект `Range` используется для демонстрации трех способов использования метода `load` для загрузки данных.

|Intent |Пример команды | Эффект |
|:--|:--|:--|
|Загрузить одно свойство |`myRange.load("values");` | Загружает одно свойство, в данном случае двумерный массив значений в этом диапазоне. |
|Загрузить несколько свойств |`myRange.load("values, rowCount, columnCount");`| Загружает все свойства из списка, разделенного запятыми, в этом примере значения, количество строк и количество столбцов. |
|Загрузить все | `myRange.load();`|Загружает все свойства в диапазоне. Это не рекомендуемое решение, так как оно замедлит работу скрипта за счет получения ненужных данных. Используйте это только при тестировании скрипта или если вам нужно каждое свойство от объекта. |

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

Вы также можете загрузить свойства всей коллекции. Каждый объект коллекции в API async имеет `items` свойство, которое представляет собой массив, содержащий объекты в этой коллекции. Использование `items` в качестве начала иерархического вызова (`items\myProperty`) для `load` загружает указанные свойства для каждого из этих элементов. В следующем примере загружается свойство `resolved` для каждых `Comment` объектов в `CommentCollection` объекте рабочего листа.

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

Методы в API async, которые возвращают информацию из рабочей книги, имеют аналогичную схему `load` / `sync` парадигмы. Например, `TableCollection.getCount` получает количество таблиц в коллекции. `getCount` возвращает `ClientResult<number>`. Это означает, что свойство `value` возвращаемого [`ClientResult`](/javascript/api/office/officeextension.clientresult?view=excel-js-online&preserve-view=true) выражено числом. Сценарий не может получить доступ к этому значению, пока не вызовет `context.sync()`. По аналогии с загрузкой свойства, `value` — это локальное пустое значение до вызова `sync`.

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

## <a name="convert-async-scripts-to-the-current-model"></a>Преобразование скриптов async в текущую модель

Текущая модель API не `load` `sync` используется, или `RequestContext` . Это значительно упрощает написание и обслуживание скриптов. Лучшим ресурсом для преобразования старых скриптов является [корпорация Майкрософт&A.](/answers/topics/office-scripts-dev.html) Там вы можете обратиться к сообществу за помощью в конкретных сценариях. Следующее руководство должно помочь наметить общие шаги, которые необходимо предпринять.

1. Создайте новый скрипт и скопировать старый код async в него. Убедитесь в том, чтобы не включать `main` старую подпись метода, используя ток `function main(workbook: ExcelScript.Workbook)` вместо.

2. Удалите все `load` и `sync` звонки. Они больше не нужны.

3. Все свойства удалены. Теперь доступ к этим объектам `get` и `set` методам, так что вам нужно переключить эти ссылки свойств на вызовы метода. Например, вместо того, чтобы устанавливать цвет заполнения ячейки через доступ к свойству, как это: `mySheet.getRange("A2:C2").format.fill.color = "blue";` , Теперь вы будете использовать методы, как это: `mySheet.getRange("A2:C2").getFormat().getFill().setColor("blue");`

4. Классы коллекции были заменены массивами. Методы `add` `get` и методы этих классов коллекции были перемещены на объект, который владел коллекцией, поэтому ваши ссылки должны быть соответствующим образом обновлены. Например, чтобы получить диаграмму под названием "MyChart" с первого листа в рабочей книге, используйте следующий код: `workbook.getWorksheets()[0].getChart("MyChart");` . Обратите внимание `[0]` на доступ к первому значению `Worksheet[]` возвращенного `getWorksheets()` .

5. Некоторые методы были переименованы для ясности и добавлены для удобства. Для получения более [подробной информации Office ссылку на API для](/javascript/api/office-scripts/overview) всех скриптов.

## <a name="office-scripts-async-api-reference-documentation"></a>Office Скрипты async справочная документация API

API async эквивалентны тем, которые используются Office дополнительных дополнительных ва-си. Справочная документация найдена [в Excel разделе ссылки Office Add-ins JavaScript API.](/javascript/api/excel?view=excel-js-online&preserve-view=true)
