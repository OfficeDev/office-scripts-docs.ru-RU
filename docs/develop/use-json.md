---
title: Использование JSON для передачи данных в скрипты Office и из них
description: Узнайте, как структурировать данные в объекты JSON для использования с внешними вызовами и Power Automate
ms.date: 06/10/2022
ms.localizationpriority: medium
ms.openlocfilehash: 753097183a18f5d20ca2c78a3748c7a1d968ad42
ms.sourcegitcommit: dd01979d34b3499360d2f79a56f8a8f24f480eed
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 06/15/2022
ms.locfileid: "66088160"
---
# <a name="use-json-to-pass-data-to-and-from-office-scripts"></a>Использование JSON для передачи данных в скрипты Office и из них

[JSON (нотация объектов JavaScript)](https://www.w3schools.com/whatis/whatis_json.asp) — это формат для хранения и передачи данных. Каждый объект JSON представляет собой коллекцию пар "имя-значение", которые можно определить при создании. JSON полезен для Office, так как он может обрабатывать произвольные сложности диапазонов, таблиц и других шаблонов данных в Excel. JSON позволяет анализировать входящие данные из веб-служб [](external-calls.md) и передавать сложные объекты [Power Automate потоков](power-automate-integration.md).

В этой статье основное внимание уделяется использованию JSON с Office скриптами. Мы рекомендуем сначала получить дополнительные сведения о формате из таких статей, как [введение в JSON](https://www.w3schools.com/js/js_json_intro.asp) из W3 Schools.

## <a name="parse-json-data-into-a-range-or-table"></a>Анализ данных JSON в диапазоне или таблице

Массивы объектов JSON обеспечивают согласованный способ передачи строк табличных данных между приложениями и веб-службами. В таких случаях каждый объект JSON представляет строку, а свойства — столбцы. Скрипт Office может выполнять циклический цикл по массиву JSON и повторно соблюсти его в виде 2D-массива. Затем этот массив устанавливается как значения диапазона и сохраняется в книге. Имена свойств также можно добавить в качестве заголовков для создания таблицы.

В следующем скрипте показаны данные JSON, преобразуемые в таблицу. Обратите внимание, что данные не взяты из внешнего источника. Это рассматривается далее в этой статье.

```typescript
/**
 * Sample JSON data. This would be replaced by external calls or
 * parameters getting data from Power Automate in a production script.
 */
const jsonData = [
  { "Action": "Edit", /* Action property with value of "Edit". */
    "N": 3370, /* N property with value of 3370. */
    "Percent": 17.85 /* Percent property with value of 17.85. */
  },
  // The rest of the object entries follow the same pattern.
  { "Action": "Paste", "N": 1171, "Percent": 6.2 },
  { "Action": "Clear", "N": 599, "Percent": 3.17 },
  { "Action": "Insert", "N": 352, "Percent": 1.86 },
  { "Action": "Delete", "N": 350, "Percent": 1.85 },
  { "Action": "Refresh", "N": 314, "Percent": 1.66 },
  { "Action": "Fill", "N": 286, "Percent": 1.51 },
];

/**
 * This script converts JSON data to an Excel table.
 */
function main(workbook: ExcelScript.Workbook) {
  // Create a new worksheet to store the imported data.
  const newSheet = workbook.addWorksheet();
  newSheet.activate();

  // Determine the data's shape by getting the properties in one object.
  // This assumes all the JSON objects have the same properties.
  const columnNames = getPropertiesFromJson(jsonData[0]);

  // Create the table headers using the property names.
  const headerRange = newSheet.getRangeByIndexes(0, 0, 1, columnNames.length);
  headerRange.setValues([columnNames]);

  // Create a new table with the headers.
  const newTable = newSheet.addTable(headerRange, true);

  // Add each object in the array of JSON objects to the table.
  const tableValues = jsonData.map(row => convertJsonToRow(row));
  newTable.addRows(-1, tableValues);
}

/**
 * This function turns a JSON object into an array to be used as a table row.
 */
function convertJsonToRow(obj: object) {
  const array: (string | number)[] = [];

  // Loop over each property and get the value. Their order will be the same as the column headers.
  for (let value in obj) {
    array.push(obj[value]);
  }
  return array;
}

/**
 * This function gets the property names from a single JSON object.
 */
function getPropertiesFromJson(obj: object) {
  const propertyArray: string[] = [];
  
  // Loop over each property in the object and store the property name in an array.
  for (let property in obj) {
    propertyArray.push(property);
  }

  return propertyArray;
}
```

> [!TIP]
> Если вы знаете структуру JSON, можно создать собственный интерфейс, чтобы упростить получение определенных свойств. Шаги преобразования JSON в массив можно заменить типобезопасными ссылками. В следующем фрагменте кода показаны эти шаги (теперь закомментированные), замененные вызовами, использующими новый `ActionRow` интерфейс. Обратите внимание, что это делает функцию `convertJsonToRow` больше не обязательной.
>
> ```typescript
>   // const tableValues = jsonData.map(row => convertJsonToRow(row));
>   // newTable.addRows(-1, tableValues);
>   // }
>
>      const actionRows: ActionRow[] = jsonData as ActionRow[];
>      // Add each object in the array of JSON objects to the table.
>      const tableValues = actionRows.map(row => [row.Action, row.N, row.Percent]);
>      newTable.addRows(-1, tableValues);
>    }
>    
>    interface ActionRow {
>      Action: string;
>      N: number;
>      Percent: number;
>    }
> ```

### <a name="get-json-data-from-external-sources"></a>Получение данных JSON из внешних источников

Существует два способа импорта данных JSON в книгу с помощью Office скрипта.

- В качестве [параметра](power-automate-integration.md#main-parameters-pass-data-to-a-script) с Power Automate потоком.
- Вызов внешней `fetch` [веб-службы](external-calls.md).

#### <a name="modify-the-sample-to-work-with-power-automate"></a>Измените пример для работы с Power Automate

Данные JSON в Power Automate могут передаваться в виде универсального массива объектов. Добавьте свойство `object[]` в скрипт, чтобы принять эти данные.

```typescript
// For Power Automate, replace the main signature in the previous sample with this one
// and remove the sample data.
function main(workbook: ExcelScript.Workbook, jsonData: object[]) {
```

Затем вы увидите параметр в соединителе Power Automate для добавления `jsonData` в действие **запуска скрипта**.

:::image type="content" source="../images/json-parameter-power-automate.png" alt-text="Соединитель Excel Online (Business), отображающий действие запуска скрипта с параметром jsonData.":::

#### <a name="modify-the-sample-to-use-a-fetch-call"></a>Изменение примера для использования вызова `fetch`

Веб-службы могут отвечать на звонки `fetch` с помощью данных JSON. Это дает скрипту необходимые данные, сохраняя при этом Excel. Дополнительные сведения о внешних `fetch` вызовах и внешних вызовах см. в разделе [Office API](external-calls.md).

```typescript
// For external services, replace the main signature in the previous sample with this one,
// add the fetch call, and remove the sample data.
async function main(workbook: ExcelScript.Workbook) {
  // Replace WEB_SERVICE_URL with the URL of whatever service you need to call.
  const response = await fetch('WEB_SERVICE_URL');
  const jsonData: object[] = await response.json();
```

## <a name="create-json-from-a-range"></a>Создание JSON из диапазона

Строки и столбцы листа часто подразумевают связи между значениями данных. Строка таблицы концептуально сопоставляется с программным объектом, каждый столбец которого является свойством этого объекта. Рассмотрим следующую таблицу данных. Каждая строка представляет транзакцию, записанную в электронной таблице.

|ID |Date     |Amount |Поставщик                        |
|:--|:--------|:------|:-----------------------------|
|1  |6/1/2022 |43,54 долл. США |Лучше всего подходит для вас компания "Органическая компания" |
|2  |6/3/2022 |67,23 долл. США |Лима-Лима и Нью-Куа       |
|3  |6/3/2022 |37,12 долл. США |Лучше всего подходит для вас компания "Органическая компания" |
|4  |6/6/2022 |86,95 долл. США |Виноградник Coho                 |
|5  |6/7/2022 |13,64 долл. США |Лима-Лима и Нью-Куа       |

Каждая транзакция (каждая строка) имеет набор свойств, связанных с ней: "ID", "Date", "Amount" и "Vendor". Его можно с моделированием в Office как объект.

```typescript
// An interface that wraps transaction details as JSON.
interface Transaction {
  "ID": string;
  "Date": number;
  "Amount": number;
  "Vendor": string;
}
```

Строки в образце таблицы соответствуют свойствам в интерфейсе, поэтому скрипт может легко преобразовать каждую строку в `Transaction` объект. Это полезно при выводе данных для Power Automate. Следующий скрипт выполняет итерацию по каждой строке таблицы и добавляет ее в .`Transaction[]`

```typescript
function main(workbook: ExcelScript.Workbook) {
  // Get the table on the current worksheet.
  const table = workbook.getActiveWorksheet().getTables()[0];

  // Create an array of Transactions and add each row to it.
  let transactions: Transaction[] = [];
  const dataValues = table.getRangeBetweenHeaderAndTotal().getValues();
  for (let i = 0; i < dataValues.length; i++) {
    let row = dataValues[i];
    let currentTransaction: Transaction = {
      ID: row[table.getColumnByName("ID").getIndex()] as string,
      Date: row[table.getColumnByName("Date").getIndex()] as number,
      Amount: row[table.getColumnByName("Amount").getIndex()] as number,
      Vendor: row[table.getColumnByName("Vendor").getIndex()] as string
    };
    transactions.push(currentTransaction);
  }

  // Do something with the Transaction objects, such as return them to a Power Automate flow.
  console.log(transactions);
}

// An interface that wraps transaction details as JSON.
interface Transaction {
  "ID": string;
  "Date": number;
  "Amount": number;
  "Vendor": string;
}
```

:::image type="content" source="../images/create-json-console-output.png" alt-text="Выходные данные консоли из предыдущего скрипта, в котором показаны значения свойств объекта.":::

### <a name="use-a-generic-object"></a>Использование универсального объекта

В предыдущем примере предполагается, что значения заголовков таблицы согласованы. Если таблица содержит переменные столбцы, необходимо создать универсальный объект JSON. В следующем скрипте показан скрипт, который регистрирует любую таблицу в виде JSON.

```typescript
function main(workbook: ExcelScript.Workbook) {
  // Get the table on the current worksheet.
  const table = workbook.getActiveWorksheet().getTables()[0];

  // Use the table header names as JSON properties.
  const tableHeaders = table.getHeaderRowRange().getValues()[0] as string[];
  
  // Get each data row in the table.
  const dataValues = table.getRangeBetweenHeaderAndTotal().getValues();
  let jsonArray: object[] = [];

  // For each row, create a JSON object and assign each property to it based on the table headers.
  for (let i = 0; i < dataValues.length; i++) {
    // Create a blank generic JSON object.
    let jsonObject: { [key: string]: string } = {};
    for (let j = 0; j < dataValues[i].length; j++) {
      jsonObject[tableHeaders[j]] = dataValues[i][j] as string;
    }

    jsonArray.push(jsonObject);
  }

  // Do something with the objects, such as return them to a Power Automate flow.
  console.log(jsonArray);
}

```

## <a name="see-also"></a>См. также

- [Поддержка внешнего вызова API в сценариях Office](external-calls.md)
- [Пример. Использование вызовов внешних выборок в Office скриптах](../resources/samples/external-fetch-calls.md)
- [Выполнение Office с помощью Power Automate](power-automate-integration.md)