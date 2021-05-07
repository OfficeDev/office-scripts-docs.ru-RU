---
title: Выходные Excel как JSON
description: Узнайте, как вывод данных Excel таблицы как JSON для использования в Power Automate.
ms.date: 03/18/2021
localization_priority: Normal
ms.openlocfilehash: c6b033a68fdbde2b053f65d1a54db58da6c93b2e
ms.sourcegitcommit: f7a7aebfb687f2a35dbed07ed62ff352a114525a
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 05/06/2021
ms.locfileid: "52232538"
---
# <a name="output-excel-table-data-as-json-for-usage-in-power-automate"></a>Данные Excel таблицы как JSON для использования в Power Automate

Excel таблицы могут быть представлены в виде массива объектов в виде JSON. Каждый объект представляет строку в таблице. Это помогает извлекать данные из Excel в согласованном формате, который виден пользователю. Затем данные могут быть переданы другим системам Power Automate потоками.

_Данные таблицы ввода_

:::image type="content" source="../../images/table-input.png" alt-text="Таблица, показывающая данные таблицы ввода":::

Вариант этого примера также включает гиперссылки в одном из столбцов таблицы. Это позволяет всплыть в JSON дополнительные уровни данных ячейки.

_Данные таблицы ввода, включаемой гиперссылки_

:::image type="content" source="../../images/table-hyperlink-view.png" alt-text="Таблица, показывающая столбец данных таблицы, отформатированный как гиперссылки":::

_Диалоговое окно для редактирования гиперссылки_

:::image type="content" source="../../images/table-hyperlink-edit.png" alt-text="Диалоговое окно Изменить гиперссылку, отображающий параметры для изменения гиперссылки":::

## <a name="sample-excel-file"></a>Пример Excel файла

Скачайте файл <a href="table-data-with-hyperlinks.xlsx">table-data-with-hyperlinks.xlsx, </a> используемый в этих примерах, и попробуйте его самостоятельно!

## <a name="sample-code-return-table-data-as-json"></a>Пример кода: данные таблицы возврата в качестве JSON

> [!NOTE]
> Вы можете изменить `interface TableData` структуру, чтобы соответствовать столбцам таблицы. Обратите внимание, что для имен столбцов с пробелами обязательно поместите ключ в кавычках, например в `"Event ID"` примере.

```TypeScript
function main(workbook: ExcelScript.Workbook): TableData[] {
  const table = workbook.getWorksheet('PlainTable').getTables()[0];
  // If you know the table name, you can also do the following:
  // const table = workbook.getTable('Table13436');
  const texts = table.getRange().getTexts();
  let returnObjects: TableData[] = [];
  if (table.getRowCount() > 0)  {
    returnObjects = returnObjectFromValues(texts);
  } 
  console.log(JSON.stringify(returnObjects));  
  return returnObjects
}

function returnObjectFromValues(values: string[][]): TableData[] {
  let objArray = [];
  let objKeys: string[] = [];
  for (let i = 0; i < values.length; i++) {
    if (i === 0) {
      objKeys = values[i]
      continue;
    }
    let obj = {}
    for (let j = 0; j < values[i].length; j++) {
      obj[objKeys[j]] = values[i][j]
    }
    objArray.push(obj);
  }
  return objArray as TableData[];
}

interface BasicObj {
  [key: string]: string
}

interface TableData {
  "Event ID": string
  Date: string
  Location: string
  Capacity: string
  Speakers: string
}
```

### <a name="sample-output"></a>Пример выходных данных

```json
[{
    "Event ID": "E107",
    "Date": "2020-12-10",
    "Location": "Montgomery",
    "Capacity": "10",
    "Speakers&quot;: &quot;Debra Berger"
}, {
    "Event ID": "E108",
    "Date": "2020-12-11",
    "Location": "Montgomery",
    "Capacity": "10",
    "Speakers&quot;: &quot;Delia Dennis"
}, {
    "Event ID": "E109",
    "Date": "2020-12-12",
    "Location": "Montgomery",
    "Capacity": "10",
    "Speakers&quot;: &quot;Diego Siciliani"
}, {
    "Event ID": "E110",
    "Date": "2020-12-13",
    "Location": "Boise",
    "Capacity": "25",
    "Speakers&quot;: &quot;Gerhart Moller"
}, {
    "Event ID": "E111",
    "Date": "2020-12-14",
    "Location": "Salt Lake City",
    "Capacity": "20",
    "Speakers&quot;: &quot;Grady Archie"
}, {
    "Event ID": "E112",
    "Date": "2020-12-15",
    "Location": "Fremont",
    "Capacity": "25",
    "Speakers&quot;: &quot;Irvin Sayers"
}, {
    "Event ID": "E113",
    "Date": "2020-12-16",
    "Location": "Salt Lake City",
    "Capacity": "20",
    "Speakers&quot;: &quot;Isaiah Langer"
}, {
    "Event ID": "E114",
    "Date": "2020-12-17",
    "Location": "Salt Lake City",
    "Capacity": "20",
    "Speakers&quot;: &quot;Johanna Lorenz"
}]
```

## <a name="sample-code-return-table-data-as-json-with-hyperlink-text"></a>Пример кода. Возвращаем данные таблицы как JSON с текстом гиперссылки

> [!NOTE]
> Сценарий всегда извлекает гиперссылки из 4-го столбца (индекс 0) таблицы. Вы можете изменить этот порядок или включить несколько столбцов в качестве данных гиперссылки, изменяя код под комментарием `// For the 4th column (0 index), extract the hyperlink and use that instead of text.`

```TypeScript
function main(workbook: ExcelScript.Workbook): TableData[] {
  const table = workbook.getWorksheet('WithHyperLink').getTables()[0];
  const range = table.getRange();
  // If you know the table name, you can also do the following:
  // const table = workbook.getTable('Table13436');
  const texts = table.getRange().getTexts();
  let returnObjects: TableData[] = [];
  if (table.getRowCount() > 0)  {
    returnObjects = returnObjectFromValues(texts, range);
  } 
  console.log(JSON.stringify(returnObjects));  
  return returnObjects
}

function returnObjectFromValues(values: string[][], range: ExcelScript.Range): TableData[] {
  let objArray = [];
  let objKeys: string[] = [];
  for (let i = 0; i < values.length; i++) {
    if (i === 0) {
      objKeys = values[i]
      continue;
    }
    let obj = {}
    for (let j = 0; j < values[i].length; j++) {
      // For the 4th column (0 index), extract the hyperlink and use that instead of text. 
      if (j === 4) {
        obj[objKeys[j]] = range.getCell(i, j).getHyperlink().address;
      } else {
        obj[objKeys[j]] = values[i][j];
      }
    }
    objArray.push(obj);
  }
  return objArray as TableData[];
}

interface BasicObj {
  [key: string]: string
}

interface TableData {
  "Event ID": string
  Date: string
  Location: string
  Capacity: string
  "Search link": string
  Speakers: string
}
```

### <a name="sample-output"></a>Пример выходных данных

```json
[{
    "Event ID": "E107",
    "Date": "2020-12-10",
    "Location": "Montgomery",
    "Capacity": "10",
    "Search link": "https://www.google.com/search?q=Montgomery",
    "Speakers&quot;: &quot;Debra Berger"
}, {
    "Event ID": "E108",
    "Date": "2020-12-11",
    "Location": "Montgomery",
    "Capacity": "10",
    "Search link": "https://www.google.com/search?q=Montgomery",
    "Speakers&quot;: &quot;Delia Dennis"
}, {
    "Event ID": "E109",
    "Date": "2020-12-12",
    "Location": "Montgomery",
    "Capacity": "10",
    "Search link": "https://www.google.com/search?q=Montgomery",
    "Speakers&quot;: &quot;Diego Siciliani"
}, {
    "Event ID": "E110",
    "Date": "2020-12-13",
    "Location": "Boise",
    "Capacity": "25",
    "Search link": "https://www.google.com/search?q=Boise",
    "Speakers&quot;: &quot;Gerhart Moller"
}, {
    "Event ID": "E111",
    "Date": "2020-12-14",
    "Location": "Salt Lake City",
    "Capacity": "20",
    "Search link": "https://www.google.com/search?q=salt+lake+city",
    "Speakers&quot;: &quot;Grady Archie"
}, {
    "Event ID": "E112",
    "Date": "2020-12-15",
    "Location": "Fremont",
    "Capacity": "25",
    "Search link": "https://www.google.com/search?q=Fremont",
    "Speakers&quot;: &quot;Irvin Sayers"
}, {
    "Event ID": "E113",
    "Date": "2020-12-16",
    "Location": "Salt Lake City",
    "Capacity": "20",
    "Search link": "https://www.google.com/search?q=salt+lake+city",
    "Speakers&quot;: &quot;Isaiah Langer"
}, {
    "Event ID": "E114",
    "Date": "2020-12-17",
    "Location": "Salt Lake City",
    "Capacity": "20",
    "Search link": "https://www.google.com/search?q=salt+lake+city",
    "Speakers&quot;: &quot;Johanna Lorenz"
}]
```

## <a name="use-in-power-automate"></a>Использование в Power Automate

О том, как использовать такой сценарий в Power Automate, см. в [Power Automate.](../../tutorials/excel-power-automate-returns.md#create-an-automated-workflow-with-power-automate)
