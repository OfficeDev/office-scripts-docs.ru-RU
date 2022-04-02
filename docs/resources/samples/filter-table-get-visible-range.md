---
title: Фильтр Excel таблицы и получить видимый диапазон
description: Узнайте, как использовать Office скрипты для фильтрации таблицы Excel и получения видимого диапазона в качестве массива объектов.
ms.date: 03/10/2022
ms.localizationpriority: medium
ms.openlocfilehash: 09adbdabb64f9cf15b8219cfd3ef2dfa35d30fe2
ms.sourcegitcommit: 7023b9e23499806901a5ecf8ebc460b76887cca6
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/31/2022
ms.locfileid: "64585837"
---
# <a name="filter-excel-table-and-get-visible-range-as-a-json-object"></a>Фильтр Excel таблицы и получить видимый диапазон в качестве объекта JSON

Этот пример фильтрует таблицу Excel и возвращает видимый диапазон в качестве объекта JSON. Этот JSON может быть предоставлен потоку Power Automate как часть более крупного решения.

## <a name="example-scenario"></a>Пример сценария

* Нанесите фильтр на столбец таблицы.
* Извлекать видимый диапазон после фильтрации.
* Сборка и возвращение объекта с [определенной структурой JSON](#sample-json).

## <a name="sample-excel-file"></a>Пример Excel файла

<a href="table-filter.xlsx"> Скачайтеtable-filter.xlsx</a> для готовой к использованию книги. Добавьте следующий скрипт, чтобы попробовать пример самостоятельно!

## <a name="sample-code-filter-a-table-and-get-visible-range"></a>Пример кода: фильтруем таблицу и получаем видимый диапазон

```TypeScript
function main(workbook: ExcelScript.Workbook): ReturnTemplate {
  // Get the "Station" column to use as key values in the filter.
  const table1 = workbook.getTable("Table1");
  const keyColumnValues: string [] = table1.getColumnByName('Station').getRangeBetweenHeaderAndTotal().getValues().map(value => value[0] as string);

  // Filter out repeated keys. This call to `filter` only returns the first instance of every unique element in the array.
  const uniqueKeys = keyColumnValues.filter((value, index, array) => array.indexOf(value) === index);
  console.log(uniqueKeys);

  const stationData: ReturnTemplate = {};

  // Filter the table to show only rows corresponding to each key.
  uniqueKeys.forEach((key: string) => {
    table1.getColumnByName('Station').getFilter()
      .applyValuesFilter([key]);
    
    // Get the visible view when a single filter is active.
    const rangeView = table1.getRange().getVisibleView();

    // Create a JSON object with every visible row.
    stationData[key] = returnObjectFromValues(rangeView.getValues() as string[][]);
  });

  // Remove the filters.
  table1.getColumnByName('Station').getFilter().clear();

  // Log the information and return it for a Power Automate flow.
  console.log(JSON.stringify(stationData));
  return stationData;
}

// This function converts a 2D-array of values into a generic JSON object.
function returnObjectFromValues(values: string[][]): BasicObject[] {
  let objectArray: BasicObject[] = [];
  let objectKeys: string[] = [];
  for (let i = 0; i < values.length; i++) {
    if (i === 0) {
      objectKeys = values[i]
      continue;
    }

    let object = {}
    for (let j = 0; j < values[i].length; j++) {
      object[objectKeys[j]] = values[i][j]
    }

    objectArray.push(object);
  }

  return objectArray;
}

interface BasicObject {
  [key: string] : string
}

interface ReturnTemplate {
  [key: string]: BasicObject[]
}
```

### <a name="sample-json"></a>Пример JSON

Каждый ключ представляет уникальное значение таблицы. Каждый экземпляр массива представляет строку, которая видна при применении соответствующего фильтра.

```json
{
  "Station-1": [{
    "Station": "Station-1",
    "Shift": "Morning",
    "Date": "27-Oct-20",
    "Responsible": "Debra Berger",
    "Reason&quot;: &quot;"
  }, {
    "Station": "Station-1",
    "Shift": "Afternoon",
    "Date": "27-Oct-20",
    "Responsible": "Delia Dennis",
    "Reason&quot;: &quot;"
  }, {
    "Station": "Station-1",
    "Shift": "Afternoon",
    "Date": "28-Oct-20",
    "Responsible": "Lidia Holloway",
    "Reason&quot;: &quot;"
  }],
  "Station-2": [{
    "Station": "Station-2",
    "Shift": "Morning",
    "Date": "27-Oct-20",
    "Responsible": "Gerhart Moller",
    "Reason&quot;: &quot;"
  }, {
    "Station": "Station-2",
    "Shift": "Afternoon",
    "Date": "28-Oct-20",
    "Responsible": "Grady Archie",
    "Reason&quot;: &quot;"
  }],
  "Station-3": [{
    "Station": "Station-3",
    "Shift": "Morning",
    "Date": "27-Oct-20",
    "Responsible": "Isaiah Langer",
    "Reason&quot;: &quot;"
  }]
}
```

## <a name="training-video-filter-an-excel-table-and-get-the-visible-range"></a>Обучающее видео. Фильтр Excel таблицу и получить видимый диапазон

[Посмотрите, как суди Рамамурти (Sudhi Ramamurthy) пройдите этот пример на YouTube](https://youtu.be/Mv7BrvPq84A).
