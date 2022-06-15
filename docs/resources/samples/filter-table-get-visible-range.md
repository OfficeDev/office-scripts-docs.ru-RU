---
title: Фильтрация Excel таблицы и получение видимого диапазона
description: Узнайте, как использовать Office скрипты для фильтрации Excel таблицы и получения видимого диапазона в виде массива объектов.
ms.date: 03/10/2022
ms.localizationpriority: medium
ms.openlocfilehash: 103ec97111720ab872c0be843aa0573781d98c44
ms.sourcegitcommit: dd01979d34b3499360d2f79a56f8a8f24f480eed
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 06/15/2022
ms.locfileid: "66088087"
---
# <a name="filter-excel-table-and-get-visible-range-as-a-json-object"></a>Фильтрация Excel таблицы и получение видимого диапазона в виде объекта JSON

Этот пример фильтрует Excel таблицу и возвращает видимый диапазон в виде [объекта JSON](https://www.w3schools.com/whatis/whatis_json.asp). Этот JSON может быть предоставлен Power Automate как часть более крупного решения.

## <a name="example-scenario"></a>Пример сценария

* Применение фильтра к столбцу таблицы.
* Извлеките видимый диапазон после фильтрации.
* Соберите и вернитесь к объекту с [определенной структурой JSON](#sample-json).

## <a name="sample-excel-file"></a>Пример Excel файла

<a href="table-filter.xlsx"> Скачайтеtable-filter.xlsx</a> для готовой к использованию книги. Добавьте следующий скрипт, чтобы попробовать пример самостоятельно!

## <a name="sample-code-filter-a-table-and-get-visible-range"></a>Пример кода: фильтрация таблицы и получение видимого диапазона

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

Каждый ключ представляет уникальное значение таблицы. Каждый экземпляр массива представляет строку, видимую при применении соответствующего фильтра. Дополнительные сведения о работе с JSON см. в статье "Использование JSON для передачи данных в Office [скрипты и из них"](../../develop/use-json.md).

```json
{
  "Station-1": [{
    "Station": "Station-1",
    "Shift": "Morning",
    "Date": "27-Oct-20",
    "Responsible": "Debra Berger",
    "Reason": ""
  }, {
    "Station": "Station-1",
    "Shift": "Afternoon",
    "Date": "27-Oct-20",
    "Responsible": "Delia Dennis",
    "Reason": ""
  }, {
    "Station": "Station-1",
    "Shift": "Afternoon",
    "Date": "28-Oct-20",
    "Responsible": "Lidia Holloway",
    "Reason": ""
  }],
  "Station-2": [{
    "Station": "Station-2",
    "Shift": "Morning",
    "Date": "27-Oct-20",
    "Responsible": "Gerhart Moller",
    "Reason": ""
  }, {
    "Station": "Station-2",
    "Shift": "Afternoon",
    "Date": "28-Oct-20",
    "Responsible": "Grady Archie",
    "Reason": ""
  }],
  "Station-3": [{
    "Station": "Station-3",
    "Shift": "Morning",
    "Date": "27-Oct-20",
    "Responsible": "Isaiah Langer",
    "Reason": ""
  }]
}
```

## <a name="training-video-filter-an-excel-table-and-get-the-visible-range"></a>Обучающее видео. Фильтрация Excel таблицы и получение видимого диапазона

[Просмотрите этот пример на YouTube](https://youtu.be/Mv7BrvPq84A), чтобы просмотреть этот пример.
