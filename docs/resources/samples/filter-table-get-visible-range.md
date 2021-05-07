---
title: Фильтр Excel таблицы и получить видимый диапазон
description: Узнайте, как использовать Office скрипты для фильтрации таблицы Excel и получения видимого диапазона в качестве массива объектов.
ms.date: 04/28/2021
localization_priority: Normal
ms.openlocfilehash: a310857e6055b3da57c353dc7ad78a6fbdd86d4e
ms.sourcegitcommit: f7a7aebfb687f2a35dbed07ed62ff352a114525a
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 05/06/2021
ms.locfileid: "52232377"
---
# <a name="filter-excel-table-and-get-visible-range-as-a-json-object"></a>Фильтр Excel таблицы и получить видимый диапазон в качестве объекта JSON

Этот пример фильтрует таблицу Excel и возвращает видимый диапазон в качестве объекта JSON. Этот JSON может быть предоставлен потоку Power Automate как часть более крупного решения.

## <a name="example-scenario"></a>Пример сценария

* Нанесите фильтр на столбец таблицы.
* Извлекать видимый диапазон после фильтрации.
* Сборка и возвращение объекта с [определенной структурой JSON.](#sample-json)

## <a name="sample-code-filter-a-table-and-get-visible-range"></a>Пример кода: фильтруем таблицу и получаем видимый диапазон

Следующий сценарий фильтрует таблицу и получает видимый диапазон.

Скачайте пример файла <a href="table-filter.xlsx">table-filter.xlsx</a> и используйте его с помощью этого скрипта, чтобы попробовать его самостоятельно!

```TypeScript
function main(workbook: ExcelScript.Workbook): ReturnTemplate {
  const table1 = workbook.getTable("Table1");
  const keyColumnValues: string [] = table1.getColumnByName('Station').getRangeBetweenHeaderAndTotal().getValues().map(v => v[0] as string);
  const uniqueKeys = keyColumnValues.filter((v, i, a) => a.indexOf(v) === i);

  console.log(uniqueKeys);
  const returnObj: ReturnTemplate = {}

  uniqueKeys.forEach((key: string) => {
    table1.getColumnByName('Station').getFilter()
      .applyValuesFilter([key]);
    const rangeView = table1.getRange().getVisibleView();
    returnObj[key] = returnObjectFromValues(rangeView.getValues() as string[][]);
  })
  table1.getColumnByName('Station').getFilter().clear();
  console.log(JSON.stringify(returnObj));
  return returnObj
}

function returnObjectFromValues(values: string[][]): BasicObj[] {
  let objArray = [];
  let objKeys: string[] = [];
  for (let i=0; i < values.length; i++) {
    if (i===0) {
      objKeys = values[i]
      continue;
    }
    let obj = {}
    for (let j=0; j < values[i].length; j++) {
      obj[objKeys[j]] = values[i][j]
    }
    objArray.push(obj);
  }
  return objArray;
}

interface BasicObj {
  [key: string] : string
}

interface ReturnTemplate {
  [key: string]: BasicObj[]
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

## <a name="training-video-filter-an-excel-table-and-get-the-visible-range"></a>Обучающее видео: фильтровать таблицу Excel и получить видимый диапазон

[Смотреть Sudhi Ramamurthy ходить через этот пример на YouTube](https://youtu.be/Mv7BrvPq84A).
