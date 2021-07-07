---
title: Фильтр Excel таблицы и получить видимый диапазон
description: Узнайте, как использовать Office скрипты для фильтрации таблицы Excel и получения видимого диапазона в качестве массива объектов.
ms.date: 06/29/2021
localization_priority: Normal
ms.openlocfilehash: b19b826f95c7e7aeb331130fde05afaafe500c3d
ms.sourcegitcommit: 211c157ca746e266eeb079f5fa1925a1e35ab702
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/07/2021
ms.locfileid: "53313955"
---
# <a name="filter-excel-table-and-get-visible-range-as-a-json-object"></a><span data-ttu-id="42496-103">Фильтр Excel таблицы и получить видимый диапазон в качестве объекта JSON</span><span class="sxs-lookup"><span data-stu-id="42496-103">Filter Excel table and get visible range as a JSON object</span></span>

<span data-ttu-id="42496-104">Этот пример фильтрует таблицу Excel и возвращает видимый диапазон в качестве объекта JSON.</span><span class="sxs-lookup"><span data-stu-id="42496-104">This sample filters an Excel table and returns the visible range as a JSON object.</span></span> <span data-ttu-id="42496-105">Этот JSON может быть предоставлен потоку Power Automate как часть более крупного решения.</span><span class="sxs-lookup"><span data-stu-id="42496-105">This JSON could be provided to a Power Automate flow as part of a larger solution.</span></span>

## <a name="example-scenario"></a><span data-ttu-id="42496-106">Пример сценария</span><span class="sxs-lookup"><span data-stu-id="42496-106">Example scenario</span></span>

* <span data-ttu-id="42496-107">Нанесите фильтр на столбец таблицы.</span><span class="sxs-lookup"><span data-stu-id="42496-107">Apply a filter to a table column.</span></span>
* <span data-ttu-id="42496-108">Извлекать видимый диапазон после фильтрации.</span><span class="sxs-lookup"><span data-stu-id="42496-108">Extract the visible range after filtering.</span></span>
* <span data-ttu-id="42496-109">Сборка и возвращение объекта с [определенной структурой JSON.](#sample-json)</span><span class="sxs-lookup"><span data-stu-id="42496-109">Assemble and return an object with a [specific JSON structure](#sample-json).</span></span>

## <a name="sample-excel-file"></a><span data-ttu-id="42496-110">Пример Excel файла</span><span class="sxs-lookup"><span data-stu-id="42496-110">Sample Excel file</span></span>

<span data-ttu-id="42496-111">Скачайте <a href="table-filter.xlsx">table-filter.xlsx</a> для готовой к использованию книги.</span><span class="sxs-lookup"><span data-stu-id="42496-111">Download <a href="table-filter.xlsx">table-filter.xlsx</a> for a ready-to-use workbook.</span></span> <span data-ttu-id="42496-112">Добавьте следующий скрипт, чтобы попробовать пример самостоятельно!</span><span class="sxs-lookup"><span data-stu-id="42496-112">Add the following script to try the sample yourself!</span></span>

## <a name="sample-code-filter-a-table-and-get-visible-range"></a><span data-ttu-id="42496-113">Пример кода: фильтруем таблицу и получаем видимый диапазон</span><span class="sxs-lookup"><span data-stu-id="42496-113">Sample code: Filter a table and get visible range</span></span>

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
  let objectArray = [];
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

### <a name="sample-json"></a><span data-ttu-id="42496-114">Пример JSON</span><span class="sxs-lookup"><span data-stu-id="42496-114">Sample JSON</span></span>

<span data-ttu-id="42496-115">Каждый ключ представляет уникальное значение таблицы.</span><span class="sxs-lookup"><span data-stu-id="42496-115">Each key represents a unique value of a table.</span></span> <span data-ttu-id="42496-116">Каждый экземпляр массива представляет строку, которая видна при применении соответствующего фильтра.</span><span class="sxs-lookup"><span data-stu-id="42496-116">Each array instance represents the row that is visible when the corresponding filter is applied.</span></span>

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

## <a name="training-video-filter-an-excel-table-and-get-the-visible-range"></a><span data-ttu-id="42496-117">Обучающее видео: фильтровать таблицу Excel и получить видимый диапазон</span><span class="sxs-lookup"><span data-stu-id="42496-117">Training video: Filter an Excel table and get the visible range</span></span>

<span data-ttu-id="42496-118">[Смотреть Sudhi Ramamurthy ходить через этот пример на YouTube](https://youtu.be/Mv7BrvPq84A).</span><span class="sxs-lookup"><span data-stu-id="42496-118">[Watch Sudhi Ramamurthy walk through this sample on YouTube](https://youtu.be/Mv7BrvPq84A).</span></span>
