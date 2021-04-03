---
title: Выходные данные Excel в качестве JSON
description: Узнайте, как вывод данных таблиц Excel в качестве JSON для использования в Power Automate.
ms.date: 03/18/2021
localization_priority: Normal
ms.openlocfilehash: 678506fee0b6a41ede8245fb360d485d635e2d64
ms.sourcegitcommit: 5d24e77df70aa2c1c982275d53213c2a9323ff86
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 04/02/2021
ms.locfileid: "51571428"
---
# <a name="output-excel-table-data-as-json-for-usage-in-power-automate"></a><span data-ttu-id="61780-103">Данные таблицы Output Excel в качестве JSON для использования в Power Automate</span><span class="sxs-lookup"><span data-stu-id="61780-103">Output Excel table data as JSON for usage in Power Automate</span></span>

<span data-ttu-id="61780-104">Данные таблиц Excel могут быть представлены в виде массива объектов в виде JSON.</span><span class="sxs-lookup"><span data-stu-id="61780-104">Excel table data can be represented as an array of objects in the form of JSON.</span></span> <span data-ttu-id="61780-105">Каждый объект представляет строку в таблице.</span><span class="sxs-lookup"><span data-stu-id="61780-105">Each object represents a row in the table.</span></span> <span data-ttu-id="61780-106">Это помогает извлекать данные из Excel в согласованном формате, который виден пользователю.</span><span class="sxs-lookup"><span data-stu-id="61780-106">This helps extract the data from Excel in a consistent format that is visible to the user.</span></span> <span data-ttu-id="61780-107">Затем данные могут быть переданы другим системам с помощью потоков Power Automate.</span><span class="sxs-lookup"><span data-stu-id="61780-107">The data can then be given to other systems through Power Automate flows.</span></span>

<span data-ttu-id="61780-108">_Данные таблицы ввода_</span><span class="sxs-lookup"><span data-stu-id="61780-108">_Input table data_</span></span>

![Снимок экрана, отображающий данные таблицы ввода](../../images/table-input.png)

<span data-ttu-id="61780-110">Вариант этого примера также включает гиперссылки в одном из столбцов таблицы.</span><span class="sxs-lookup"><span data-stu-id="61780-110">A variation of this sample also includes the hyperlinks in one of the table columns.</span></span> <span data-ttu-id="61780-111">Это позволяет всплыть в JSON дополнительные уровни данных ячейки.</span><span class="sxs-lookup"><span data-stu-id="61780-111">This allows additional levels of cell data to be surfaced in the JSON.</span></span>

<span data-ttu-id="61780-112">_Данные таблицы ввода, включаемой гиперссылки_</span><span class="sxs-lookup"><span data-stu-id="61780-112">_Input table data that includes hyperlinks_</span></span>

![Снимок экрана, отображающий данные таблицы, включающие гиперссылки](../../images/table-hyperlink-view.png)

<span data-ttu-id="61780-114">_Диалоговое окно для редактирования гиперссылки_</span><span class="sxs-lookup"><span data-stu-id="61780-114">_Dialog to edit hyperlink_</span></span>

![Снимок экрана, отображающий диалоговое окно для редактирования гиперссылки](../../images/table-hyperlink-edit.png)

## <a name="sample-excel-file"></a><span data-ttu-id="61780-116">Пример файла Excel</span><span class="sxs-lookup"><span data-stu-id="61780-116">Sample Excel file</span></span>

<span data-ttu-id="61780-117">Скачайте файл <a href="table-data-with-hyperlinks.xlsx">table-data-with-hyperlinks.xlsx, </a> используемый в этих примерах, и попробуйте его самостоятельно!</span><span class="sxs-lookup"><span data-stu-id="61780-117">Download the file <a href="table-data-with-hyperlinks.xlsx">table-data-with-hyperlinks.xlsx</a> used in these samples and try it out yourself!</span></span>

## <a name="sample-code-return-table-data-as-json"></a><span data-ttu-id="61780-118">Пример кода: данные таблицы возврата в качестве JSON</span><span class="sxs-lookup"><span data-stu-id="61780-118">Sample code: Return table data as JSON</span></span>

> [!NOTE]
> <span data-ttu-id="61780-119">Вы можете изменить `interface TableData` структуру, чтобы соответствовать столбцам таблицы.</span><span class="sxs-lookup"><span data-stu-id="61780-119">You can change the `interface TableData` structure to match your table columns.</span></span> <span data-ttu-id="61780-120">Обратите внимание, что для имен столбцов с пробелами обязательно поместите ключ в кавычках, например в `"Event ID"` примере.</span><span class="sxs-lookup"><span data-stu-id="61780-120">Note that for column names with spaces, be sure to place your key in quotation marks, such as with `"Event ID"` in the sample.</span></span>

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

### <a name="sample-output"></a><span data-ttu-id="61780-121">Пример выходных данных</span><span class="sxs-lookup"><span data-stu-id="61780-121">Sample output</span></span>

```json
[{
    "Event ID": "E107",
    "Date": "2020-12-10",
    "Location": "Montgomery",
    "Capacity": "10",
    "Speakers": "Debra Berger"
}, {
    "Event ID": "E108",
    "Date": "2020-12-11",
    "Location": "Montgomery",
    "Capacity": "10",
    "Speakers": "Delia Dennis"
}, {
    "Event ID": "E109",
    "Date": "2020-12-12",
    "Location": "Montgomery",
    "Capacity": "10",
    "Speakers": "Diego Siciliani"
}, {
    "Event ID": "E110",
    "Date": "2020-12-13",
    "Location": "Boise",
    "Capacity": "25",
    "Speakers": "Gerhart Moller"
}, {
    "Event ID": "E111",
    "Date": "2020-12-14",
    "Location": "Salt Lake City",
    "Capacity": "20",
    "Speakers": "Grady Archie"
}, {
    "Event ID": "E112",
    "Date": "2020-12-15",
    "Location": "Fremont",
    "Capacity": "25",
    "Speakers": "Irvin Sayers"
}, {
    "Event ID": "E113",
    "Date": "2020-12-16",
    "Location": "Salt Lake City",
    "Capacity": "20",
    "Speakers": "Isaiah Langer"
}, {
    "Event ID": "E114",
    "Date": "2020-12-17",
    "Location": "Salt Lake City",
    "Capacity": "20",
    "Speakers": "Johanna Lorenz"
}]
```

## <a name="sample-code-return-table-data-as-json-with-hyperlink-text"></a><span data-ttu-id="61780-122">Пример кода. Возвращаем данные таблицы как JSON с текстом гиперссылки</span><span class="sxs-lookup"><span data-stu-id="61780-122">Sample code: Return table data as JSON with hyperlink text</span></span>

> [!NOTE]
> <span data-ttu-id="61780-123">Сценарий всегда извлекает гиперссылки из 4-го столбца (индекс 0) таблицы.</span><span class="sxs-lookup"><span data-stu-id="61780-123">The script always extracts hyperlinks from the 4th column (0 index) of the table.</span></span> <span data-ttu-id="61780-124">Вы можете изменить этот порядок или включить несколько столбцов в качестве данных гиперссылки, изменяя код под комментарием `// For the 4th column (0 index), extract the hyperlink and use that instead of text.`</span><span class="sxs-lookup"><span data-stu-id="61780-124">You can change that order or include multiple columns as hyperlink data by modifying the code under the comment `// For the 4th column (0 index), extract the hyperlink and use that instead of text.`</span></span>

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

### <a name="sample-output"></a><span data-ttu-id="61780-125">Пример выходных данных</span><span class="sxs-lookup"><span data-stu-id="61780-125">Sample output</span></span>

```json
[{
    "Event ID": "E107",
    "Date": "2020-12-10",
    "Location": "Montgomery",
    "Capacity": "10",
    "Search link": "https://www.google.com/search?q=Montgomery",
    "Speakers": "Debra Berger"
}, {
    "Event ID": "E108",
    "Date": "2020-12-11",
    "Location": "Montgomery",
    "Capacity": "10",
    "Search link": "https://www.google.com/search?q=Montgomery",
    "Speakers": "Delia Dennis"
}, {
    "Event ID": "E109",
    "Date": "2020-12-12",
    "Location": "Montgomery",
    "Capacity": "10",
    "Search link": "https://www.google.com/search?q=Montgomery",
    "Speakers": "Diego Siciliani"
}, {
    "Event ID": "E110",
    "Date": "2020-12-13",
    "Location": "Boise",
    "Capacity": "25",
    "Search link": "https://www.google.com/search?q=Boise",
    "Speakers": "Gerhart Moller"
}, {
    "Event ID": "E111",
    "Date": "2020-12-14",
    "Location": "Salt Lake City",
    "Capacity": "20",
    "Search link": "https://www.google.com/search?q=salt+lake+city",
    "Speakers": "Grady Archie"
}, {
    "Event ID": "E112",
    "Date": "2020-12-15",
    "Location": "Fremont",
    "Capacity": "25",
    "Search link": "https://www.google.com/search?q=Fremont",
    "Speakers": "Irvin Sayers"
}, {
    "Event ID": "E113",
    "Date": "2020-12-16",
    "Location": "Salt Lake City",
    "Capacity": "20",
    "Search link": "https://www.google.com/search?q=salt+lake+city",
    "Speakers": "Isaiah Langer"
}, {
    "Event ID": "E114",
    "Date": "2020-12-17",
    "Location": "Salt Lake City",
    "Capacity": "20",
    "Search link": "https://www.google.com/search?q=salt+lake+city",
    "Speakers": "Johanna Lorenz"
}]
```

## <a name="use-in-power-automate"></a><span data-ttu-id="61780-126">Использование в Power Automate</span><span class="sxs-lookup"><span data-stu-id="61780-126">Use in Power Automate</span></span>

<span data-ttu-id="61780-127">О том, как использовать такой скрипт в Power Automate, см. в статью Создание автоматизированного рабочего процесса [с помощью Power Automate.](../../tutorials/excel-power-automate-returns.md#create-an-automated-workflow-with-power-automate)</span><span class="sxs-lookup"><span data-stu-id="61780-127">For how to use such a script in Power Automate, see [Create an automated workflow with Power Automate](../../tutorials/excel-power-automate-returns.md#create-an-automated-workflow-with-power-automate).</span></span>
