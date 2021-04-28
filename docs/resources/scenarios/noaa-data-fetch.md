---
title: 'Office Пример сценария: Graph данных уровня воды из NOAA'
description: Пример, который извлекает данные JSON из базы данных NOAA и использует их для создания диаграммы.
ms.date: 04/26/2021
localization_priority: Normal
ms.openlocfilehash: 8aea11f42bf2a81fa53cbf4f6ee7280213b97085
ms.sourcegitcommit: d466b82f27bc61aeba193f902c9bc65ecbf60e4e
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 04/28/2021
ms.locfileid: "52066303"
---
# <a name="office-scripts-sample-scenario-fetch-and-graph-water-level-data-from-noaa"></a><span data-ttu-id="985cd-103">Office Сценарий примера сценариев: извлечение и график данных уровня воды из NOAA</span><span class="sxs-lookup"><span data-stu-id="985cd-103">Office Scripts sample scenario: Fetch and graph water-level data from NOAA</span></span>

<span data-ttu-id="985cd-104">В этом сценарии необходимо совместить уровень воды на станции [National Oceanic and Atmospheric Administration's Seattle.](https://tidesandcurrents.noaa.gov/stationhome.html?id=9447130)</span><span class="sxs-lookup"><span data-stu-id="985cd-104">In this scenario, you need to plot the water level at the [National Oceanic and Atmospheric Administration's Seattle station](https://tidesandcurrents.noaa.gov/stationhome.html?id=9447130).</span></span> <span data-ttu-id="985cd-105">Внешние данные используются для заполнения таблицы и создания диаграммы.</span><span class="sxs-lookup"><span data-stu-id="985cd-105">You'll use external data to populate a spreadsheet and create a chart.</span></span>

<span data-ttu-id="985cd-106">Вы разработает сценарий, который использует команду для запроса базы `fetch` [данных NOAA Tides и Currents.](https://tidesandcurrents.noaa.gov/)</span><span class="sxs-lookup"><span data-stu-id="985cd-106">You'll develop a script that uses the `fetch` command to query the [NOAA Tides and Currents database](https://tidesandcurrents.noaa.gov/).</span></span> <span data-ttu-id="985cd-107">Это позволит получить уровень воды, записанный через заданный промежуток времени.</span><span class="sxs-lookup"><span data-stu-id="985cd-107">That will get the water level recorded across a given time span.</span></span> <span data-ttu-id="985cd-108">Сведения будут возвращены в качестве JSON, поэтому часть сценария будет переводить их в значения диапазона.</span><span class="sxs-lookup"><span data-stu-id="985cd-108">The information will be returned as JSON, so part of the script will translate that into range values.</span></span> <span data-ttu-id="985cd-109">После того как данные будут в таблице, они будут использоваться для сделайте диаграмму.</span><span class="sxs-lookup"><span data-stu-id="985cd-109">Once the data is in the spreadsheet, it will be used to make a chart.</span></span>

## <a name="scripting-skills-covered"></a><span data-ttu-id="985cd-110">Навыки скриптов, охватываемых</span><span class="sxs-lookup"><span data-stu-id="985cd-110">Scripting skills covered</span></span>

- <span data-ttu-id="985cd-111">Внешние вызовы API `fetch` ()</span><span class="sxs-lookup"><span data-stu-id="985cd-111">External API calls (`fetch`)</span></span>
- <span data-ttu-id="985cd-112">Размыв JSON</span><span class="sxs-lookup"><span data-stu-id="985cd-112">JSON parsing</span></span>
- <span data-ttu-id="985cd-113">Диаграммы</span><span class="sxs-lookup"><span data-stu-id="985cd-113">Charts</span></span>

## <a name="setup-instructions"></a><span data-ttu-id="985cd-114">Инструкции по настройке</span><span class="sxs-lookup"><span data-stu-id="985cd-114">Setup instructions</span></span>

1. <span data-ttu-id="985cd-115">Откройте книгу с помощью Excel в Интернете.</span><span class="sxs-lookup"><span data-stu-id="985cd-115">Open the workbook with Excel on the web.</span></span>

1. <span data-ttu-id="985cd-116">В **вкладке Automate** выберите **Все скрипты**.</span><span class="sxs-lookup"><span data-stu-id="985cd-116">Under the **Automate** tab, select **All Scripts**.</span></span>

1. <span data-ttu-id="985cd-117">В области **задач редактора** кода выберите **Новый скрипт** и вклеите следующий скрипт в редактор.</span><span class="sxs-lookup"><span data-stu-id="985cd-117">In the **Code Editor** task pane, select **New Script** and paste the following script into the editor.</span></span>

    ```TypeScript
    /**
     * Gets data from the National Oceanic and Atmospheric Administration's Tides and Currents database. 
     * That data is used to make a chart.
     */
    async function main(workbook: ExcelScript.Workbook): Promise<void> {
      // Get the current sheet.
      let currentSheet = workbook.getActiveWorksheet();
    
      // Create selection of parameters for the fetch URL.
      // More information on the NOAA APIs is found here: 
      // https://api.tidesandcurrents.noaa.gov/api/prod/
      const option = "water_level";
      const startDate = "20201225"; /* yyyymmdd date format */
      const endDate = "20201227";
      const station = "9447130"; /* Seattle */
    
      // Construct the URL for the fetch call.
      const strQuery = `https://api.tidesandcurrents.noaa.gov/api/prod/datagetter?product=${option}&begin_date=${startDate}&end_date=${endDate}&datum=MLLW&station=${station}&units=english&time_zone=gmt&application=NOS.COOPS.TAC.WL&format=json`;
    
      console.log(strQuery);
    
      // Resolve the Promises returned by the fetch operation.
      const response = await fetch(strQuery);
      const rawJson: string = await response.json();
    
      // Translate the raw JSON into a usable state.
      const stringifiedJson = JSON.stringify(rawJson);
    
      // Note that we're only taking the data part of the JSON and excluding the metadata.
      const noaaData: NOAAData[] = JSON.parse(stringifiedJson).data;
    
      // Create table headers and format them to stand out.
      let headers = [["Time", "Level"]];
      let headerRange = currentSheet.getRange("A1:B1");
      headerRange.setValues(headers);
      headerRange.getFormat().getFill().setColor("#4472C4");
      headerRange.getFormat().getFont().setColor("white");
    
      // Insert all the data in rows from JSON.
      let noaaDataCount = noaaData.length;
      let dataToEnter = [[], []]
      for (let i = 0; i < noaaDataCount; i++) {
        let currentDataPiece = noaaData[i];
        dataToEnter[i] = [currentDataPiece.t, currentDataPiece.v];
      }
    
      let dataRange = currentSheet.getRange("A2:B" + String(noaaDataCount + 1)); /* +1 to account for the title row */
      dataRange.setValues(dataToEnter);
    
      // Format the "Time" column for timestamps.
      dataRange.getColumn(0).setNumberFormatLocal("[$-en-US]mm/dd/yyyy hh:mm AM/PM;@");
    
      // Create and format a chart with the level data.
      let chart = currentSheet.addChart(ExcelScript.ChartType.xyscatterSmooth, dataRange);
      chart.getTitle().setText("Water Level - Seattle");
      chart.setTop(0);
      chart.setLeft(300);
      chart.setWidth(500);
      chart.setHeight(300);
      chart.getAxes().getValueAxis().setShowDisplayUnitLabel(false);
      chart.getAxes().getCategoryAxis().setTextOrientation(60);
      chart.getLegend().setVisible(false);
    
      // Add a comment with the data attribution.
      currentSheet.addComment(
        "A1",
        `This data was taken from the National Oceanic and Atmospheric Administration's Tides and Currents database on ${new Date(Date.now())}.`
      );
    
      /**
       * An interface to wrap the parts of the JSON we need.
       * These properties must match the names used in the JSON.
       */ 
      interface NOAAData {
        t: string; // Time
        v: number; // Level
      }
    }
    ```

1. <span data-ttu-id="985cd-118">Переименуй сценарий в **диаграмму уровня воды NOAA** и сохраните его.</span><span class="sxs-lookup"><span data-stu-id="985cd-118">Rename the script to **NOAA Water Level Chart** and save it.</span></span>

## <a name="running-the-script"></a><span data-ttu-id="985cd-119">Выполнение скрипта</span><span class="sxs-lookup"><span data-stu-id="985cd-119">Running the script</span></span>

<span data-ttu-id="985cd-120">На любом графике запустите сценарий **диаграммы уровня воды NOAA.**</span><span class="sxs-lookup"><span data-stu-id="985cd-120">On any worksheet, run the **NOAA Water Level Chart** script.</span></span> <span data-ttu-id="985cd-121">Сценарий извлекает данные уровня воды с 25 декабря 2020 г. по 27 декабря 2020 г.</span><span class="sxs-lookup"><span data-stu-id="985cd-121">The script fetches the water level data from December 25, 2020 to December 27, 2020.</span></span> <span data-ttu-id="985cd-122">Переменные в начале сценария можно изменить, чтобы использовать разные даты `const` или получать различные сведения о станциях.</span><span class="sxs-lookup"><span data-stu-id="985cd-122">The `const` variables at the beginning of the script can be changed to use different dates or get different station information.</span></span> <span data-ttu-id="985cd-123">API [CO-OPS for Data Retrieval](https://api.tidesandcurrents.noaa.gov/api/prod/) описывает, как получить все эти данные.</span><span class="sxs-lookup"><span data-stu-id="985cd-123">The [CO-OPS API For Data Retrieval](https://api.tidesandcurrents.noaa.gov/api/prod/) describes how to get all this data.</span></span>

### <a name="after-running-the-script"></a><span data-ttu-id="985cd-124">После запуска скрипта</span><span class="sxs-lookup"><span data-stu-id="985cd-124">After running the script</span></span>

:::image type="content" source="../../images/scenario-noaa-water-level-after.png" alt-text="В таблице после запуска скрипта показаны некоторые данные уровня воды и диаграмма.":::
