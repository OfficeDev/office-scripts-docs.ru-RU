---
title: 'Пример сценария сценариев Office: данные на уровне ватерли Graph из NOAA'
description: Пример, который получает данные JSON из базы данных NOAA и использует их для создания диаграммы.
ms.date: 01/11/2021
localization_priority: Normal
ms.openlocfilehash: 5b0b4e3675cbe053368f63123d819f0dab626e60
ms.sourcegitcommit: 7580dcb8f2f97974c2a9cce25ea30d6526730e28
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 01/14/2021
ms.locfileid: "49867879"
---
# <a name="office-scripts-sample-scenario-fetch-and-graph-water-level-data-from-noaa"></a><span data-ttu-id="31082-103">Пример сценария сценариев Office: извлечение и график данных на уровне ватерли от NOAA</span><span class="sxs-lookup"><span data-stu-id="31082-103">Office Scripts sample scenario: Fetch and graph water-level data from NOAA</span></span>

<span data-ttu-id="31082-104">В этом сценарии необходимо выровнеть уровень вехи на станции ["National Wateric and Seattle Administration" в Сиэтле.](https://tidesandcurrents.noaa.gov/stationhome.html?id=9447130)</span><span class="sxs-lookup"><span data-stu-id="31082-104">In this scenario, you need to plot the water level at the [National Oceanic and Atmospheric Administration's Seattle station](https://tidesandcurrents.noaa.gov/stationhome.html?id=9447130).</span></span> <span data-ttu-id="31082-105">Внешние данные используются для заполнения таблицы и создания диаграммы.</span><span class="sxs-lookup"><span data-stu-id="31082-105">You'll use external data to populate a spreadsheet and create a chart.</span></span>

<span data-ttu-id="31082-106">Вы разработайте сценарий, использующий команду для запроса базы данных `fetch` [noAA — "Подмайки" и "Currents".](https://tidesandcurrents.noaa.gov/)</span><span class="sxs-lookup"><span data-stu-id="31082-106">You'll develop a script that uses the `fetch` command to query the [NOAA Tides and Currents database](https://tidesandcurrents.noaa.gov/).</span></span> <span data-ttu-id="31082-107">Это позволит фиксировать уровень в течение заданного периода времени.</span><span class="sxs-lookup"><span data-stu-id="31082-107">That will get the water level recorded across a given time span.</span></span> <span data-ttu-id="31082-108">Сведения будут возвращены в качестве JSON, поэтому часть скрипта преобразует их в значения диапазона.</span><span class="sxs-lookup"><span data-stu-id="31082-108">The information will be returned as JSON, so part of the script will translate that into range values.</span></span> <span data-ttu-id="31082-109">После того как данные посетят электронные таблицы, они будут использоваться для начертки диаграммы.</span><span class="sxs-lookup"><span data-stu-id="31082-109">Once the data is in the spreadsheet, it will be used to make a chart.</span></span>

## <a name="scripting-skills-covered"></a><span data-ttu-id="31082-110">Навыки написания сценариев</span><span class="sxs-lookup"><span data-stu-id="31082-110">Scripting skills covered</span></span>

- <span data-ttu-id="31082-111">Внешние вызовы API ( `fetch` )</span><span class="sxs-lookup"><span data-stu-id="31082-111">External API calls (`fetch`)</span></span>
- <span data-ttu-id="31082-112">Разбиение по JSON</span><span class="sxs-lookup"><span data-stu-id="31082-112">JSON parsing</span></span>
- <span data-ttu-id="31082-113">Диаграммы</span><span class="sxs-lookup"><span data-stu-id="31082-113">Charts</span></span>

## <a name="setup-instructions"></a><span data-ttu-id="31082-114">Инструкции по установке</span><span class="sxs-lookup"><span data-stu-id="31082-114">Setup instructions</span></span>

1. <span data-ttu-id="31082-115">Откройте книгу с помощью Excel в Интернете.</span><span class="sxs-lookup"><span data-stu-id="31082-115">Open the workbook with Excel on the web.</span></span>

1. <span data-ttu-id="31082-116">На **вкладке "Автоматизация"** выберите **"Все сценарии".**</span><span class="sxs-lookup"><span data-stu-id="31082-116">Under the **Automate** tab, select **All Scripts**.</span></span>

1. <span data-ttu-id="31082-117">В области **задач редактора** кода выберите **"Новый** сценарий" и в paste the following script into the editor.</span><span class="sxs-lookup"><span data-stu-id="31082-117">In the **Code Editor** task pane, select **New Script** and paste the following script into the editor.</span></span>

    ```typescript
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
      const rawJson = await response.json();
    
      // Translate the raw JSON into a usable state.
      const stringifiedJson = JSON.stringify(rawJson);
      const noaaData = JSON.parse(stringifiedJson);
    
      // Create table headers and format them to stand out.
      let headers = [["Time", "Level"]];
      let headerRange = currentSheet.getRange("A1:B1");
      headerRange.setValues(headers);
      headerRange.getFormat().getFill().setColor("#4472C4");
      headerRange.getFormat().getFont().setColor("white");
    
      // Insert all the data in rows from JSON.
      let noaaDataCount = noaaData.data.length;
      let dataToEnter = [[], []]
      for (let i = 0; i < noaaDataCount; i++) {
        let currentDataPiece = noaaData.data[i];
        dataToEnter[i] = [currentDataPiece.t, currentDataPiece.v];
      }
    
      let dataRange = currentSheet.getRange("A2:B" + String(noaaDataCount + 1)); /* +1 to account for the title row */
      dataRange.setValues(dataToEnter);
      
      // Format the "Time" column for timestamps.
      dataRange.getColumn(0).setNumberFormatLocal("[$-en-US]mm/dd/yyyy hh:mm AM/PM;@");
    
      // Create and format a chart with the level data.
      let chart = currentSheet.addChart(ExcelScript.ChartType.xyscatterSmooth,dataRange);
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
    }
    ```

1. <span data-ttu-id="31082-118">Переименуем сценарий в **диаграмму уровня ватерли NOAA** и сохраните его.</span><span class="sxs-lookup"><span data-stu-id="31082-118">Rename the script to **NOAA Water Level Chart** and save it.</span></span>

## <a name="running-the-script"></a><span data-ttu-id="31082-119">Выполнение скрипта</span><span class="sxs-lookup"><span data-stu-id="31082-119">Running the script</span></span>

<span data-ttu-id="31082-120">На любом из них запустите сценарий диаграммы **на уровне ватерли NOAA.**</span><span class="sxs-lookup"><span data-stu-id="31082-120">On any worksheet, run the **NOAA Water Level Chart** script.</span></span> <span data-ttu-id="31082-121">Сценарий извлекает данные об уровне вехи с 25 декабря 2020 г. по 27 декабря 2020 г.</span><span class="sxs-lookup"><span data-stu-id="31082-121">The script fetches the water level data from December 25, 2020 to December 27, 2020.</span></span> <span data-ttu-id="31082-122">Переменные в начале сценария можно изменить, чтобы использовать `const` разные даты или получить другую информацию о станции.</span><span class="sxs-lookup"><span data-stu-id="31082-122">The `const` variables at the beginning of the script can be changed to use different dates or get different station information.</span></span> <span data-ttu-id="31082-123">API [CO-OPS для](https://api.tidesandcurrents.noaa.gov/api/prod/) и получения данных описывает, как получить все эти данные.</span><span class="sxs-lookup"><span data-stu-id="31082-123">The [CO-OPS API For Data Retrieval](https://api.tidesandcurrents.noaa.gov/api/prod/) describes how to get all this data.</span></span>

### <a name="after-running-the-script"></a><span data-ttu-id="31082-124">После запуска сценария</span><span class="sxs-lookup"><span data-stu-id="31082-124">After running the script</span></span>

![На этом графике после запуска сценария показаны некоторые данные об уровне ватерли и диаграмма.](../../images/scenario-noaa-water-level-after.png)
