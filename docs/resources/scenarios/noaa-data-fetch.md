---
title: 'Office Пример сценария: Graph данных уровня воды из NOAA'
description: Пример, который извлекает данные JSON из базы данных NOAA и использует их для создания диаграммы.
ms.date: 06/29/2021
localization_priority: Normal
ms.openlocfilehash: aca2a966e6a0588c597d6f01f52f2c73354de0c1
ms.sourcegitcommit: 211c157ca746e266eeb079f5fa1925a1e35ab702
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/07/2021
ms.locfileid: "53313766"
---
# <a name="office-scripts-sample-scenario-fetch-and-graph-water-level-data-from-noaa"></a>Office Сценарий примера сценариев: извлечение и график данных уровня воды из NOAA

В этом сценарии необходимо совместить уровень воды на станции [National Oceanic and Atmospheric Administration's Seattle.](https://tidesandcurrents.noaa.gov/stationhome.html?id=9447130) Внешние данные используются для заполнения таблицы и создания диаграммы.

Вы разработает сценарий, который использует команду для запроса базы `fetch` [данных NOAA Tides и Currents.](https://tidesandcurrents.noaa.gov/) Это позволит получить уровень воды, записанный через заданный промежуток времени. Сведения будут возвращены в качестве JSON, поэтому часть сценария будет переводить их в значения диапазона. После того как данные будут в таблице, они будут использоваться для сделайте диаграмму.

## <a name="scripting-skills-covered"></a>Навыки скриптов, охватываемых

- Внешние вызовы API `fetch` ()
- Размыв JSON
- Диаграммы

## <a name="setup-instructions"></a>Инструкции по настройке

1. Откройте книгу с помощью Excel в Интернете.

1. В **вкладке Automate** выберите **Новый скрипт** и вклеите следующий скрипт в редактор.

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

1. Переименуй сценарий в **диаграмму уровня воды NOAA** и сохраните его.

## <a name="running-the-script"></a>Выполнение скрипта

На любом графике запустите сценарий **диаграммы уровня воды NOAA.** Сценарий извлекает данные уровня воды с 25 декабря 2020 г. по 27 декабря 2020 г. Переменные в начале сценария можно изменить, чтобы использовать разные даты `const` или получать различные сведения о станциях. API [CO-OPS for Data Retrieval](https://api.tidesandcurrents.noaa.gov/api/prod/) описывает, как получить все эти данные.

### <a name="after-running-the-script"></a>После запуска скрипта

:::image type="content" source="../../images/scenario-noaa-water-level-after.png" alt-text="В таблице после запуска скрипта показаны некоторые данные уровня воды и диаграмма.":::
