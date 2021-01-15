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
# <a name="office-scripts-sample-scenario-fetch-and-graph-water-level-data-from-noaa"></a>Пример сценария сценариев Office: извлечение и график данных на уровне ватерли от NOAA

В этом сценарии необходимо выровнеть уровень вехи на станции ["National Wateric and Seattle Administration" в Сиэтле.](https://tidesandcurrents.noaa.gov/stationhome.html?id=9447130) Внешние данные используются для заполнения таблицы и создания диаграммы.

Вы разработайте сценарий, использующий команду для запроса базы данных `fetch` [noAA — "Подмайки" и "Currents".](https://tidesandcurrents.noaa.gov/) Это позволит фиксировать уровень в течение заданного периода времени. Сведения будут возвращены в качестве JSON, поэтому часть скрипта преобразует их в значения диапазона. После того как данные посетят электронные таблицы, они будут использоваться для начертки диаграммы.

## <a name="scripting-skills-covered"></a>Навыки написания сценариев

- Внешние вызовы API ( `fetch` )
- Разбиение по JSON
- Диаграммы

## <a name="setup-instructions"></a>Инструкции по установке

1. Откройте книгу с помощью Excel в Интернете.

1. На **вкладке "Автоматизация"** выберите **"Все сценарии".**

1. В области **задач редактора** кода выберите **"Новый** сценарий" и в paste the following script into the editor.

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

1. Переименуем сценарий в **диаграмму уровня ватерли NOAA** и сохраните его.

## <a name="running-the-script"></a>Выполнение скрипта

На любом из них запустите сценарий диаграммы **на уровне ватерли NOAA.** Сценарий извлекает данные об уровне вехи с 25 декабря 2020 г. по 27 декабря 2020 г. Переменные в начале сценария можно изменить, чтобы использовать `const` разные даты или получить другую информацию о станции. API [CO-OPS для](https://api.tidesandcurrents.noaa.gov/api/prod/) и получения данных описывает, как получить все эти данные.

### <a name="after-running-the-script"></a>После запуска сценария

![На этом графике после запуска сценария показаны некоторые данные об уровне ватерли и диаграмма.](../../images/scenario-noaa-water-level-after.png)
