---
title: 'Office сценария сценариев: Graph данных на уровне воды из NOAA'
description: Пример, который извлекает данные JSON из базы данных NOAA и использует их для создания диаграммы.
ms.date: 03/08/2022
ms.localizationpriority: medium
ms.openlocfilehash: b4181edae7d8a46ae381ddfb1a2893b03faffd9b
ms.sourcegitcommit: dd01979d34b3499360d2f79a56f8a8f24f480eed
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 06/15/2022
ms.locfileid: "66088101"
---
# <a name="office-scripts-sample-scenario-fetch-and-graph-water-level-data-from-noaa"></a>Office сценария сценариев: получение и граф данных на уровне воды из NOAA

В этом сценарии необходимо отобразить уровень воды на станции Сиэтла национального океанического и [верхнего уровня](https://tidesandcurrents.noaa.gov/stationhome.html?id=9447130). Вы будете использовать внешние данные для заполнения электронной таблицы и создания диаграммы.

Вы разработаете сценарий, использующий команду `fetch` для запроса базы данных [NOAA Tides и Currents](https://tidesandcurrents.noaa.gov/). Это позволит получить уровень воды, записанный в заданный промежуток времени. Сведения будут возвращены в [формате JSON](https://www.w3schools.com/whatis/whatis_json.asp), поэтому часть скрипта преобразует их в значения диапазона. После того как данные будут внесены в электронную таблицу, они будут использоваться для создания диаграммы.

Дополнительные сведения о работе с JSON см. в статье "Использование JSON для передачи данных в Office [скрипты и из них"](../../develop/use-json.md).

## <a name="scripting-skills-covered"></a>Рассматриваются навыки навыков на написание скриптов

- Вызовы внешних API (`fetch`)
- Синтаксический анализ JSON
- Диаграммы

## <a name="setup-instructions"></a>Инструкции по настройке

1. Откройте книгу с помощью Excel в Интернете.

1. На **вкладке "Автоматизация** " выберите " **Новый скрипт** " и вставьте следующий скрипт в редактор.

    ```TypeScript
    /**
     * Gets data from the National Oceanic and Atmospheric Administration's Tides and Currents database. 
     * That data is used to make a chart.
     */
    async function main(workbook: ExcelScript.Workbook) {
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

1. Переименуйте скрипт в **диаграмму уровня воды NOAA** и сохраните его.

## <a name="running-the-script"></a>Выполнение скрипта

На любом листе запустите скрипт **диаграммы уровня воды NOAA** . Скрипт извлекает данные об уровне воды с 25 декабря 2020 г. по 27 декабря 2020 г. Переменные `const` в начале скрипта можно изменить для использования разных дат или получения разных сведений о станции. [API CO-OPS для](https://api.tidesandcurrents.noaa.gov/api/prod/) извлечения данных описывает, как получить все эти данные.

### <a name="after-running-the-script"></a>После выполнения скрипта

:::image type="content" source="../../images/scenario-noaa-water-level-after.png" alt-text="На листе после выполнения скрипта отображаются некоторые данные об уровне воды и диаграмма.":::
