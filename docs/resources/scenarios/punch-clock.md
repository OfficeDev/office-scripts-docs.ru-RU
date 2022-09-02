---
title: 'Пример сценария сценариев Office: кнопка "Часы"'
description: В этом примере добавляется кнопка "Часы" и пользователь может выполнять тактовую синхронизацию с использованием текущего времени.
ms.date: 04/26/2022
ms.localizationpriority: medium
ms.openlocfilehash: ac128a33b653506b6168bd4acfe1713bf6d26759
ms.sourcegitcommit: a6504f8b0d6b717457c6e0b5306c35ad3900914e
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/02/2022
ms.locfileid: "67572684"
---
# <a name="office-scripts-sample-scenario-punch-clock-button"></a>Пример сценария сценариев Office: кнопка "Часы"

Идея сценария и сценарий, используемые в этом примере, были предоставлены членом сообщества сценариев Office [Брайаном Gonzalez](https://github.com/b-gonzalez).

В этом сценарии вы создадите лист времени для сотрудника, который позволяет ему записывать время начала и окончания с помощью нажатия [кнопки](../../develop/script-buttons.md). В зависимости от того, что было записано ранее, нажатие кнопки будет начинаться с их дня (часы в) или завершить свой день (время ожидания). Этот пример работает как для Excel в Интернете, так и для Windows.

:::image type="content" source="../../images/punch-clock-sample-3.png" alt-text="Таблица с тремя столбцами (&quot;Часы в&quot;, &quot;Время ожидания&quot; и &quot;Длительность&quot;) и кнопкой &quot;Часы по часам&quot; в книге.":::

## <a name="setup-instructions"></a>Инструкции по настройке

1. [ Скачайтеpunch-clock-sample.xlsx](punch-clock-sample.xlsx) в OneDrive.

    :::image type="content" source="../../images/punch-clock-sample-1.png" alt-text="Таблица с тремя столбцами: &quot;Clock In&quot;, &quot;Clock Out&quot; и &quot;Duration&quot;.":::

1. Откройте книгу в Excel в Интернете.

1. На **вкладке "Автоматизация** " выберите " **Новый скрипт** " и вставьте следующий скрипт в редактор.

    ```typescript
    /**
     * This script records either the start or end time of a shift, 
     * depending on what is filled out in the table. 
     * It is intended to be used with a Script Button.
     */
    function main(workbook: ExcelScript.Workbook) {
      // Get the first table in the timesheet.
      const timeSheet = workbook.getWorksheet("MyTimeSheet");
      const timeTable = timeSheet.getTables()[0];
    
      // Get the appropriate table columns.
      const clockInColumn = timeTable.getColumnByName("Clock In");
      const clockOutColumn = timeTable.getColumnByName("Clock Out");
      const durationColumn = timeTable.getColumnByName("Duration");
    
      // Get the last rows for the Clock In and Clock Out columns.
      let clockInLastRow = clockInColumn.getRangeBetweenHeaderAndTotal().getLastRow();
      let clockOutLastRow = clockOutColumn.getRangeBetweenHeaderAndTotal().getLastRow();
    
      // Get the current date to use as the start or end time.
      let date: Date = new Date();
    
      // Add the current time to a column based on the state of the table.
      if (clockInLastRow.getValue() as string === "") {
        // If the Clock In column has an empty value in the table, add a start time.
        clockInLastRow.setValue(date.toLocaleString());
      } else if (clockOutLastRow.getValue() as string === "") {
        // If the Clock Out column has an empty value in the table, 
        // add an end time and calculate the shift duration.
        clockOutLastRow.setValue(date.toLocaleString());
        const clockInTime = new Date(clockInLastRow.getValue() as string);
        const clockOutTime  = new Date(clockOutLastRow.getValue() as string);
        const clockDuration = Math.abs((clockOutTime.getTime() - clockInTime.getTime()));
    
        let durationString = getDurationMessage(clockDuration);
        durationColumn.getRangeBetweenHeaderAndTotal().getLastRow().setValue(durationString);
      } else {
        // If both columns are full, add a new row, then add a start time.
        timeTable.addRow()
        clockInLastRow.getOffsetRange(1, 0).setValue(date.toLocaleString());
      }
    }
    
    /**
     * A function to write a time duration as a string.
     */
    function getDurationMessage(delta: number) {
      // Adapted from here:
      // https://stackoverflow.com/questions/13903897/javascript-return-number-of-days-hours-minutes-seconds-between-two-dates
    
      delta = delta / 1000;
      let durationString = "";
    
      let days = Math.floor(delta / 86400);
      delta -= days * 86400;
    
      let hours = Math.floor(delta / 3600) % 24;
      delta -= hours * 3600;
    
      let minutes = Math.floor(delta / 60) % 60;
    
      if (days >= 1) {
        durationString += days;
        durationString += (days > 1 ? " days" : " day");
    
        if (hours >= 1 && minutes >= 1) {
          durationString += ", ";
        }
        else if (hours >= 1 || minutes > 1) {
          durationString += " and ";
        }
      }
    
      if (hours >= 1) {
        durationString += hours;
        durationString += (hours > 1 ? " hours" : " hour");
        if (minutes >= 1) {
          durationString += " and ";
        }
      }
    
      if (minutes >= 1) {
        durationString += minutes;
        durationString += (minutes > 1 ? " minutes" : " minute");
      }
    
      return durationString;
    }
    ```

1. Переименуйте скрипт в "Часы".

1. Сохраните скрипт.

1. В книге выберите ячейку **E2**.

1. Добавьте кнопку сценария. Перейдите в меню **"Дополнительные параметры" (...)** **на странице** сведений о скрипте и нажмите **кнопку "Добавить"**.

    :::image type="content" source="../../images/punch-clock-sample-2.png" alt-text="Меню &quot;Дополнительные параметры&quot; и кнопка &quot;Добавить кнопку&quot;.":::

1. Сохраните книгу.

## <a name="run-the-script"></a>Запустите сценарий

Нажмите **кнопку "Часы"** , чтобы запустить сценарий. Оно регистрирует текущее время в разделе "В часах" или "Время ожидания" в зависимости от того, что было введено ранее.

:::image type="content" source="../../images/punch-clock-sample-3.png" alt-text="Таблица и кнопка &quot;Часы&quot; в книге.":::

> [!NOTE]
> Длительность записывается, только если она превышает минуту. Вручную измените время "В часах", чтобы проверить большее время.
