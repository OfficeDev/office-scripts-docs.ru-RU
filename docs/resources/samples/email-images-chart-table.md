---
title: Отправьте изображения диаграммы и таблицы Excel электронной почте
description: Узнайте, как использовать Office и Power Automate для извлечения и электронной почты изображения диаграммы Excel таблицы.
ms.date: 05/06/2021
localization_priority: Normal
ms.openlocfilehash: 54b6b67a0f211f2dc6c881bab17ff23220619e6e
ms.sourcegitcommit: 4687693f02fc90a57ba30c461f35046e02e6f5fb
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 05/19/2021
ms.locfileid: "52545780"
---
# <a name="use-office-scripts-and-power-automate-to-email-images-of-a-chart-and-table"></a>Используйте Office и Power Automate по электронной почте изображения диаграммы и таблицы

Этот пример использует Office скрипты и Power Automate для создания диаграммы. Затем он отправляют изображения диаграммы и ее базовой таблицы.

## <a name="example-scenario"></a>Пример сценария

* Рассчитайте, чтобы получить последние результаты.
* Создайте диаграмму.
* Получите изображения диаграммы и таблицы.
* Отправьте изображения по электронной почте Power Automate.

_Входные данные_

:::image type="content" source="../../images/input-data.png" alt-text="Лист с указанием таблицы входных данных":::

_Диаграмма вывода_

:::image type="content" source="../../images/chart-created.png" alt-text="Диаграмма столбца создана, показывающая сумму, примеская заказчиком":::

_Электронная почта, полученная через Power Automate поток_

:::image type="content" source="../../images/email-received.png" alt-text="Письмо, отправленное потоком, показывающим Excel диаграмму, встроенную в тело":::

## <a name="solution"></a>Решение

Это решение состоит из двух частей:

1. [Сценарий Office для расчета и извлечения Excel диаграммы и таблицы](#sample-code-calculate-and-extract-excel-chart-and-table)
1. Поток Power Automate для вызова скрипта и электронной почты результатов. Пример того, как это сделать, смотрите Создание [автоматизированного рабочего процесса с Power Automate.](../../tutorials/excel-power-automate-returns.md#create-an-automated-workflow-with-power-automate)

## <a name="sample-code-calculate-and-extract-excel-chart-and-table"></a>Пример кода: Рассчитайте и извлекайте Excel диаграмму и таблицу

Следующий скрипт вычисляет и извлекает Excel и таблицу.

Скачать пример файла <a href="email-chart-table.xlsx">email-chart-table.xlsx</a> использовать его с этим скриптом, чтобы попробовать его самостоятельно!

```TypeScript
function main(workbook: ExcelScript.Workbook): ReportImages {
  // Recalculate the workbook to ensure all tables and charts are updated.
  workbook.getApplication().calculate(ExcelScript.CalculationType.full);
  
  // Get the data from the "InvoiceAmounts" table.
  let sheet1 = workbook.getWorksheet("Sheet1");
  const table = workbook.getWorksheet('InvoiceAmounts').getTables()[0];
  const rows = table.getRange().getTexts();

  // Get only the "Customer Name" and "Amount due" columns, then remove the "Total" row.
  const selectColumns = rows.map((row) => {
    return [row[2], row[5]];
  });
  table.setShowTotals(true);
  selectColumns.splice(selectColumns.length-1, 1);
  console.log(selectColumns);

  // Delete the "ChartSheet" worksheet if it's present, then recreate it.
  workbook.getWorksheet('ChartSheet')?.delete();
  const chartSheet = workbook.addWorksheet('ChartSheet');

  // Add the selected data to the new worksheet.
  const targetRange = chartSheet.getRange('A1').getResizedRange(selectColumns.length-1, selectColumns[0].length-1);
  targetRange.setValues(selectColumns);

  // Insert the chart on sheet 'ChartSheet' at cell "D1".
  let chart_2 = chartSheet.addChart(ExcelScript.ChartType.columnClustered, targetRange);
  chart_2.setPosition('D1');

  // Get images of the chart and table, then return them for a Power Automate flow.
  const chartImage = chart_2.getImage();
  const tableImage = table.getRange().getImage();
  return {chartImage, tableImage};
}

// The interface for table and chart images.
interface ReportImages {
  chartImage: string
  tableImage: string
}
```

## <a name="power-automate-flow-email-the-chart-and-table-images"></a>Power Automate поток: Электронная почта диаграммы и таблицы изображения

Этот поток запускает скрипт и отправляют возвращенные изображения по электронной почте.

1. Создайте новый **мгновенный поток облаков.**
1. Выберите **Вручную вызвать поток и** нажмите **Создать**.
1. Добавьте **новый шаг,** который использует **Excel Online (Бизнес)** с **действием сценария** Run. Используйте следующие значения для действия:
    * **Расположение**: OneDrive для бизнеса
    * **Библиотека документов**: OneDrive
    * **Файл**: Ваша трудовая [книжка (выбрана с помощью выбранного файла)](../../testing/power-automate-troubleshooting.md#select-workbooks-with-the-file-browser-control)
    * **Сценарий**: Ваше имя скрипта

    :::image type="content" source="../../images/email-chart-sample-flow-1.png" alt-text="Завершенный разъем Excel Online (Бизнес) в Power Automate":::
1. Этот образец использует Outlook в качестве почтового клиента. Вы можете использовать любой разъем электронной Power Automate поддерживает, но остальные шаги предполагают, что вы выбрали Outlook. Добавьте **новый шаг,** который использует **Office 365 Outlook** и отправить и **отправить (V2)** действий. Используйте следующие значения для действия:
    * **Для:** Ваш тестовый адрес электронной почты (или личная электронная почта)
    * **Тема**: Пожалуйста, просмотрите данные отчета
    * Для поля **тела** выберите "Code View" `</>` () и введите следующее:

    ```HTML
    <p>Please review the following report data:<br>
    <br>
    Chart:<br>
    <br>
    <img src="data:image/png;base64,@{outputs('Run_script')?['body/result/chartImage']}"/>
    <br>
    Data:<br>
    <br>
    <img src="data:image/png;base64,@{outputs('Run_script')?['body/result/tableImage']}"/>
    <br>
    </p>
    ```

    :::image type="content" source="../../images/email-chart-sample-flow-2.png" alt-text="Завершенный Office 365 Outlook в Power Automate":::
1. Сохранить поток и попробовать его.

## <a name="training-video-extract-and-email-images-of-chart-and-table"></a>Учебное видео: Выдержка и изображения электронной почты диаграммы и таблицы

[Смотреть Судхи Рамамурти ходить через этот образец на YouTube](https://youtu.be/152GJyqc-Kw).
