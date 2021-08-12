---
title: Отправить по электронной почте изображения Excel и таблицы
description: Узнайте, как использовать Office скрипты и Power Automate для извлечения и отправки по электронной почте изображений Excel диаграммы и таблицы.
ms.date: 06/29/2021
localization_priority: Normal
ms.openlocfilehash: 1fffd81426c8850cb605e2dbc0f9bf023a4a3692c12f8bd7c8dcc9ec45236ab8
ms.sourcegitcommit: 75f7ed8c2d23a104acc293f8ce29ea580b4fcdc5
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 08/11/2021
ms.locfileid: "57846743"
---
# <a name="use-office-scripts-and-power-automate-to-email-images-of-a-chart-and-table"></a>Использование Office и Power Automate для отправки изображений диаграммы и таблицы по электронной почте

В этом примере Office скрипты и Power Automate для создания диаграммы. Затем он передает по электронной почте изображения диаграммы и базовой таблицы.

## <a name="example-scenario"></a>Пример сценария

* Вычислять, чтобы получить последние результаты.
* Создание диаграммы.
* Получите изображения диаграммы и таблицы.
* Отправьте изображения по электронной почте Power Automate.

_Входные данные_

:::image type="content" source="../../images/input-data.png" alt-text="Таблица, показывающая таблицу входных данных.":::

_Диаграмма вывода_

:::image type="content" source="../../images/chart-created.png" alt-text="Диаграмма столбцов, созданная с указанием суммы, которая должна быть засвеяна клиентом.":::

_Электронная почта, полученная Power Automate потока_

:::image type="content" source="../../images/email-received.png" alt-text="Сообщение, отправленное потоком, с указанием Excel, встроенного в тело.":::

## <a name="solution"></a>Решение

Это решение состоит из двух частей:

1. [Сценарий Office для вычисления и извлечения Excel диаграммы и таблицы](#sample-code-calculate-and-extract-excel-chart-and-table)
1. Поток Power Automate для вызова скрипта и отправки результатов по электронной почте. Пример этого см. в примере [Create an automated workflow with Power Automate.](../../tutorials/excel-power-automate-returns.md#create-an-automated-workflow-with-power-automate)

## <a name="sample-excel-file"></a>Пример Excel файла

Скачайте <a href="email-chart-table.xlsx">email-chart-table.xlsx</a> для готовой к использованию книги. Добавьте следующий скрипт, чтобы попробовать пример самостоятельно!

## <a name="sample-code-calculate-and-extract-excel-chart-and-table"></a>Пример кода. Вычислять и извлекать Excel диаграмму и таблицу

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

## <a name="power-automate-flow-email-the-chart-and-table-images"></a>Power Automate потока: отправить по электронной почте изображения диаграммы и таблицы

Этот поток запускает сценарий и передает возвращаемые изображения по электронной почте.

1. Создайте новый **поток мгновенных облаков.**
1. Выберите **вручную вызвать поток и** выберите **Создать**.
1. Добавьте новый **шаг,** использующий **соединителю Excel Online (Бизнес)** с действием **сценария Run.** Используйте следующие значения для действия.
    * **Расположение**: OneDrive для бизнеса
    * **Библиотека документов**: OneDrive
    * **Файл**: Ваша книга [(выбрана с помощью выбора файла)](../../testing/power-automate-troubleshooting.md#select-workbooks-with-the-file-browser-control)
    * **Сценарий:** имя сценария

    :::image type="content" source="../../images/email-chart-sample-flow-1.png" alt-text="Завершенный соедините Excel Online (Бизнес) в Power Automate.":::
1. В этом примере Outlook как клиент электронной почты. Можно использовать любые соединители электронной почты Power Automate поддерживает, но остальные действия предполагают, что вы выбрали Outlook. Добавьте новый **шаг,** использующий **соединителю Office 365 Outlook** и действие Отправка и электронная почта **(V2).** Используйте следующие значения для действия.
    * **Чтобы:** ваша тестовая учетная запись электронной почты (или личная электронная почта)
    * **Тема:** Просмотрите отчетные данные
    * Для поля **Body** выберите "Представление кода" () и `</>` введите следующее:

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

    :::image type="content" source="../../images/email-chart-sample-flow-2.png" alt-text="Завершенный соедините Office 365 Outlook в Power Automate.":::
1. Сохраните поток и попробуйте его. Используйте **кнопку Test** на странице редактора потока или запустите поток через вкладку **Мои потоки.** Не забудьте разрешить доступ при запросе.

## <a name="training-video-extract-and-email-images-of-chart-and-table"></a>Обучающее видео: извлечение и отправка изображений диаграммы и таблицы по электронной почте

[Смотреть Sudhi Ramamurthy ходить через этот пример на YouTube](https://youtu.be/152GJyqc-Kw).
