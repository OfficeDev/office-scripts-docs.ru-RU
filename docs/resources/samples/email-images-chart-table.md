---
title: Отправить по электронной почте изображения Excel и таблицы
description: Узнайте, как использовать Office скрипты и Power Automate для извлечения и отправки по электронной почте изображений Excel диаграммы и таблицы.
ms.date: 04/28/2021
localization_priority: Normal
ms.openlocfilehash: b49b6670562d117bb3dd6dcf894c54432bc5ceaa
ms.sourcegitcommit: f7a7aebfb687f2a35dbed07ed62ff352a114525a
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 05/06/2021
ms.locfileid: "52232594"
---
# <a name="use-office-scripts-and-power-automate-to-email-images-of-a-chart-and-table"></a>Использование Office и Power Automate для отправки изображений диаграммы и таблицы по электронной почте

В этом примере Office скрипты и Power Automate для создания диаграммы. Затем он передает по электронной почте изображения диаграммы и базовой таблицы.

## <a name="example-scenario"></a>Пример сценария

* Вычислять, чтобы получить последние результаты.
* Создание диаграммы.
* Получите изображения диаграммы и таблицы.
* Отправьте изображения по электронной почте Power Automate.

_Входные данные_

:::image type="content" source="../../images/input-data.png" alt-text="Таблица, показывающая таблицу входных данных":::

_Диаграмма вывода_

:::image type="content" source="../../images/chart-created.png" alt-text="Диаграмма столбцов, созданная с указанием суммы, за которую должен высмеять клиент":::

_Электронная почта, полученная Power Automate потока_

:::image type="content" source="../../images/email-received.png" alt-text="Сообщение, отправленное потоком, с указанием Excel, встроенного в тело":::

## <a name="solution"></a>Решение

Это решение состоит из двух частей:

1. [Сценарий Office для вычисления и извлечения Excel диаграммы и таблицы](#sample-code-calculate-and-extract-excel-chart-and-table)
1. Поток Power Automate для вызова скрипта и отправки результатов по электронной почте. Пример этого см. в примере [Create an automated workflow with Power Automate.](../../tutorials/excel-power-automate-returns.md#create-an-automated-workflow-with-power-automate)

## <a name="sample-code-calculate-and-extract-excel-chart-and-table"></a>Пример кода. Вычислять и извлекать Excel диаграмму и таблицу

Следующий сценарий вычисляет и извлекает Excel диаграмму и таблицу.

Скачайте пример файла <a href="email-chart-table.xlsx">email-chart-table.xlsx</a> и используйте его с помощью этого скрипта, чтобы попробовать его самостоятельно!

```TypeScript
function main(workbook: ExcelScript.Workbook): ReportImages {

  workbook.getApplication().calculate(ExcelScript.CalculationType.full);
  
  let sheet1 = workbook.getWorksheet("Sheet1");
  const table = workbook.getWorksheet('InvoiceAmounts').getTables()[0];
  const rows = table.getRange().getTexts();

  const selectColumns = rows.map((row) => {
    return [row[2], row[5]];
  });
  table.setShowTotals(true);
  selectColumns.splice(selectColumns.length-1, 1);
  console.log(selectColumns);

  workbook.getWorksheet('ChartSheet')?.delete();
  const chartSheet = workbook.addWorksheet('ChartSheet');
  const targetRange = updateRange(chartSheet, selectColumns);

  // Insert chart on sheet 'Sheet1'.
  let chart_2 = chartSheet.addChart(ExcelScript.ChartType.columnClustered, targetRange);
  chart_2.setPosition('D1');
  const chartImage = chart_2.getImage();
  const tableImage = table.getRange().getImage();
  return {
    chartImage,
    tableImage
  }
}

function updateRange(sheet: ExcelScript.Worksheet, data: string[][]): ExcelScript.Range {
  const targetRange = sheet.getRange('A1').getResizedRange(data.length-1, data[0].length-1);
  targetRange.setValues(data);
  return targetRange;
}

interface ReportImages {
  chartImage: string
  tableImage: string
}
```

## <a name="power-automate-flow-email-the-chart-and-table-images"></a>Power Automate потока: отправить по электронной почте изображения диаграммы и таблицы

Этот поток запускает сценарий и передает возвращаемые изображения по электронной почте.

1. Создайте новый **поток мгновенных облаков.**
1. Выберите **вручную вызвать поток и** нажмите **кнопку Создать**.
1. Добавьте новый **шаг,** использующий **соединителю Excel Online (Бизнес)** с действием **Запуска скрипта (предварительного просмотра).** Используйте следующие значения для действия:
    * **Расположение**: OneDrive для бизнеса
    * **Библиотека документов**: OneDrive
    * **Файл**: Ваша книга [(выбрана с помощью выбора файла)](../../testing/power-automate-troubleshooting.md#select-workbooks-with-the-file-browser-control)
    * **Сценарий:** имя сценария

    :::image type="content" source="../../images/email-chart-sample-flow-1.png" alt-text="Завершенный соедините Excel Online (Бизнес) в Power Automate":::
1. В этом примере Outlook как клиент электронной почты. Можно использовать любые соединители электронной почты Power Automate поддерживает, но остальные действия предполагают, что вы выбрали Outlook. Добавьте новый **шаг,** использующий **соединителю Office 365 Outlook** и действие Отправка и электронная почта **(V2).** Используйте следующие значения для действия:
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

    :::image type="content" source="../../images/email-chart-sample-flow-2.png" alt-text="Завершенный соедините Office 365 Outlook в Power Automate":::
1. Сохраните поток и попробуйте его.

## <a name="training-video-extract-and-email-images-of-chart-and-table"></a>Обучающее видео: извлечение и отправка изображений диаграммы и таблицы по электронной почте

[Смотреть Sudhi Ramamurthy ходить через этот пример на YouTube](https://youtu.be/152GJyqc-Kw).
