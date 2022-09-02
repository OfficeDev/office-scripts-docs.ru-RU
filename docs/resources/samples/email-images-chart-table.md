---
title: Email изображения диаграммы и таблицы Excel
description: Узнайте, как использовать сценарии Office и Power Automate для извлечения и отправки изображений диаграммы и таблицы Excel по электронной почте.
ms.date: 06/29/2021
ms.localizationpriority: medium
ms.openlocfilehash: dbf9135723a735321c99991d94f4b4387d800702
ms.sourcegitcommit: a6504f8b0d6b717457c6e0b5306c35ad3900914e
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/02/2022
ms.locfileid: "67572467"
---
# <a name="use-office-scripts-and-power-automate-to-email-images-of-a-chart-and-table"></a>Использование сценариев Office и Power Automate для отправки изображений диаграммы и таблицы по электронной почте

В этом примере для создания диаграммы используются сценарии Office и Power Automate. Затем он будет отправлять изображения диаграммы и ее базовой таблицы по электронной почте.

## <a name="example-scenario"></a>Пример сценария

* Вычисление для получения последних результатов.
* Создание диаграммы.
* Получение изображений диаграмм и таблиц.
* Email изображения с помощью Power Automate.

_Входные данные_

:::image type="content" source="../../images/input-data.png" alt-text="Лист с таблицей входных данных.":::

_Выходная диаграмма_

:::image type="content" source="../../images/chart-created.png" alt-text="Гистограмма, созданная с указанием суммы, выполнившегося клиентом.":::

_Email, полученные с помощью потока Power Automate_

:::image type="content" source="../../images/email-received.png" alt-text="Сообщение электронной почты, отправленное потоком, отображающее диаграмму Excel, внедренную в текст.":::

## <a name="solution"></a>Решение

Это решение состоит из двух частей:

1. [Скрипт Office для вычисления и извлечения диаграммы и таблицы Excel](#sample-code-calculate-and-extract-excel-chart-and-table)
1. Поток Power Automate для вызова скрипта и отправки результатов по электронной почте. Пример того, как это сделать, см. в статье "Создание автоматизированного рабочего процесса [с помощью Power Automate"](../../tutorials/excel-power-automate-returns.md#create-an-automated-workflow-with-power-automate).

## <a name="sample-excel-file"></a>Пример файла Excel

[ Скачайтеemail-chart-table.xlsx](email-chart-table.xlsx) для готовой к использованию книги. Добавьте следующий скрипт, чтобы попробовать пример самостоятельно!

## <a name="sample-code-calculate-and-extract-excel-chart-and-table"></a>Пример кода: вычисление и извлечение диаграммы и таблицы Excel

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

## <a name="power-automate-flow-email-the-chart-and-table-images"></a>Поток Power Automate: Email изображений диаграмм и таблиц

Этот поток запускает сценарий и по электронной почте возвращает возвращаемые изображения.

1. Создайте новый **мгновенный облачный поток**.
1. Выберите **"Вручную активировать поток" и** нажмите кнопку **"Создать"**.
1. Добавьте новый **шаг, использующий** соединитель **Excel Online (business)** с действием **запуска скрипта** . Используйте следующие значения для действия.
    * **Расположение**: OneDrive для бизнеса
    * **Библиотека документов**: OneDrive
    * **Файл**: книга ([выбрана с помощью выбора файла](../../testing/power-automate-troubleshooting.md#select-workbooks-with-the-file-browser-control))
    * **Сценарий**: имя скрипта

    :::image type="content" source="../../images/email-chart-sample-flow-1.png" alt-text="Завершенный соединитель Excel Online (business) в Power Automate.":::
1. В этом примере в качестве почтового клиента используется Outlook. Вы можете использовать любой соединитель электронной почты, поддерживаемый Power Automate, но в остальных шагах предполагается, что вы выбрали Outlook. Добавьте новый **шаг,** использующий **соединитель Office 365 Outlook** и действие **отправки и** отправки электронной почты (V2). Используйте следующие значения для действия.
    * **To**: Ваша тестовая учетная запись электронной почты (или личная электронная почта)
    * **Тема**. Просмотрите данные отчета
    * В поле **"Текст** " выберите "Представление кода" (`</>`) и введите следующее:

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

    :::image type="content" source="../../images/email-chart-sample-flow-2.png" alt-text="Завершенный Office 365 Outlook в Power Automate.":::
1. Сохраните поток и попробуйте его. Нажмите **кнопку "** Тест" на странице редактора потоков или запустите поток на **вкладке "Мои потоки** ". Не забудьте разрешить доступ при появлении запроса.

## <a name="training-video-extract-and-email-images-of-chart-and-table"></a>Обучающее видео: извлечение и отправка изображений диаграммы и таблицы по электронной почте

[Просмотрите этот пример на YouTube](https://youtu.be/152GJyqc-Kw), чтобы просмотреть этот пример.
