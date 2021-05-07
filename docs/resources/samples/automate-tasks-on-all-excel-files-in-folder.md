---
title: Запуск сценария для всех файлов Excel в папке
description: Узнайте, как запустить сценарий для всех Excel файлов в папке на OneDrive для бизнеса.
ms.date: 04/28/2021
localization_priority: Normal
ms.openlocfilehash: a6b869e2b346635e2b28fa7c6273c1a86a5bc5c5
ms.sourcegitcommit: f7a7aebfb687f2a35dbed07ed62ff352a114525a
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 05/06/2021
ms.locfileid: "52232629"
---
# <a name="run-a-script-on-all-excel-files-in-a-folder"></a>Запуск сценария для всех файлов Excel в папке

Этот проект выполняет набор задач автоматизации для всех файлов, расположенных в папке на OneDrive для бизнеса. Его также можно использовать в SharePoint папке.
Он выполняет вычисления Excel файлов, добавляет форматирование и вставляет комментарий, @mentions [коллеге.](https://support.microsoft.com/office/90701709-5dc1-41c7-aa48-b01d4a46e8c7)

Скачайте <a href="https://github.com/OfficeDev/office-scripts-docs/blob/master/docs/resources/samples/highlight-alert-excel-files.zip?raw=true"> файлhighlight-alert-excel-files.zip,</a>извлеките файлы в папку с названием **Sales,** используемую в этом примере, и попробуйте ее самостоятельно!

## <a name="sample-code-add-formatting-and-insert-comment"></a>Пример кода: добавление форматирования и вставки комментариев

Это сценарий, который выполняется в каждой отдельной книге.

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  let table1 = workbook.getTable("Table1");
  const rowCount = table1.getRowCount();
  if (rowCount === 0) {
    return;
  }
  workbook.getApplication().calculate(ExcelScript.CalculationType.full);

  const amountDueCol = table1.getColumnByName('Amount Due');
  const amountDueValues = amountDueCol.getRangeBetweenHeaderAndTotal().getValues();

  let highestValue = amountDueValues[0][0];
  let row = 0;
  for (let i = 1; i < amountDueValues.length; i++) {
    if (amountDueValues[i][0] > highestValue) {
      highestValue = amountDueValues[i][0];
      row = i;
    }
  }
  // Set fill color to FFFF00 for range in table Table1 cell in row 0 on column "Amount due".
  table1.getColumn("Amount due")
    .getRangeBetweenHeaderAndTotal()
    .getRow(row)
    .getFormat()
    .getFill()
    .setColor("FFFF00");
  let selectedSheet = workbook.getActiveWorksheet();
  // Insert comment at cell InvoiceAmounts!F2.
  workbook.addComment(table1.getColumn("Amount due")
    .getRangeBetweenHeaderAndTotal()
    .getRow(row), {
    mentions: [{
      email: "AdeleV@M365x904181.OnMicrosoft.com",
      id: 0,
      name: "Adele Vance"
    }],
    richContent: "<at id=\"0\">Adele Vance</at> Please review this amount"
  }, ExcelScript.ContentType.mention);
}
```

## <a name="power-automate-flow-run-the-script-on-every-workbook-in-the-folder"></a>Power Automate: запустите сценарий для каждой книги в папке

Этот поток запускает сценарий для каждой книги в папке "Продажи".

1. Создайте новый **поток мгновенных облаков.**
1. Выберите **вручную вызвать поток и** нажмите **кнопку Создать**.
1. Добавьте новый **шаг,** использующий **соединителю OneDrive для бизнеса** и файлы **List в действии папки.**

    :::image type="content" source="../../images/all-files-in-folder-sample-flow-1.png" alt-text="Завершенный OneDrive для бизнеса в Power Automate":::
1. Выберите папку "Продажи" с извлеченными книгами.
1. Чтобы убедиться, что выбраны только книги, выберите **новый** шаг, а затем выберите **Условие** и установите следующие значения:
    1. **Имя** (значение OneDrive файла)
    1. "заканчивается"
    1. xlsx.

    :::image type="content" source="../../images/all-files-in-folder-sample-flow-2.png" alt-text="Блок Power Automate, который применяет последующие действия к каждому файлу":::
1. В **филиале If Yes** **добавьте соединителю Excel Online (Бизнес)** с действием Сценарий запуска **(предварительного просмотра).** Используйте следующие значения для действия:
    1. **Расположение**: OneDrive для бизнеса
    1. **Библиотека документов**: OneDrive
    1. **Файл**: **Id** (OneDrive файла)
    1. **Сценарий:** имя сценария

    :::image type="content" source="../../images/all-files-in-folder-sample-flow-3.png" alt-text="Завершенный соедините Excel Online (Бизнес) в Power Automate":::
1. Сохраните поток и попробуйте его.

## <a name="training-video-run-a-script-on-all-excel-files-in-a-folder"></a>Обучающее видео: запустите сценарий для всех Excel файлов в папке

[Смотреть Sudhi Ramamurthy ходить через этот пример на YouTube](https://youtu.be/xMg711o7k6w).
