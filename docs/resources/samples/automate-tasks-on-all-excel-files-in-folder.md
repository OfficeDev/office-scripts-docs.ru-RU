---
title: Запустите сценарий для всех файлов Excel в папке
description: Узнайте, как запустить сценарий для всех файлов Excel в папке OneDrive для бизнеса.
ms.date: 03/31/2021
localization_priority: Normal
ms.openlocfilehash: a11876e8241a069a7c640bbcf2c36b4842d3bd90
ms.sourcegitcommit: 5d24e77df70aa2c1c982275d53213c2a9323ff86
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 04/02/2021
ms.locfileid: "51571488"
---
# <a name="run-a-script-on-all-excel-files-in-a-folder"></a>Запустите сценарий для всех файлов Excel в папке

Этот проект выполняет набор задач автоматизации для всех файлов, расположенных в папке OneDrive для бизнеса. Он также может использоваться в папке SharePoint.
Выполняет вычисления в файлах Excel, добавляет форматирование и [](https://support.microsoft.com/office/90701709-5dc1-41c7-aa48-b01d4a46e8c7) вставляет @mentions коллегу.

## <a name="sample-code-add-formatting-and-insert-comment"></a>Пример кода: добавление форматирования и вставки комментариев

Скачайте <a href="https://github.com/OfficeDev/office-scripts-docs/blob/master/docs/resources/samples/highlight-alert-excel-files.zip?raw=true"> файлhighlight-alert-excel-files.zip,</a>извлеките файлы в папку с названием **Sales,** используемую в этом примере, и попробуйте ее самостоятельно!

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

## <a name="training-video-run-a-script-on-all-excel-files-in-a-folder"></a>Обучающее видео. Запустите сценарий для всех файлов Excel в папке

[Просмотрите пошаговую видеозапись](https://youtu.be/xMg711o7k6w) запуска скрипта для всех файлов Excel в папке OneDrive для бизнеса или SharePoint.
