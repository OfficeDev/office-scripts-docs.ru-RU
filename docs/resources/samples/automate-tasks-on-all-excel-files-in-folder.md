---
title: Запуск сценария для всех файлов Excel в папке
description: Узнайте, как запустить скрипт на всех Excel файлов в папке на OneDrive для бизнеса.
ms.date: 05/03/2021
localization_priority: Normal
ms.openlocfilehash: fb9a4deb01b52ef031cb1ba3400bd6f10de9d9f5
ms.sourcegitcommit: 4687693f02fc90a57ba30c461f35046e02e6f5fb
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 05/19/2021
ms.locfileid: "52545795"
---
# <a name="run-a-script-on-all-excel-files-in-a-folder"></a>Запуск сценария для всех файлов Excel в папке

Этот проект выполняет набор задач автоматизации на всех файлах, расположенных в папке на OneDrive для бизнеса. Он также может быть использован на SharePoint папке.
Он выполняет расчеты на Excel файлов, добавляет форматирование и вставляет комментарий, [который @mentions](https://support.microsoft.com/office/90701709-5dc1-41c7-aa48-b01d4a46e8c7) коллеге.

Скачать файл <a href="https://github.com/OfficeDev/office-scripts-docs/blob/master/docs/resources/samples/highlight-alert-excel-files.zip?raw=true">highlight-alert-excel-files.zip,</a>извлечь файлы в папку под **названием Продажи,** используемые в этом образце, и попробовать его самостоятельно!

## <a name="sample-code-add-formatting-and-insert-comment"></a>Пример кода: Добавить форматирование и вставить комментарий

Это скрипт, который работает на каждой отдельной рабочей книге.

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Get the table named "Table1" in the workbook.
  let table1 = workbook.getTable("Table1");

  // If the table is empty, end the script.
  const rowCount = table1.getRowCount();
  if (rowCount === 0) {
    return;
  }

  // Force the workbook to be completely recalculated.
  workbook.getApplication().calculate(ExcelScript.CalculationType.full);

  // Get the "Amount Due" column from the table.
  const amountDueColumn = table1.getColumnByName('Amount Due');
  const amountDueValues = amountDueColumn.getRangeBetweenHeaderAndTotal().getValues();

  // Find the highest amount that's due.
  let highestValue = amountDueValues[0][0];
  let row = 0;
  for (let i = 1; i < amountDueValues.length; i++) {
    if (amountDueValues[i][0] > highestValue) {
      highestValue = amountDueValues[i][0];
      row = i;
    }
  }

  let highestAmountDue = table1.getColumn("Amount due").getRangeBetweenHeaderAndTotal().getRow(row);

  // Set the fill color to yellow for the cell with the highest value in the "Amount Due" column.
  highestAmountDue
    .getFormat()
    .getFill()
    .setColor("FFFF00");

  // Insert an @mention comment in the cell.
  workbook.addComment(highestAmountDue, {
    mentions: [{
      email: "AdeleV@M365x904181.OnMicrosoft.com",
      id: 0,
      name: "Adele Vance"
    }],
    richContent: "<at id=\"0\">Adele Vance</at> Please review this amount"
  }, ExcelScript.ContentType.mention);
}
```

## <a name="power-automate-flow-run-the-script-on-every-workbook-in-the-folder"></a>Power Automate поток: Запустите скрипт на каждой рабочей книге в папке

Этот поток запускает скрипт на каждой рабочей книге в папке "Продажи".

1. Создайте новый **мгновенный поток облаков.**
1. Выберите **Вручную вызвать поток и** нажмите **Создать**.
1. Добавьте **новый шаг,** который использует **OneDrive для бизнеса** и файлы списка **в действии папки.**

    :::image type="content" source="../../images/all-files-in-folder-sample-flow-1.png" alt-text="Завершенный OneDrive для бизнеса в Power Automate":::
1. Выберите папку "Продажи" с извлеченными трудовыми книжками.
1. Чтобы обеспечить выбор только трудовых книжек, **выберите новый шаг,** затем **выберите Условие** и установите следующие значения:
    1. **Имя** (OneDrive имени файла)
    1. "заканчивается"
    1. "xlsx".

    :::image type="content" source="../../images/all-files-in-folder-sample-flow-2.png" alt-text="Блок Power Automate, который применяет последующие действия к каждому файлу":::
1. Под **ветвью If yes** добавьте **разъем Excel Online (Business)** с **действием сценария Run.** Используйте следующие значения для действия:
    1. **Расположение**: OneDrive для бизнеса
    1. **Библиотека документов**: OneDrive
    1. **Файл**: **Id** (OneDrive значение идентификатора файла)
    1. **Сценарий**: Ваше имя скрипта

    :::image type="content" source="../../images/all-files-in-folder-sample-flow-3.png" alt-text="Завершенный разъем Excel Online (Бизнес) в Power Automate":::
1. Сохранить поток и попробовать его.

## <a name="training-video-run-a-script-on-all-excel-files-in-a-folder"></a>Учебное видео: Запустите сценарий на всех Excel файлов в папке

[Смотреть Судхи Рамамурти ходить через этот образец на YouTube](https://youtu.be/xMg711o7k6w).
