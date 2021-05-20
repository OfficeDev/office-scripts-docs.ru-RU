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
# <a name="run-a-script-on-all-excel-files-in-a-folder"></a><span data-ttu-id="3663e-103">Запуск сценария для всех файлов Excel в папке</span><span class="sxs-lookup"><span data-stu-id="3663e-103">Run a script on all Excel files in a folder</span></span>

<span data-ttu-id="3663e-104">Этот проект выполняет набор задач автоматизации на всех файлах, расположенных в папке на OneDrive для бизнеса.</span><span class="sxs-lookup"><span data-stu-id="3663e-104">This project performs a set of automation tasks on all files situated in a folder on OneDrive for Business.</span></span> <span data-ttu-id="3663e-105">Он также может быть использован на SharePoint папке.</span><span class="sxs-lookup"><span data-stu-id="3663e-105">It could also be used on a SharePoint folder.</span></span>
<span data-ttu-id="3663e-106">Он выполняет расчеты на Excel файлов, добавляет форматирование и вставляет комментарий, [который @mentions](https://support.microsoft.com/office/90701709-5dc1-41c7-aa48-b01d4a46e8c7) коллеге.</span><span class="sxs-lookup"><span data-stu-id="3663e-106">It performs calculations on the Excel files, adds formatting, and inserts a comment that [@mentions](https://support.microsoft.com/office/90701709-5dc1-41c7-aa48-b01d4a46e8c7) a colleague.</span></span>

<span data-ttu-id="3663e-107">Скачать файл <a href="https://github.com/OfficeDev/office-scripts-docs/blob/master/docs/resources/samples/highlight-alert-excel-files.zip?raw=true">highlight-alert-excel-files.zip,</a>извлечь файлы в папку под **названием Продажи,** используемые в этом образце, и попробовать его самостоятельно!</span><span class="sxs-lookup"><span data-stu-id="3663e-107">Download the file <a href="https://github.com/OfficeDev/office-scripts-docs/blob/master/docs/resources/samples/highlight-alert-excel-files.zip?raw=true">highlight-alert-excel-files.zip</a>, extract the files to a folder titled **Sales** used in this sample, and try it out yourself!</span></span>

## <a name="sample-code-add-formatting-and-insert-comment"></a><span data-ttu-id="3663e-108">Пример кода: Добавить форматирование и вставить комментарий</span><span class="sxs-lookup"><span data-stu-id="3663e-108">Sample code: Add formatting and insert comment</span></span>

<span data-ttu-id="3663e-109">Это скрипт, который работает на каждой отдельной рабочей книге.</span><span class="sxs-lookup"><span data-stu-id="3663e-109">This is the script that runs on each individual workbook.</span></span>

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

## <a name="power-automate-flow-run-the-script-on-every-workbook-in-the-folder"></a><span data-ttu-id="3663e-110">Power Automate поток: Запустите скрипт на каждой рабочей книге в папке</span><span class="sxs-lookup"><span data-stu-id="3663e-110">Power Automate flow: Run the script on every workbook in the folder</span></span>

<span data-ttu-id="3663e-111">Этот поток запускает скрипт на каждой рабочей книге в папке "Продажи".</span><span class="sxs-lookup"><span data-stu-id="3663e-111">This flow runs the script on every workbook in the "Sales" folder.</span></span>

1. <span data-ttu-id="3663e-112">Создайте новый **мгновенный поток облаков.**</span><span class="sxs-lookup"><span data-stu-id="3663e-112">Create a new **Instant cloud flow**.</span></span>
1. <span data-ttu-id="3663e-113">Выберите **Вручную вызвать поток и** нажмите **Создать**.</span><span class="sxs-lookup"><span data-stu-id="3663e-113">Select **Manually trigger a flow** and press **Create**.</span></span>
1. <span data-ttu-id="3663e-114">Добавьте **новый шаг,** который использует **OneDrive для бизнеса** и файлы списка **в действии папки.**</span><span class="sxs-lookup"><span data-stu-id="3663e-114">Add a **New step** that uses the **OneDrive for Business** connector and the **List files in folder** action.</span></span>

    :::image type="content" source="../../images/all-files-in-folder-sample-flow-1.png" alt-text="Завершенный OneDrive для бизнеса в Power Automate":::
1. <span data-ttu-id="3663e-116">Выберите папку "Продажи" с извлеченными трудовыми книжками.</span><span class="sxs-lookup"><span data-stu-id="3663e-116">Select the "Sales" folder with the extracted workbooks.</span></span>
1. <span data-ttu-id="3663e-117">Чтобы обеспечить выбор только трудовых книжек, **выберите новый шаг,** затем **выберите Условие** и установите следующие значения:</span><span class="sxs-lookup"><span data-stu-id="3663e-117">To ensure only workbooks are selected, choose **New step**, then select **Condition** and set the following values:</span></span>
    1. <span data-ttu-id="3663e-118">**Имя** (OneDrive имени файла)</span><span class="sxs-lookup"><span data-stu-id="3663e-118">**Name** (the OneDrive file name value)</span></span>
    1. <span data-ttu-id="3663e-119">"заканчивается"</span><span class="sxs-lookup"><span data-stu-id="3663e-119">"ends with"</span></span>
    1. <span data-ttu-id="3663e-120">"xlsx".</span><span class="sxs-lookup"><span data-stu-id="3663e-120">"xlsx".</span></span>

    :::image type="content" source="../../images/all-files-in-folder-sample-flow-2.png" alt-text="Блок Power Automate, который применяет последующие действия к каждому файлу":::
1. <span data-ttu-id="3663e-122">Под **ветвью If yes** добавьте **разъем Excel Online (Business)** с **действием сценария Run.**</span><span class="sxs-lookup"><span data-stu-id="3663e-122">Under the **If yes** branch, add the **Excel Online (Business)** connector with the **Run script** action.</span></span> <span data-ttu-id="3663e-123">Используйте следующие значения для действия:</span><span class="sxs-lookup"><span data-stu-id="3663e-123">Use the following values for the action:</span></span>
    1. <span data-ttu-id="3663e-124">**Расположение**: OneDrive для бизнеса</span><span class="sxs-lookup"><span data-stu-id="3663e-124">**Location**: OneDrive for Business</span></span>
    1. <span data-ttu-id="3663e-125">**Библиотека документов**: OneDrive</span><span class="sxs-lookup"><span data-stu-id="3663e-125">**Document Library**: OneDrive</span></span>
    1. <span data-ttu-id="3663e-126">**Файл**: **Id** (OneDrive значение идентификатора файла)</span><span class="sxs-lookup"><span data-stu-id="3663e-126">**File**: **Id** (the OneDrive file ID value)</span></span>
    1. <span data-ttu-id="3663e-127">**Сценарий**: Ваше имя скрипта</span><span class="sxs-lookup"><span data-stu-id="3663e-127">**Script**: Your script name</span></span>

    :::image type="content" source="../../images/all-files-in-folder-sample-flow-3.png" alt-text="Завершенный разъем Excel Online (Бизнес) в Power Automate":::
1. <span data-ttu-id="3663e-129">Сохранить поток и попробовать его.</span><span class="sxs-lookup"><span data-stu-id="3663e-129">Save the flow and try it out.</span></span>

## <a name="training-video-run-a-script-on-all-excel-files-in-a-folder"></a><span data-ttu-id="3663e-130">Учебное видео: Запустите сценарий на всех Excel файлов в папке</span><span class="sxs-lookup"><span data-stu-id="3663e-130">Training video: Run a script on all Excel files in a folder</span></span>

<span data-ttu-id="3663e-131">[Смотреть Судхи Рамамурти ходить через этот образец на YouTube](https://youtu.be/xMg711o7k6w).</span><span class="sxs-lookup"><span data-stu-id="3663e-131">[Watch Sudhi Ramamurthy walk through this sample on YouTube](https://youtu.be/xMg711o7k6w).</span></span>
