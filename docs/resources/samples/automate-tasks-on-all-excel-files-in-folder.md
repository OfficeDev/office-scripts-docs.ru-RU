---
title: Запуск сценария для всех файлов Excel в папке
description: Узнайте, как запустить сценарий для всех Excel файлов в папке на OneDrive для бизнеса.
ms.date: 04/02/2021
localization_priority: Normal
ms.openlocfilehash: 6376dcac0eb36c04c2b60b2717d18cd730a0a8ee
ms.sourcegitcommit: 1f003c9924e651600c913d84094506125f1055ab
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 04/26/2021
ms.locfileid: "52026857"
---
# <a name="run-a-script-on-all-excel-files-in-a-folder"></a><span data-ttu-id="b1ebc-103">Запуск сценария для всех файлов Excel в папке</span><span class="sxs-lookup"><span data-stu-id="b1ebc-103">Run a script on all Excel files in a folder</span></span>

<span data-ttu-id="b1ebc-104">Этот проект выполняет набор задач автоматизации для всех файлов, расположенных в папке на OneDrive для бизнеса.</span><span class="sxs-lookup"><span data-stu-id="b1ebc-104">This project performs a set of automation tasks on all files situated in a folder on OneDrive for Business.</span></span> <span data-ttu-id="b1ebc-105">Его также можно использовать в SharePoint папке.</span><span class="sxs-lookup"><span data-stu-id="b1ebc-105">It could also be used on a SharePoint folder.</span></span>
<span data-ttu-id="b1ebc-106">Он выполняет вычисления Excel файлов, добавляет форматирование и вставляет комментарий, @mentions [коллеге.](https://support.microsoft.com/office/90701709-5dc1-41c7-aa48-b01d4a46e8c7)</span><span class="sxs-lookup"><span data-stu-id="b1ebc-106">It performs calculations on the Excel files, adds formatting, and inserts a comment that [@mentions](https://support.microsoft.com/office/90701709-5dc1-41c7-aa48-b01d4a46e8c7) a colleague.</span></span>

<span data-ttu-id="b1ebc-107">Скачайте <a href="https://github.com/OfficeDev/office-scripts-docs/blob/master/docs/resources/samples/highlight-alert-excel-files.zip?raw=true"> файлhighlight-alert-excel-files.zip,</a>извлеките файлы в папку с названием **Sales,** используемую в этом примере, и попробуйте ее самостоятельно!</span><span class="sxs-lookup"><span data-stu-id="b1ebc-107">Download the file <a href="https://github.com/OfficeDev/office-scripts-docs/blob/master/docs/resources/samples/highlight-alert-excel-files.zip?raw=true">highlight-alert-excel-files.zip</a>, extract the files to a folder titled **Sales** used in this sample, and try it out yourself!</span></span>

## <a name="sample-code-add-formatting-and-insert-comment"></a><span data-ttu-id="b1ebc-108">Пример кода: добавление форматирования и вставки комментариев</span><span class="sxs-lookup"><span data-stu-id="b1ebc-108">Sample code: Add formatting and insert comment</span></span>

<span data-ttu-id="b1ebc-109">Это сценарий, который выполняется в каждой отдельной книге.</span><span class="sxs-lookup"><span data-stu-id="b1ebc-109">This is the script that runs on each individual workbook.</span></span>

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

## <a name="power-automate-flow-run-the-script-on-every-workbook-in-the-folder"></a><span data-ttu-id="b1ebc-110">Power Automate: запустите сценарий для каждой книги в папке</span><span class="sxs-lookup"><span data-stu-id="b1ebc-110">Power Automate flow: Run the script on every workbook in the folder</span></span>

<span data-ttu-id="b1ebc-111">Этот поток запускает сценарий для каждой книги в папке "Продажи".</span><span class="sxs-lookup"><span data-stu-id="b1ebc-111">This flow runs the script on every workbook in the "Sales" folder.</span></span>

1. <span data-ttu-id="b1ebc-112">Создайте новый **поток мгновенных облаков.**</span><span class="sxs-lookup"><span data-stu-id="b1ebc-112">Create a new **Instant cloud flow**.</span></span>
1. <span data-ttu-id="b1ebc-113">Выберите **вручную вызвать поток и** нажмите **кнопку Создать**.</span><span class="sxs-lookup"><span data-stu-id="b1ebc-113">Select **Manually trigger a flow** and press **Create**.</span></span>
1. <span data-ttu-id="b1ebc-114">Добавьте новый **шаг,** использующий **соединителю OneDrive для бизнеса** и файлы **List в действии папки.**</span><span class="sxs-lookup"><span data-stu-id="b1ebc-114">Add a **New step** that uses the **OneDrive for Business** connector and the **List files in folder** action.</span></span>

    :::image type="content" source="../../images/all-files-in-folder-sample-flow-1.png" alt-text="Завершенный OneDrive для бизнеса в Power Automate.":::
1. <span data-ttu-id="b1ebc-116">Выберите папку "Продажи" с извлеченными книгами.</span><span class="sxs-lookup"><span data-stu-id="b1ebc-116">Select the "Sales" folder with the extracted workbooks.</span></span>
1. <span data-ttu-id="b1ebc-117">Чтобы убедиться, что выбраны только книги, выберите **новый** шаг, а затем выберите **Условие** и установите следующие значения:</span><span class="sxs-lookup"><span data-stu-id="b1ebc-117">To ensure only workbooks are selected, choose **New step**, then select **Condition** and set the following values:</span></span>
    1. <span data-ttu-id="b1ebc-118">**Имя** (значение OneDrive файла)</span><span class="sxs-lookup"><span data-stu-id="b1ebc-118">**Name** (the OneDrive file name value)</span></span>
    1. <span data-ttu-id="b1ebc-119">"заканчивается"</span><span class="sxs-lookup"><span data-stu-id="b1ebc-119">"ends with"</span></span>
    1. <span data-ttu-id="b1ebc-120">xlsx.</span><span class="sxs-lookup"><span data-stu-id="b1ebc-120">"xlsx".</span></span>

    :::image type="content" source="../../images/all-files-in-folder-sample-flow-2.png" alt-text="Блок Power Automate, который применяет последующие действия к каждому файлу.":::
1. <span data-ttu-id="b1ebc-122">В **филиале If Yes** **добавьте соединителю Excel Online (Бизнес)** с действием Сценарий запуска **(предварительного просмотра).**</span><span class="sxs-lookup"><span data-stu-id="b1ebc-122">Under the **If yes** branch, add the **Excel Online (Business)** connector with the **Run script (preview)** action.</span></span> <span data-ttu-id="b1ebc-123">Используйте следующие значения для действия:</span><span class="sxs-lookup"><span data-stu-id="b1ebc-123">Use the following values for the action:</span></span>
    1. <span data-ttu-id="b1ebc-124">**Расположение**: OneDrive для бизнеса</span><span class="sxs-lookup"><span data-stu-id="b1ebc-124">**Location**: OneDrive for Business</span></span>
    1. <span data-ttu-id="b1ebc-125">**Библиотека документов**: OneDrive</span><span class="sxs-lookup"><span data-stu-id="b1ebc-125">**Document Library**: OneDrive</span></span>
    1. <span data-ttu-id="b1ebc-126">**Файл**: **Id** (OneDrive файла)</span><span class="sxs-lookup"><span data-stu-id="b1ebc-126">**File**: **Id** (the OneDrive file ID value)</span></span>
    1. <span data-ttu-id="b1ebc-127">**Сценарий:** имя сценария</span><span class="sxs-lookup"><span data-stu-id="b1ebc-127">**Script**: Your script name</span></span>

    :::image type="content" source="../../images/all-files-in-folder-sample-flow-3.png" alt-text="Завершенный соедините Excel Online (Бизнес) в Power Automate.":::
1. <span data-ttu-id="b1ebc-129">Сохраните поток и попробуйте его.</span><span class="sxs-lookup"><span data-stu-id="b1ebc-129">Save the flow and try it out.</span></span>

## <a name="training-video-run-a-script-on-all-excel-files-in-a-folder"></a><span data-ttu-id="b1ebc-130">Обучающее видео: запустите сценарий для всех Excel файлов в папке</span><span class="sxs-lookup"><span data-stu-id="b1ebc-130">Training video: Run a script on all Excel files in a folder</span></span>

<span data-ttu-id="b1ebc-131">[Просмотрите пошаговую](https://youtu.be/xMg711o7k6w) видеозапись запуска сценария для всех Excel файлов в OneDrive для бизнеса или SharePoint папке.</span><span class="sxs-lookup"><span data-stu-id="b1ebc-131">[Watch step-by-step video](https://youtu.be/xMg711o7k6w) on how to run a script on all Excel files in a OneDrive for Business or SharePoint folder.</span></span>
