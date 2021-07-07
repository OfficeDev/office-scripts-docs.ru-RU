---
title: Запуск сценария для всех файлов Excel в папке
description: Узнайте, как запустить сценарий для всех Excel файлов в папке на OneDrive для бизнеса.
ms.date: 06/29/2021
localization_priority: Normal
ms.openlocfilehash: bf9c0c486dacced5c3017b267ea65dfd215a5197
ms.sourcegitcommit: 211c157ca746e266eeb079f5fa1925a1e35ab702
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/07/2021
ms.locfileid: "53313899"
---
# <a name="run-a-script-on-all-excel-files-in-a-folder"></a><span data-ttu-id="14261-103">Запуск сценария для всех файлов Excel в папке</span><span class="sxs-lookup"><span data-stu-id="14261-103">Run a script on all Excel files in a folder</span></span>

<span data-ttu-id="14261-104">Этот проект выполняет набор задач автоматизации для всех файлов, расположенных в папке на OneDrive для бизнеса.</span><span class="sxs-lookup"><span data-stu-id="14261-104">This project performs a set of automation tasks on all files situated in a folder on OneDrive for Business.</span></span> <span data-ttu-id="14261-105">Его также можно использовать в SharePoint папке.</span><span class="sxs-lookup"><span data-stu-id="14261-105">It could also be used on a SharePoint folder.</span></span>
<span data-ttu-id="14261-106">Он выполняет вычисления Excel файлов, добавляет форматирование и вставляет комментарий, @mentions [коллеге.](https://support.microsoft.com/office/90701709-5dc1-41c7-aa48-b01d4a46e8c7)</span><span class="sxs-lookup"><span data-stu-id="14261-106">It performs calculations on the Excel files, adds formatting, and inserts a comment that [@mentions](https://support.microsoft.com/office/90701709-5dc1-41c7-aa48-b01d4a46e8c7) a colleague.</span></span>

## <a name="sample-excel-files"></a><span data-ttu-id="14261-107">Пример Excel файлов</span><span class="sxs-lookup"><span data-stu-id="14261-107">Sample Excel files</span></span>

<span data-ttu-id="14261-108">Скачайте <a href="https://github.com/OfficeDev/office-scripts-docs/blob/master/docs/resources/samples/highlight-alert-excel-files.zip?raw=true">highlight-alert-excel-files.zip</a> для всех книг, необходимых для этого примера.</span><span class="sxs-lookup"><span data-stu-id="14261-108">Download <a href="https://github.com/OfficeDev/office-scripts-docs/blob/master/docs/resources/samples/highlight-alert-excel-files.zip?raw=true">highlight-alert-excel-files.zip</a> for all the workbooks you'll need for this sample.</span></span> <span data-ttu-id="14261-109">Извлечение этих файлов в папку с названием **Sales**.</span><span class="sxs-lookup"><span data-stu-id="14261-109">Extract those files to a folder titled **Sales**.</span></span> <span data-ttu-id="14261-110">Добавьте следующий сценарий в свою коллекцию скриптов, чтобы попробовать пример самостоятельно!</span><span class="sxs-lookup"><span data-stu-id="14261-110">Add the following script to your script collection to try the sample yourself!</span></span>

## <a name="sample-code-add-formatting-and-insert-comment"></a><span data-ttu-id="14261-111">Пример кода: добавление форматирования и вставки комментариев</span><span class="sxs-lookup"><span data-stu-id="14261-111">Sample code: Add formatting and insert comment</span></span>

<span data-ttu-id="14261-112">Это сценарий, который выполняется в каждой отдельной книге.</span><span class="sxs-lookup"><span data-stu-id="14261-112">This is the script that runs on each individual workbook.</span></span>

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

## <a name="power-automate-flow-run-the-script-on-every-workbook-in-the-folder"></a><span data-ttu-id="14261-113">Power Automate: запустите сценарий для каждой книги в папке</span><span class="sxs-lookup"><span data-stu-id="14261-113">Power Automate flow: Run the script on every workbook in the folder</span></span>

<span data-ttu-id="14261-114">Этот поток запускает сценарий для каждой книги в папке "Продажи".</span><span class="sxs-lookup"><span data-stu-id="14261-114">This flow runs the script on every workbook in the "Sales" folder.</span></span>

1. <span data-ttu-id="14261-115">Создайте новый **поток мгновенных облаков.**</span><span class="sxs-lookup"><span data-stu-id="14261-115">Create a new **Instant cloud flow**.</span></span>
1. <span data-ttu-id="14261-116">Выберите **вручную вызвать поток и** выберите **Создать**.</span><span class="sxs-lookup"><span data-stu-id="14261-116">Choose **Manually trigger a flow** and select **Create**.</span></span>
1. <span data-ttu-id="14261-117">Добавьте новый **шаг,** использующий **соединителю OneDrive для бизнеса** и файлы **List в действии папки.**</span><span class="sxs-lookup"><span data-stu-id="14261-117">Add a **New step** that uses the **OneDrive for Business** connector and the **List files in folder** action.</span></span>

    :::image type="content" source="../../images/all-files-in-folder-sample-flow-1.png" alt-text="Завершенный OneDrive для бизнеса в Power Automate.":::
1. <span data-ttu-id="14261-119">Выберите папку "Продажи" с извлеченными книгами.</span><span class="sxs-lookup"><span data-stu-id="14261-119">Select the "Sales" folder with the extracted workbooks.</span></span>
1. <span data-ttu-id="14261-120">Чтобы убедиться, что выбраны только книги, выберите **новый** шаг, а затем выберите **Условие** и установите следующие значения:</span><span class="sxs-lookup"><span data-stu-id="14261-120">To ensure only workbooks are selected, choose **New step**, then select **Condition** and set the following values:</span></span>
    1. <span data-ttu-id="14261-121">**Имя** (значение OneDrive файла)</span><span class="sxs-lookup"><span data-stu-id="14261-121">**Name** (the OneDrive file name value)</span></span>
    1. <span data-ttu-id="14261-122">"заканчивается"</span><span class="sxs-lookup"><span data-stu-id="14261-122">"ends with"</span></span>
    1. <span data-ttu-id="14261-123">xlsx.</span><span class="sxs-lookup"><span data-stu-id="14261-123">"xlsx".</span></span>

    :::image type="content" source="../../images/all-files-in-folder-sample-flow-2.png" alt-text="Блок Power Automate, который применяет последующие действия к каждому файлу.":::
1. <span data-ttu-id="14261-125">В **филиале If Yes** **добавьте соединителю Excel Online (Бизнес)** с действием **сценария Run.**</span><span class="sxs-lookup"><span data-stu-id="14261-125">Under the **If yes** branch, add the **Excel Online (Business)** connector with the **Run script** action.</span></span> <span data-ttu-id="14261-126">Используйте следующие значения для действия:</span><span class="sxs-lookup"><span data-stu-id="14261-126">Use the following values for the action:</span></span>
    1. <span data-ttu-id="14261-127">**Расположение**: OneDrive для бизнеса</span><span class="sxs-lookup"><span data-stu-id="14261-127">**Location**: OneDrive for Business</span></span>
    1. <span data-ttu-id="14261-128">**Библиотека документов**: OneDrive</span><span class="sxs-lookup"><span data-stu-id="14261-128">**Document Library**: OneDrive</span></span>
    1. <span data-ttu-id="14261-129">**Файл**: **Id** (OneDrive файла)</span><span class="sxs-lookup"><span data-stu-id="14261-129">**File**: **Id** (the OneDrive file ID value)</span></span>
    1. <span data-ttu-id="14261-130">**Сценарий:** имя сценария</span><span class="sxs-lookup"><span data-stu-id="14261-130">**Script**: Your script name</span></span>

    :::image type="content" source="../../images/all-files-in-folder-sample-flow-3.png" alt-text="Завершенный соедините Excel Online (Бизнес) в Power Automate.":::
1. <span data-ttu-id="14261-132">Сохраните поток и попробуйте его. Используйте **кнопку Test** на странице редактора потока или запустите поток через вкладку **Мои потоки.** Не забудьте разрешить доступ при запросе.</span><span class="sxs-lookup"><span data-stu-id="14261-132">Save the flow and try it out. Use the **Test** button on the flow editor page or run the flow through your **My flows** tab. Be sure to allow access when prompted.</span></span>

## <a name="training-video-run-a-script-on-all-excel-files-in-a-folder"></a><span data-ttu-id="14261-133">Обучающее видео: запустите сценарий для всех Excel файлов в папке</span><span class="sxs-lookup"><span data-stu-id="14261-133">Training video: Run a script on all Excel files in a folder</span></span>

<span data-ttu-id="14261-134">[Смотреть Sudhi Ramamurthy ходить через этот пример на YouTube](https://youtu.be/xMg711o7k6w).</span><span class="sxs-lookup"><span data-stu-id="14261-134">[Watch Sudhi Ramamurthy walk through this sample on YouTube](https://youtu.be/xMg711o7k6w).</span></span>
