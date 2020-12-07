---
title: Передача данных сценариям в автоматически запускаемых рабочих процессах Power Automate
description: Учебное руководство, посвященное запуску сценариев Office для Excel в Интернете с помощью Power Automate при получении электронной почты с дальнейшей передачей данных рабочего процесса в сценарий.
ms.date: 11/30/2020
localization_priority: Priority
ms.openlocfilehash: b73f40c70669fedbe8a0adcf346995cb20b62d37
ms.sourcegitcommit: af487756dffea0f8f0cd62710c586842cb08073c
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 12/04/2020
ms.locfileid: "49571481"
---
# <a name="pass-data-to-scripts-in-an-automatically-run-power-automate-flow-preview"></a><span data-ttu-id="32f18-103">Передача данных сценариям в автоматически запускаемых рабочих процессах Power Automate (предварительная версия)</span><span class="sxs-lookup"><span data-stu-id="32f18-103">Pass data to scripts in an automatically-run Power Automate flow (preview)</span></span>

<span data-ttu-id="32f18-104">В этом руководстве объясняется, как использовать сценарий Office для Excel в Интернете с помощью автоматизированных рабочих процессов [Power Automate](https://flow.microsoft.com).</span><span class="sxs-lookup"><span data-stu-id="32f18-104">This tutorial teaches you how to use an Office Script for Excel on the web with an automated [Power Automate](https://flow.microsoft.com) workflow.</span></span> <span data-ttu-id="32f18-105">Сценарий будет автоматически выполняться каждый раз при получении электронной почты. Данные из сообщений электронной почты будут записываться в книгу Excel.</span><span class="sxs-lookup"><span data-stu-id="32f18-105">Your script will automatically run each time you receive an email, recording information from the email in an Excel workbook.</span></span> <span data-ttu-id="32f18-106">Возможность передавать данные из других приложений в сценарии Office предоставляет вам значительную гибкость и свободу в автоматизированных процессах.</span><span class="sxs-lookup"><span data-stu-id="32f18-106">Being able to pass data from other applications into an Office Script gives you a great deal of flexibility and freedom in your automated processes.</span></span>

> [!TIP]
> <span data-ttu-id="32f18-107">Если вы только приступили к работе со сценариями Office, рекомендуем начать с учебника [Запись, редактирование и создание сценариев Office в Excel в Интернете](excel-tutorial.md).</span><span class="sxs-lookup"><span data-stu-id="32f18-107">If you are new to Office Scripts, we recommend starting with the [Record, edit, and create Office Scripts in Excel on the web](excel-tutorial.md) tutorial.</span></span> <span data-ttu-id="32f18-108">Если вы впервые используете Power Automate, рекомендуем начать с учебника [Вызов сценариев из активированного вручную потока Power Automate](excel-power-automate-manual.md).</span><span class="sxs-lookup"><span data-stu-id="32f18-108">If you are new to Power Automate, we recommend starting with the [Call scripts from a manual Power Automate flow](excel-power-automate-manual.md) tutorial.</span></span> <span data-ttu-id="32f18-109">[Сценарии Office используют TypeScript](../overview/code-editor-environment.md), и этот учебник предназначен для пользователей с начальным и средним уровнем знаний по JavaScript или TypeScript.</span><span class="sxs-lookup"><span data-stu-id="32f18-109">[Office Scripts use TypeScript](../overview/code-editor-environment.md) and this tutorial is intended for people with beginner to intermediate-level knowledge of JavaScript or TypeScript.</span></span> <span data-ttu-id="32f18-110">Если вы впервые работаете с JavaScript, рекомендуем начать с [учебника Mozilla по JavaScript](https://developer.mozilla.org/docs/Web/JavaScript/Guide/Introduction).</span><span class="sxs-lookup"><span data-stu-id="32f18-110">If you're new to JavaScript, we recommend starting with the [Mozilla JavaScript tutorial](https://developer.mozilla.org/docs/Web/JavaScript/Guide/Introduction).</span></span>

## <a name="prerequisites"></a><span data-ttu-id="32f18-111">Предварительные условия</span><span class="sxs-lookup"><span data-stu-id="32f18-111">Prerequisites</span></span>

[!INCLUDE [Tutorial prerequisites](../includes/power-automate-tutorial-prerequisites.md)]

## <a name="prepare-the-workbook"></a><span data-ttu-id="32f18-112">Подготовка книги</span><span class="sxs-lookup"><span data-stu-id="32f18-112">Prepare the workbook</span></span>

<span data-ttu-id="32f18-113">Power Automate не может использовать [относительные ссылки](../develop/power-automate-integration.md#avoid-using-relative-references), такие как `Workbook.getActiveWorksheet`, для доступа к компонентам книги.</span><span class="sxs-lookup"><span data-stu-id="32f18-113">Power Automate can't use [relative references](../develop/power-automate-integration.md#avoid-using-relative-references) like `Workbook.getActiveWorksheet` to access workbook components.</span></span> <span data-ttu-id="32f18-114">Поэтому нужно, чтобы в книге и в таблице были согласованные имена, на которые сможет ссылаться Power Automate.</span><span class="sxs-lookup"><span data-stu-id="32f18-114">So, we need a workbook and worksheet with consistent names for Power Automate to reference.</span></span>

1. <span data-ttu-id="32f18-115">Создайте новую книгу с именем **MyWorkbook**.</span><span class="sxs-lookup"><span data-stu-id="32f18-115">Create a new workbook named **MyWorkbook**.</span></span>

2. <span data-ttu-id="32f18-116">Перейдите на вкладку **Автоматизация** и выберите **Редактор кода**.</span><span class="sxs-lookup"><span data-stu-id="32f18-116">Go to the **Automate** tab and select **Code Editor**.</span></span>

3. <span data-ttu-id="32f18-117">Выберите **Новый сценарий**.</span><span class="sxs-lookup"><span data-stu-id="32f18-117">Select **New Script**.</span></span>

4. <span data-ttu-id="32f18-118">Замените имеющийся код на следующий и нажмите кнопку **Выполнить**.</span><span class="sxs-lookup"><span data-stu-id="32f18-118">Replace the existing code with the following script and press **Run**.</span></span> <span data-ttu-id="32f18-119">При том будет создана книга с нужными именами листа, таблицы и сводной таблицы.</span><span class="sxs-lookup"><span data-stu-id="32f18-119">This will setup the workbook with consistent worksheet, table, and PivotTable names.</span></span>

    ```TypeScript
    function main(workbook: ExcelScript.Workbook) {
      // Add a new worksheet to store our email table
      let emailsSheet = workbook.addWorksheet("Emails");

      // Add data and create a table
      emailsSheet.getRange("A1:D1").setValues([
        ["Date", "Day of the week", "Email address", "Subject"]
      ]);
      let newTable = workbook.addTable(emailsSheet.getRange("A1:D2"), true);
      newTable.setName("EmailTable");

      // Add a new PivotTable to a new worksheet
      let pivotWorksheet = workbook.addWorksheet("Subjects");
      let newPivotTable = workbook.addPivotTable("Pivot", "EmailTable", pivotWorksheet.getRange("A3:C20"));

      // Setup the pivot hierarchies
      newPivotTable.addRowHierarchy(newPivotTable.getHierarchy("Day of the week"));
      newPivotTable.addRowHierarchy(newPivotTable.getHierarchy("Email address"));
      newPivotTable.addDataHierarchy(newPivotTable.getHierarchy("Subject"));
    }
    ```

## <a name="create-an-office-script"></a><span data-ttu-id="32f18-120">Создание сценария Office</span><span class="sxs-lookup"><span data-stu-id="32f18-120">Create an Office Script</span></span>

<span data-ttu-id="32f18-121">Создадим сценарий, записывающий информацию из электронной почты.</span><span class="sxs-lookup"><span data-stu-id="32f18-121">Let's create a script that logs information from an email.</span></span> <span data-ttu-id="32f18-122">Предположим, что нужно узнать, в какие дни недели мы получаем больше всего почты, и сколько уникальных отправителей отправляют ее.</span><span class="sxs-lookup"><span data-stu-id="32f18-122">We want to know how which days of the week we receive the most mail and how many unique senders are sending that mail.</span></span> <span data-ttu-id="32f18-123">В нашей книге содержится таблица со столбцами **Дата**, **День недели**, **Адрес электронной почты** и **Тема**.</span><span class="sxs-lookup"><span data-stu-id="32f18-123">Our workbook has a table with **Date**, **Day of the week**, **Email address**, and **Subject** columns.</span></span> <span data-ttu-id="32f18-124">Кроме того, в книге содержится сводная таблица, содержащая **День недели** и **Адрес электронной почты** (это иерархии строк).</span><span class="sxs-lookup"><span data-stu-id="32f18-124">Our worksheet also has a PivotTable that is pivoting on the **Day of the week** and **Email address** (those are the row hierarchies).</span></span> <span data-ttu-id="32f18-125">Количество уникальных **тем** — это отображаемая объединенная информация (иерархия данных).</span><span class="sxs-lookup"><span data-stu-id="32f18-125">The count of unique **Subjects** is the aggregated information being displayed (the data hierarchy).</span></span> <span data-ttu-id="32f18-126">Наш сценарий будет обновлять эту сводную таблицу после обновления таблицы электронной почты.</span><span class="sxs-lookup"><span data-stu-id="32f18-126">We'll have our script refresh that PivotTable after updating the email table.</span></span>

1. <span data-ttu-id="32f18-127">В окне **Редактор кода** выберите **Создать сценарий**.</span><span class="sxs-lookup"><span data-stu-id="32f18-127">From within the **Code Editor**, select **New Script**.</span></span>

2. <span data-ttu-id="32f18-128">Поток, который мы создадим на более позднем этапе, будет отправлять данные о каждом полученном сообщении электронной почты.</span><span class="sxs-lookup"><span data-stu-id="32f18-128">The flow that we'll create later in the tutorial will send our script information about each email that's received.</span></span> <span data-ttu-id="32f18-129">Сценарий должен обращаться к этим входным данным с помощью параметров в функции `main`.</span><span class="sxs-lookup"><span data-stu-id="32f18-129">The script needs to accept that input through parameters in the `main` function.</span></span> <span data-ttu-id="32f18-130">Замените сценарий по умолчанию следующим сценарием.</span><span class="sxs-lookup"><span data-stu-id="32f18-130">Replace the default script with the following script:</span></span>

    ```TypeScript
    function main(
      workbook: ExcelScript.Workbook,
      from: string,
      dateReceived: string,
      subject: string) {

    }
    ```

3. <span data-ttu-id="32f18-131">Этому сценарию требуется доступ к таблице книги и к сводной таблице.</span><span class="sxs-lookup"><span data-stu-id="32f18-131">The script needs access to the workbook's table and PivotTable.</span></span> <span data-ttu-id="32f18-132">Добавьте следующий код в текст сценария после открывающего символа `{`:</span><span class="sxs-lookup"><span data-stu-id="32f18-132">Add the following code to the body of the script, after the opening `{`:</span></span>

    ```TypeScript
    // Get the email table.
    let emailWorksheet = workbook.getWorksheet("Emails");
    let table = emailWorksheet.getTable("EmailTable");
  
    // Get the PivotTable.
    let pivotTableWorksheet = workbook.getWorksheet("Subjects");
    let pivotTable = pivotTableWorksheet.getPivotTable("Pivot");
    ```

4. <span data-ttu-id="32f18-133">Параметр `dateReceived` относится к типу `string`.</span><span class="sxs-lookup"><span data-stu-id="32f18-133">The `dateReceived` parameter is of type `string`.</span></span> <span data-ttu-id="32f18-134">Преобразуем его в объекту [`Date`](../develop/javascript-objects.md#date), чтобы можно было удобно получать день недели.</span><span class="sxs-lookup"><span data-stu-id="32f18-134">Let's convert that to a [`Date` object](../develop/javascript-objects.md#date) so we can easily get the day of the week.</span></span> <span data-ttu-id="32f18-135">После этого нужно будет сопоставить значение номера дня с более читаемой версией.</span><span class="sxs-lookup"><span data-stu-id="32f18-135">After doing that, we'll need to map the day's number value to a more readable version.</span></span> <span data-ttu-id="32f18-136">Добавьте следующий код в конце сценария перед закрывающим символом `}`</span><span class="sxs-lookup"><span data-stu-id="32f18-136">Add the following code to the end of your script, before the closing `}`:</span></span>

    ```TypeScript
      // Parse the received date string to determine the day of the week.
      let emailDate = new Date(dateReceived);
      let dayName = emailDate.toLocaleDateString("en-US", { weekday: 'long' });
    ```

5. <span data-ttu-id="32f18-137">Строка `subject` может включать тег ответа "RE:".</span><span class="sxs-lookup"><span data-stu-id="32f18-137">The `subject` string may include the "RE:" reply tag.</span></span> <span data-ttu-id="32f18-138">Давайте удалим этот тег из строки, чтобы у сообщений электронной почте в одной и той же беседе была одинаковая тема для таблицы.</span><span class="sxs-lookup"><span data-stu-id="32f18-138">Let's remove that from the string so that emails in the same thread have the same subject for the table.</span></span> <span data-ttu-id="32f18-139">Добавьте следующий код в конце сценария перед закрывающим символом `}`</span><span class="sxs-lookup"><span data-stu-id="32f18-139">Add the following code to the end of your script, before the closing `}`:</span></span>

    ```TypeScript
    // Remove the reply tag from the email subject to group emails on the same thread.
    let subjectText = subject.replace("Re: ", "");
    subjectText = subjectText.replace("RE: ", "");
    ```

6. <span data-ttu-id="32f18-140">Теперь, когда данные электронной почты отформатированы по нашему желанию, добавим строку в таблицу электронной почты.</span><span class="sxs-lookup"><span data-stu-id="32f18-140">Now that the email data has been formatted to our liking, let's add a row to the email table.</span></span> <span data-ttu-id="32f18-141">Добавьте следующий код в конце сценария перед закрывающим символом `}`</span><span class="sxs-lookup"><span data-stu-id="32f18-141">Add the following code to the end of your script, before the closing `}`:</span></span>

    ```TypeScript
    // Add the parsed text to the table.
    table.addRow(-1, [dateReceived, dayName, from, subjectText]);
    ```

7. <span data-ttu-id="32f18-142">Теперь нужно обновить сводную таблицу.</span><span class="sxs-lookup"><span data-stu-id="32f18-142">Finally, let's make sure the PivotTable is refreshed.</span></span> <span data-ttu-id="32f18-143">Добавьте следующий код в конце сценария перед закрывающим символом `}`</span><span class="sxs-lookup"><span data-stu-id="32f18-143">Add the following code to the end of your script, before the closing `}`:</span></span>

    ```TypeScript
    // Refresh the PivotTable to include the new row.
    pivotTable.refresh();
    ```

8. <span data-ttu-id="32f18-144">Переименуйте сценарий в **Запись электронной почты** и нажмите кнопку **Сохранить сценарий**.</span><span class="sxs-lookup"><span data-stu-id="32f18-144">Rename your script **Record Email** and press **Save script**.</span></span>

<span data-ttu-id="32f18-145">Теперь сценарий готов для рабочего процесса Power Automate.</span><span class="sxs-lookup"><span data-stu-id="32f18-145">Your script is now ready for a Power Automate workflow.</span></span> <span data-ttu-id="32f18-146">Сценарий должен выглядеть примерно так:</span><span class="sxs-lookup"><span data-stu-id="32f18-146">It should look like the following script:</span></span>

```TypeScript
function main(
  workbook: ExcelScript.Workbook,
  from: string,
  dateReceived: string,
  subject: string) {
  // Get the email table.
  let emailWorksheet = workbook.getWorksheet("Emails");
  let table = emailWorksheet.getTable("EmailTable");

  // Get the PivotTable.
  let pivotTableWorksheet = workbook.getWorksheet("Subjects");
  let pivotTable = pivotTableWorksheet.getPivotTable("Pivot");

  // Parse the received date string to determine the day of the week.
  let emailDate = new Date(dateReceived);
  let dayName = emailDate.toLocaleDateString("en-US", { weekday: 'long' });

  // Remove the reply tag from the email subject to group emails on the same thread.
  let subjectText = subject.replace("Re: ", "");
  subjectText = subjectText.replace("RE: ", "");

  // Add the parsed text to the table.
  table.addRow(-1, [dateReceived, dayName, from, subjectText]);

  // Refresh the PivotTable to include the new row.
  pivotTable.refresh();
}
```

## <a name="create-an-automated-workflow-with-power-automate"></a><span data-ttu-id="32f18-147">Создание автоматизированного рабочего процесса с помощью Power Automate</span><span class="sxs-lookup"><span data-stu-id="32f18-147">Create an automated workflow with Power Automate</span></span>

1. <span data-ttu-id="32f18-148">Войдите на [сайт Power Automate](https://flow.microsoft.com).</span><span class="sxs-lookup"><span data-stu-id="32f18-148">Sign in to the [Power Automate site](https://flow.microsoft.com).</span></span>

2. <span data-ttu-id="32f18-149">В меню в левой части экрана выберите **Создать**.</span><span class="sxs-lookup"><span data-stu-id="32f18-149">In the menu that's displayed on the left side of the screen, press **Create**.</span></span> <span data-ttu-id="32f18-150">При этом откроется список способов создания новых рабочих процессов.</span><span class="sxs-lookup"><span data-stu-id="32f18-150">This brings you to list of ways to create new workflows.</span></span>

    ![Кнопка "Создать" в Power Automate.](../images/power-automate-tutorial-1.png)

3. <span data-ttu-id="32f18-152">В разделе **Начать с пустого** выберите **Автоматизированный рабочий процесс**.</span><span class="sxs-lookup"><span data-stu-id="32f18-152">In the **Start from blank** section, select **Automated flow**.</span></span> <span data-ttu-id="32f18-153">В этом случае создается рабочий процесс, запускаемый каким-либо событием, например получением сообщения электронной почты.</span><span class="sxs-lookup"><span data-stu-id="32f18-153">This creates a workflow triggered by an event, such as receiving an email.</span></span>

    ![Функция "Автоматизированный рабочий процесс" в Power Automate.](../images/power-automate-params-tutorial-1.png)

4. <span data-ttu-id="32f18-155">В появившемся диалоговом окне введите имя рабочего процесса в текстовом поле **Имя рабочего процесса**.</span><span class="sxs-lookup"><span data-stu-id="32f18-155">In the dialog window that appears, enter a name for your flow in the **Flow name** text box.</span></span> <span data-ttu-id="32f18-156">Затем выберите **При получении новой электронной почты** в списке параметров **Выберите триггер рабочего процесса**.</span><span class="sxs-lookup"><span data-stu-id="32f18-156">Then select **When a new email arrives** from the list of options under **Choose your flow's trigger**.</span></span> <span data-ttu-id="32f18-157">Может потребоваться найти этот параметр с помощью поля поиска.</span><span class="sxs-lookup"><span data-stu-id="32f18-157">You may need to search for the option using the search box.</span></span> <span data-ttu-id="32f18-158">Затем нажмите кнопку **Создать**.</span><span class="sxs-lookup"><span data-stu-id="32f18-158">Finally, press **Create**.</span></span>

    ![Часть окна "Создание автоматизированного рабочего процесса" в Power Automate с параметром "получение нового сообщения электронной почты".](../images/power-automate-params-tutorial-2.png)

    > [!NOTE]
    > <span data-ttu-id="32f18-160">В этом учебном руководстве используется Outlook.</span><span class="sxs-lookup"><span data-stu-id="32f18-160">This tutorial uses Outlook.</span></span> <span data-ttu-id="32f18-161">Можно использовать любую предпочитаемую вами службу электронной почты, хотя в этом случае некоторые параметры могут отличаться.</span><span class="sxs-lookup"><span data-stu-id="32f18-161">Feel free to use your preferred email service instead, though some options may be different.</span></span>

5. <span data-ttu-id="32f18-162">Нажмите кнопку **Новый шаг**.</span><span class="sxs-lookup"><span data-stu-id="32f18-162">Press **New step**.</span></span>

6. <span data-ttu-id="32f18-163">Перейдите на вкладку **Стандартные** и выберите **Excel Online (бизнес)**.</span><span class="sxs-lookup"><span data-stu-id="32f18-163">Select the **Standard** tab, then select **Excel Online (Business)**.</span></span>

    ![Функция Power Automate для Excel Online (бизнес).](../images/power-automate-tutorial-4.png)

7. <span data-ttu-id="32f18-165">В разделе **Действия** выберите **Запустить сценарий (предварительная версия)**.</span><span class="sxs-lookup"><span data-stu-id="32f18-165">Under **Actions**, select **Run script (preview)**.</span></span>

    ![Вариант действия Power Automate "Запуск сценария" (предварительная версия).](../images/power-automate-tutorial-5.png)

8. <span data-ttu-id="32f18-167">Затем выберите книгу, сценарий и исходные аргументы сценария для использования на следующем шаге.</span><span class="sxs-lookup"><span data-stu-id="32f18-167">Next, you'll select the workbook, script, and script input arguments to use in the flow step.</span></span> <span data-ttu-id="32f18-168">В этом учебнике вы будете использовать книгу, созданную в OneDrive, но вы можете воспользоваться любой книгой в OneDrive или на сайте SharePoint.</span><span class="sxs-lookup"><span data-stu-id="32f18-168">For the tutorial, you'll use the workbook you created in your OneDrive, but you could use any workbook in a OneDrive or SharePoint site.</span></span> <span data-ttu-id="32f18-169">Укажите следующие параметры для соединителя **Запуск сценария**.</span><span class="sxs-lookup"><span data-stu-id="32f18-169">Specify the following settings for the **Run script** connector:</span></span>

    - <span data-ttu-id="32f18-170">**Расположение**: OneDrive для бизнеса</span><span class="sxs-lookup"><span data-stu-id="32f18-170">**Location**: OneDrive for Business</span></span>
    - <span data-ttu-id="32f18-171">**Библиотека документов**: OneDrive</span><span class="sxs-lookup"><span data-stu-id="32f18-171">**Document Library**: OneDrive</span></span>
    - <span data-ttu-id="32f18-172">**Файл**: MyWorkbook.xlsx *(выбран с помощью браузера файлов)*</span><span class="sxs-lookup"><span data-stu-id="32f18-172">**File**: MyWorkbook.xlsx *(Chosen through the file browser)*</span></span>
    - <span data-ttu-id="32f18-173">**Сценарий**: запись электронной почты</span><span class="sxs-lookup"><span data-stu-id="32f18-173">**Script**: Record Email</span></span>
    - <span data-ttu-id="32f18-174">**от**: От *(динамическое содержимое из Outlook)*</span><span class="sxs-lookup"><span data-stu-id="32f18-174">**from**: From *(dynamic content from Outlook)*</span></span>
    - <span data-ttu-id="32f18-175">**dateReceived**: Время получения *(динамическое содержимое из Outlook)*</span><span class="sxs-lookup"><span data-stu-id="32f18-175">**dateReceived**: Received Time *(dynamic content from Outlook)*</span></span>
    - <span data-ttu-id="32f18-176">**тема**: Тема *(динамическое содержимое из Outlook)*</span><span class="sxs-lookup"><span data-stu-id="32f18-176">**subject**: Subject *(dynamic content from Outlook)*</span></span>

    <span data-ttu-id="32f18-177">*Обратите внимание, что эти параметры сценария будут отображаться только после выбора сценария.*</span><span class="sxs-lookup"><span data-stu-id="32f18-177">*Note that the parameters for the script will only appear once the script is selected.*</span></span>

    ![Параметры варианта действия Power Automate "Запуск сценария" (предварительная версия).](../images/power-automate-params-tutorial-3.png)

9. <span data-ttu-id="32f18-179">Нажмите кнопку **Сохранить**.</span><span class="sxs-lookup"><span data-stu-id="32f18-179">Press **Save**.</span></span>

<span data-ttu-id="32f18-180">Теперь рабочий процесс включен.</span><span class="sxs-lookup"><span data-stu-id="32f18-180">Your flow is now enabled.</span></span> <span data-ttu-id="32f18-181">Он будет автоматически выполнять сценарий каждый раз при получении сообщения электронной почты через Outlook.</span><span class="sxs-lookup"><span data-stu-id="32f18-181">It will automatically run your script each time you receive an email through Outlook.</span></span>

## <a name="manage-the-script-in-power-automate"></a><span data-ttu-id="32f18-182">Управление сценарием в Power Automate</span><span class="sxs-lookup"><span data-stu-id="32f18-182">Manage the script in Power Automate</span></span>

1. <span data-ttu-id="32f18-183">На главной странице Power Automate выберите **Мои рабочие процессы**.</span><span class="sxs-lookup"><span data-stu-id="32f18-183">From the main Power Automate page, select **My flows**.</span></span>

    ![Кнопка "Мои рабочие процессы" в Power Automate.](../images/power-automate-tutorial-7.png)

2. <span data-ttu-id="32f18-185">Выберите рабочий процесс.</span><span class="sxs-lookup"><span data-stu-id="32f18-185">Select your flow.</span></span> <span data-ttu-id="32f18-186">Здесь можно просмотреть историю запусков.</span><span class="sxs-lookup"><span data-stu-id="32f18-186">Here you can see the run history.</span></span> <span data-ttu-id="32f18-187">Можно обновить страницу или нажать кнопку обновления **всех запусков**, чтобы обновить историю.</span><span class="sxs-lookup"><span data-stu-id="32f18-187">You can refresh the page or press the refresh **All runs** button to update the history.</span></span> <span data-ttu-id="32f18-188">Рабочий процесс запустится вскоре после получения сообщения электронной почты.</span><span class="sxs-lookup"><span data-stu-id="32f18-188">The flow will trigger shortly after an email is received.</span></span> <span data-ttu-id="32f18-189">Проверьте рабочий процесс, отправив себе сообщение электронной почты.</span><span class="sxs-lookup"><span data-stu-id="32f18-189">Test the flow by sending yourself mail.</span></span>

<span data-ttu-id="32f18-190">При срабатывании рабочего процесса и успешном выполнении сценария должна обновляться таблица книги и сводная таблица.</span><span class="sxs-lookup"><span data-stu-id="32f18-190">When the flow is triggered and successfully runs your script, you should see the workbook's table and PivotTable update.</span></span>

![Таблица электронной почты после нескольких выполнений рабочего процесса.](../images/power-automate-params-tutorial-4.png)

![Сводная таблица после нескольких выполнений рабочего процесса.](../images/power-automate-params-tutorial-5.png)

## <a name="next-steps"></a><span data-ttu-id="32f18-193">Дальнейшие действия</span><span class="sxs-lookup"><span data-stu-id="32f18-193">Next steps</span></span>

<span data-ttu-id="32f18-194">Посетите страницу [Запуск сценариев Office с помощью Power Automate](../develop/power-automate-integration.md) для получения дополнительных сведений о подключениях сценариев Office с помощью Power Automate.</span><span class="sxs-lookup"><span data-stu-id="32f18-194">Visit [Run Office Scripts with Power Automate](../develop/power-automate-integration.md) to learn more about connecting Office Scripts with Power Automate.</span></span>

<span data-ttu-id="32f18-195">Кроме того, прочтите статью [Образец сценария автоматизированных напоминаний о задачах](../resources/scenarios/task-reminders.md), чтобы узнать, как использовать сценарии Office и Power Automate с адаптивными карточками Teams.</span><span class="sxs-lookup"><span data-stu-id="32f18-195">You can also check out the [Automated task reminders sample scenario](../resources/scenarios/task-reminders.md) to learn how to combine Office Scripts and Power Automate with Teams Adaptive Cards.</span></span>
