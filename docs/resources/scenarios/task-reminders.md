---
title: 'Пример сценария office Scripts: автоматические напоминания о задачах'
description: Пример, использующий Power Automate и Adaptive Cards, автоматизирует напоминания о задачах в таблице управления проектами.
ms.date: 11/30/2020
localization_priority: Normal
ms.openlocfilehash: a229a06e9f1f9118d57dadac8864bbc7eae7315b
ms.sourcegitcommit: 45ffe3dbd2c834b78592ad35928cf8096f5e80bc
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 04/14/2021
ms.locfileid: "51755156"
---
# <a name="office-scripts-sample-scenario-automated-task-reminders"></a><span data-ttu-id="3484c-103">Пример сценария office Scripts: автоматические напоминания о задачах</span><span class="sxs-lookup"><span data-stu-id="3484c-103">Office Scripts sample scenario: Automated task reminders</span></span>

<span data-ttu-id="3484c-104">В этом сценарии вы управляете проектом.</span><span class="sxs-lookup"><span data-stu-id="3484c-104">In this scenario you're managing a project.</span></span> <span data-ttu-id="3484c-105">Вы используете таблицу Excel для отслеживания состояния сотрудников каждый месяц.</span><span class="sxs-lookup"><span data-stu-id="3484c-105">You use an Excel worksheet to track your employees' status every month.</span></span> <span data-ttu-id="3484c-106">Часто необходимо напоминать людям о том, как заполнить их состояние, поэтому вы решили автоматизировать процесс напоминания.</span><span class="sxs-lookup"><span data-stu-id="3484c-106">You often need to remind people to fill out their status, so you've decided to automate that reminder process.</span></span>

<span data-ttu-id="3484c-107">Вы создайте поток Power Automate для сообщения людей с отсутствуют поля состояния и применить их ответы к таблице.</span><span class="sxs-lookup"><span data-stu-id="3484c-107">You'll create a Power Automate flow to message people with missing status fields and apply their responses to the spreadsheet.</span></span> <span data-ttu-id="3484c-108">Для этого вы разработает пару скриптов для обработки работы с книгой.</span><span class="sxs-lookup"><span data-stu-id="3484c-108">To do this, you'll develop a pair of scripts to handle the working with the workbook.</span></span> <span data-ttu-id="3484c-109">Первый скрипт получает список людей с пустыми состояниями, а второй сценарий добавляет строку состояния в правой строке.</span><span class="sxs-lookup"><span data-stu-id="3484c-109">The first script gets a list of people with blank statuses and the second script adds a status string to the right row.</span></span> <span data-ttu-id="3484c-110">Кроме того, с помощью [команд адаптивных](/microsoftteams/platform/task-modules-and-cards/what-are-cards) карт сотрудники введите свой статус непосредственно из уведомления.</span><span class="sxs-lookup"><span data-stu-id="3484c-110">You'll also make use of [Teams Adaptive Cards](/microsoftteams/platform/task-modules-and-cards/what-are-cards) to have employees enter their status directly from the notification.</span></span>

## <a name="scripting-skills-covered"></a><span data-ttu-id="3484c-111">Навыки скриптов, охватываемых</span><span class="sxs-lookup"><span data-stu-id="3484c-111">Scripting skills covered</span></span>

- <span data-ttu-id="3484c-112">Создание потоков в Power Automate</span><span class="sxs-lookup"><span data-stu-id="3484c-112">Create flows in Power Automate</span></span>
- <span data-ttu-id="3484c-113">Передачу данных скриптам</span><span class="sxs-lookup"><span data-stu-id="3484c-113">Pass data to scripts</span></span>
- <span data-ttu-id="3484c-114">Возвращение данных из скриптов</span><span class="sxs-lookup"><span data-stu-id="3484c-114">Return data from scripts</span></span>
- <span data-ttu-id="3484c-115">Адаптивные карты Teams</span><span class="sxs-lookup"><span data-stu-id="3484c-115">Teams Adaptive Cards</span></span>
- <span data-ttu-id="3484c-116">Таблицы</span><span class="sxs-lookup"><span data-stu-id="3484c-116">Tables</span></span>

## <a name="prerequisites"></a><span data-ttu-id="3484c-117">Предварительные условия</span><span class="sxs-lookup"><span data-stu-id="3484c-117">Prerequisites</span></span>

<span data-ttu-id="3484c-118">В этом сценарии [используются Power Automate](https://flow.microsoft.com) и [Microsoft Teams.](https://www.microsoft.com/microsoft-365/microsoft-teams/group-chat-software)</span><span class="sxs-lookup"><span data-stu-id="3484c-118">This scenario uses [Power Automate](https://flow.microsoft.com) and [Microsoft Teams](https://www.microsoft.com/microsoft-365/microsoft-teams/group-chat-software).</span></span> <span data-ttu-id="3484c-119">Вам потребуется как связанное с учетной записью, используемой для разработки сценариев Office.</span><span class="sxs-lookup"><span data-stu-id="3484c-119">You will need both associated with the account that you use for developing Office Scripts.</span></span> <span data-ttu-id="3484c-120">Чтобы получить бесплатный доступ к подписке microsoft Developer, чтобы узнать об этих приложениях и работать с ними, рассмотрите возможность присоединения к [программе разработчиков Microsoft 365.](https://developer.microsoft.com/microsoft-365/dev-program)</span><span class="sxs-lookup"><span data-stu-id="3484c-120">For free access to a Microsoft Developer subscription to learn about and work with these applications, consider joining the [Microsoft 365 Developer Program](https://developer.microsoft.com/microsoft-365/dev-program).</span></span>

## <a name="setup-instructions"></a><span data-ttu-id="3484c-121">Инструкции по настройке</span><span class="sxs-lookup"><span data-stu-id="3484c-121">Setup instructions</span></span>

1. <span data-ttu-id="3484c-122">Скачайте <a href="task-reminders.xlsx">task-reminders.xlsx</a> в OneDrive.</span><span class="sxs-lookup"><span data-stu-id="3484c-122">Download <a href="task-reminders.xlsx">task-reminders.xlsx</a> to your OneDrive.</span></span>

2. <span data-ttu-id="3484c-123">Откройте книгу в Excel в Интернете.</span><span class="sxs-lookup"><span data-stu-id="3484c-123">Open the workbook in Excel on the web.</span></span>

3. <span data-ttu-id="3484c-124">В **вкладке Automate** откройте **все скрипты.**</span><span class="sxs-lookup"><span data-stu-id="3484c-124">Under the **Automate** tab, open **All Scripts**.</span></span>

4. <span data-ttu-id="3484c-125">Сначала нам нужен сценарий для получения всех сотрудников с отчетами о состоянии, которые отсутствуют в таблице.</span><span class="sxs-lookup"><span data-stu-id="3484c-125">First, we need a script to get all the employees with status reports that are missing from the spreadsheet.</span></span> <span data-ttu-id="3484c-126">В области **задач редактора** кода нажмите **кнопку Новый скрипт** и вклеите следующий скрипт в редактор.</span><span class="sxs-lookup"><span data-stu-id="3484c-126">In the **Code Editor** task pane, press **New Script** and paste the following script into the editor.</span></span>

    ```TypeScript
    /**
     * This script looks for missing status reports in a project management table.
     *
     * @returns An array of Employee objects (containing their names and emails).
     */
    function main(workbook: ExcelScript.Workbook): Employee[] {
      // Get the first worksheet and the first table on that worksheet.
      let sheet = workbook.getFirstWorksheet()
      let table = sheet.getTables()[0];

      // Give the column indices names matching their expected content.
      const NAME_INDEX = 0;
      const EMAIL_INDEX = 1;
      const STATUS_REPORT_INDEX = 2;

      // Get the data for the whole table.
      let bodyRangeValues = table.getRangeBetweenHeaderAndTotal().getValues();

      // Create the array of Employee objects to return.
      let people: Employee[] = [];

      // Loop through the table and check each row for completion.
      for (let i = 0; i < bodyRangeValues.length; i++) {
        let row = bodyRangeValues[i];
        if (row[STATUS_REPORT_INDEX] === "") {
          // Save the email to return.
          people.push({ name: row[NAME_INDEX].toString(), email: row[EMAIL_INDEX].toString() });
        }
      }

      // Log the array to verify we're getting the right rows.
      console.log(people);

      // Return the array of Employees.
      return people;
    }

    /**
     * An interface representing an employee.
     * An array of Employees will be returned from the script
     * for the Power Automate flow.
     */
    interface Employee {
      name: string;
      email: string;
    }
    ```

5. <span data-ttu-id="3484c-127">Сохраните сценарий с именем **Get People**.</span><span class="sxs-lookup"><span data-stu-id="3484c-127">Save the script with the name **Get People**.</span></span>

6. <span data-ttu-id="3484c-128">Далее нам нужен второй скрипт для обработки карт отчетов о состоянии и вложения новых сведений в таблицу.</span><span class="sxs-lookup"><span data-stu-id="3484c-128">Next, we need a second script to process the status report cards and put the new information in the spreadsheet.</span></span> <span data-ttu-id="3484c-129">В области **задач редактора** кода нажмите **кнопку Новый скрипт** и вклеите следующий скрипт в редактор.</span><span class="sxs-lookup"><span data-stu-id="3484c-129">In the **Code Editor** task pane, press **New Script** and paste the following script into the editor.</span></span>

    ```TypeScript
    /**
     * This script applies the results of a Teams Adaptive Card about
     * a status update to a project management table.
     *
     * @param senderEmail - The email address of the employee updating their status.
     * @param statusReportResponse - The employee's status report.
     */
    function main(workbook: ExcelScript.Workbook,
      senderEmail: string,
      statusReportResponse: string) {

      // Get the first worksheet and the first table in that worksheet.
      let sheet = workbook.getFirstWorksheet();
      let table = sheet.getTables()[0];

      // Give the column indices names matching their expected content.
      const NAME_INDEX = 0;
      const EMAIL_INDEX = 1;
      const STATUS_REPORT_INDEX = 2;

      // Get the range and data for the whole table.
      let bodyRange = table.getRangeBetweenHeaderAndTotal();
      let tableRowCount = bodyRange.getRowCount();
      let bodyRangeValues = bodyRange.getValues();

      // Create a flag to denote success.
      let statusAdded = false;

      // Loop through the table and check each row for a matching email address.
      for (let i = 0; i < tableRowCount && !statusAdded; i++) {
        let row = bodyRangeValues[i];

        // Check if the row's email address matches.
        if (row[EMAIL_INDEX] === senderEmail) {
          // Add the Teams Adaptive Card response to the table.
          bodyRange.getCell(i, STATUS_REPORT_INDEX).setValues([
            [statusReportResponse]
          ]);
          statusAdded = true;
        }
      }

      // If successful, log the status update.
      if (statusAdded) {
        console.log(
          `Successfully added status report for ${senderEmail} containing: ${statusReportResponse}`
        );
      }
    }
    ```

7. <span data-ttu-id="3484c-130">Сохраните сценарий с именем **Сохранить состояние**.</span><span class="sxs-lookup"><span data-stu-id="3484c-130">Save the script with the name **Save Status**.</span></span>

8. <span data-ttu-id="3484c-131">Теперь необходимо создать поток.</span><span class="sxs-lookup"><span data-stu-id="3484c-131">Now, we need to create the flow.</span></span> <span data-ttu-id="3484c-132">Open [Power Automate](https://flow.microsoft.com/).</span><span class="sxs-lookup"><span data-stu-id="3484c-132">Open [Power Automate](https://flow.microsoft.com/).</span></span>

    > [!TIP]
    > <span data-ttu-id="3484c-133">Если вы еще не создали поток раньше, ознакомьтесь с нашим учебником Начните с помощью скриптов с [Power Automate,](../../tutorials/excel-power-automate-manual.md) чтобы узнать основы.</span><span class="sxs-lookup"><span data-stu-id="3484c-133">If you haven't created a flow before, please check out our tutorial [Start using scripts with Power Automate](../../tutorials/excel-power-automate-manual.md) to learn the basics.</span></span>

9. <span data-ttu-id="3484c-134">Создайте новый **мгновенный поток.**</span><span class="sxs-lookup"><span data-stu-id="3484c-134">Create a new **Instant flow**.</span></span>

10. <span data-ttu-id="3484c-135">Выберите **вручную вызвать поток** из параметров и нажмите **кнопку Создать**.</span><span class="sxs-lookup"><span data-stu-id="3484c-135">Choose **Manually trigger a flow** from the options and press **Create**.</span></span>

11. <span data-ttu-id="3484c-136">Потоку необходимо вызвать скрипт **Get People,** чтобы получить всех сотрудников с пустыми полями состояния.</span><span class="sxs-lookup"><span data-stu-id="3484c-136">The flow needs to call the **Get People** script to get all the employees with empty status fields.</span></span> <span data-ttu-id="3484c-137">Нажмите **кнопку Новый шаг** и выберите Excel Online **(Бизнес).**</span><span class="sxs-lookup"><span data-stu-id="3484c-137">Press **New step** and select **Excel Online (Business)**.</span></span> <span data-ttu-id="3484c-138">В разделе **Действия** выберите **Запуск сценария (предварительная версия)**.</span><span class="sxs-lookup"><span data-stu-id="3484c-138">Under **Actions**, select **Run script (preview)**.</span></span> <span data-ttu-id="3484c-139">Предоставление следующих записей для шага потока:</span><span class="sxs-lookup"><span data-stu-id="3484c-139">Provide the following entries for the flow step:</span></span>

    - <span data-ttu-id="3484c-140">**Расположение**: OneDrive для бизнеса</span><span class="sxs-lookup"><span data-stu-id="3484c-140">**Location**: OneDrive for Business</span></span>
    - <span data-ttu-id="3484c-141">**Библиотека документов**: OneDrive</span><span class="sxs-lookup"><span data-stu-id="3484c-141">**Document Library**: OneDrive</span></span>
    - <span data-ttu-id="3484c-142">**Файл**: task-reminders.xlsx *(выбранный через браузер файлов)*</span><span class="sxs-lookup"><span data-stu-id="3484c-142">**File**: task-reminders.xlsx *(Chosen through the file browser)*</span></span>
    - <span data-ttu-id="3484c-143">**Сценарий**: Get People</span><span class="sxs-lookup"><span data-stu-id="3484c-143">**Script**: Get People</span></span>

    :::image type="content" source="../../images/scenario-task-reminders-first-flow-step.png" alt-text="Поток Power Automate, показывающий первый шаг потока скрипта Run.":::

12. <span data-ttu-id="3484c-145">Далее поток должен обрабатывать каждого сотрудника в массиве, возвращаемом скриптом.</span><span class="sxs-lookup"><span data-stu-id="3484c-145">Next, the flow needs to process each Employee in the array returned by the script.</span></span> <span data-ttu-id="3484c-146">Нажмите **кнопку Новый шаг** и выберите сообщение адаптивной карты пользователю Teams и **дождись ответа.**</span><span class="sxs-lookup"><span data-stu-id="3484c-146">Press **New step** and select **Post an Adaptive Card to a Teams user and wait for a response**.</span></span>

13. <span data-ttu-id="3484c-147">В поле **Получатель** добавьте **электронную почту** из динамического контента (в выборе будет логотип Excel).</span><span class="sxs-lookup"><span data-stu-id="3484c-147">For the **Recipient** field, add **email** from the dynamic content (the selection will have the Excel logo by it).</span></span> <span data-ttu-id="3484c-148">Добавление **электронной** почты вызывает, что шаг потока будет окружен применить **к каждому блоку.**</span><span class="sxs-lookup"><span data-stu-id="3484c-148">Adding **email** causes the flow step to be surrounded by an **Apply to each** block.</span></span> <span data-ttu-id="3484c-149">Это означает, что массив будет итерирован с помощью Power Automate.</span><span class="sxs-lookup"><span data-stu-id="3484c-149">That means the array will be iterated over by Power Automate.</span></span>

14. <span data-ttu-id="3484c-150">Отправка адаптивной карты требует, чтобы JSON карты предоставлялись в качестве **сообщения.**</span><span class="sxs-lookup"><span data-stu-id="3484c-150">Sending an Adaptive Card requires the card's JSON to be provided as the **Message**.</span></span> <span data-ttu-id="3484c-151">Для создания пользовательских карт можно использовать [конструктор адаптивных](https://adaptivecards.io/designer/) карт.</span><span class="sxs-lookup"><span data-stu-id="3484c-151">You can use the [Adaptive Card Designer](https://adaptivecards.io/designer/) to create custom cards.</span></span> <span data-ttu-id="3484c-152">В этом примере используйте следующий JSON.</span><span class="sxs-lookup"><span data-stu-id="3484c-152">For this sample, use the following JSON.</span></span>  

    ```json
    {
      "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
      "type": "AdaptiveCard",
      "version": "1.0",
      "body": [
        {
          "type": "TextBlock",
          "size": "Medium",
          "weight": "Bolder",
          "text": "Update your Status Report"
        },
        {
          "type": "Image",
          "altText": "",
          "url": "https://i.imgur.com/f5RcuF3.png"
        },
        {
          "type": "TextBlock",
          "text": "This is a reminder to update your status report for this month's review. You can do so right here in this card, or by adding it directly to the spreadsheet.",
          "wrap": true
        },
        {
          "type": "Input.Text",
          "placeholder": "My status report for this month is...",
          "id": "response",
          "isMultiline": true
        }
      ],
      "actions": [
        {
          "type": "Action.Submit",
          "title": "Submit",
          "id": "submit"
        }
      ]
    }
    ```

15. <span data-ttu-id="3484c-153">Заполните оставшиеся поля следующим образом:</span><span class="sxs-lookup"><span data-stu-id="3484c-153">Fill out the remaining fields as follows:</span></span>

    - <span data-ttu-id="3484c-154">**Сообщение об обновлении.** Спасибо за отправку отчета о состоянии.</span><span class="sxs-lookup"><span data-stu-id="3484c-154">**Update message**: Thank you for submitting your status report.</span></span> <span data-ttu-id="3484c-155">Ваш ответ успешно добавлен в таблицу.</span><span class="sxs-lookup"><span data-stu-id="3484c-155">Your response has been successfully added to the spreadsheet.</span></span>
    - <span data-ttu-id="3484c-156">**Должна обновить карточку**: Да</span><span class="sxs-lookup"><span data-stu-id="3484c-156">**Should update card**: Yes</span></span>

16. <span data-ttu-id="3484c-157">В **пункте Применить** к каждому блоку после публикации адаптивной карты пользователю **Teams** и дождаться ответа, нажмите **кнопку Добавить действие**.</span><span class="sxs-lookup"><span data-stu-id="3484c-157">In the **Apply to each** block, following the **Post an Adaptive Card to a Teams user and wait for a response**, press **Add an action**.</span></span> <span data-ttu-id="3484c-158">Выберите **Excel Online (Бизнес).**</span><span class="sxs-lookup"><span data-stu-id="3484c-158">Select **Excel Online (Business)**.</span></span> <span data-ttu-id="3484c-159">В разделе **Действия** выберите **Запуск сценария (предварительная версия)**.</span><span class="sxs-lookup"><span data-stu-id="3484c-159">Under **Actions**, select **Run script (preview)**.</span></span> <span data-ttu-id="3484c-160">Предоставление следующих записей для шага потока:</span><span class="sxs-lookup"><span data-stu-id="3484c-160">Provide the following entries for the flow step:</span></span>

    - <span data-ttu-id="3484c-161">**Расположение**: OneDrive для бизнеса</span><span class="sxs-lookup"><span data-stu-id="3484c-161">**Location**: OneDrive for Business</span></span>
    - <span data-ttu-id="3484c-162">**Библиотека документов**: OneDrive</span><span class="sxs-lookup"><span data-stu-id="3484c-162">**Document Library**: OneDrive</span></span>
    - <span data-ttu-id="3484c-163">**Файл**: task-reminders.xlsx *(выбранный через браузер файлов)*</span><span class="sxs-lookup"><span data-stu-id="3484c-163">**File**: task-reminders.xlsx *(Chosen through the file browser)*</span></span>
    - <span data-ttu-id="3484c-164">**Сценарий:** сохранение состояния</span><span class="sxs-lookup"><span data-stu-id="3484c-164">**Script**: Save Status</span></span>
    - <span data-ttu-id="3484c-165">**senderEmail:** электронная *почта (динамическое содержимое из Excel)*</span><span class="sxs-lookup"><span data-stu-id="3484c-165">**senderEmail**: email *(dynamic content from Excel)*</span></span>
    - <span data-ttu-id="3484c-166">**statusReportResponse**: response *(динамический контент из Teams)*</span><span class="sxs-lookup"><span data-stu-id="3484c-166">**statusReportResponse**: response *(dynamic content from Teams)*</span></span>

    :::image type="content" source="../../images/scenario-task-reminders-last-flow-step.png" alt-text="Поток Power Automate, показывающий каждый шаг apply-to-each.":::

17. <span data-ttu-id="3484c-168">Сохраните поток.</span><span class="sxs-lookup"><span data-stu-id="3484c-168">Save the flow.</span></span>

## <a name="running-the-flow"></a><span data-ttu-id="3484c-169">Запуск потока</span><span class="sxs-lookup"><span data-stu-id="3484c-169">Running the flow</span></span>

<span data-ttu-id="3484c-170">Чтобы проверить поток, убедитесь, что в таблицах с пустым состоянием используется адрес электронной почты, привязанный к учетной записи Teams (возможно, при тестировании следует использовать собственный адрес электронной почты).</span><span class="sxs-lookup"><span data-stu-id="3484c-170">To test the flow, make sure any table rows with blank status use an email address tied to a Teams account (you should probably use your own email address while testing).</span></span>

<span data-ttu-id="3484c-171">Вы можете выбрать **Test** из конструктора потока или запустить поток со страницы **Мои потоки.**</span><span class="sxs-lookup"><span data-stu-id="3484c-171">You can either select **Test** from the flow designer, or run the flow from the **My flows** page.</span></span> <span data-ttu-id="3484c-172">После запуска потока и пользования требуемыми подключениями необходимо получить адаптивную карту от Power Automate до Teams.</span><span class="sxs-lookup"><span data-stu-id="3484c-172">After starting the flow and accepting the use of the required connections, you should receive an Adaptive Card from Power Automate through Teams.</span></span> <span data-ttu-id="3484c-173">После заполнения поля состояния в карточке поток будет продолжаться и обновлять таблицу со статусом, который вы предоставляете.</span><span class="sxs-lookup"><span data-stu-id="3484c-173">Once you fill out the status field in the card, the flow will continue and update the spreadsheet with the status you provide.</span></span>

### <a name="before-running-the-flow"></a><span data-ttu-id="3484c-174">Перед запуском потока</span><span class="sxs-lookup"><span data-stu-id="3484c-174">Before running the flow</span></span>

:::image type="content" source="../../images/scenario-task-reminders-spreadsheet-before.png" alt-text="Таблица с отчетом о состоянии, содержащим одну отсутствующую запись состояния.":::

### <a name="receiving-the-adaptive-card"></a><span data-ttu-id="3484c-176">Получение адаптивной карты</span><span class="sxs-lookup"><span data-stu-id="3484c-176">Receiving the Adaptive Card</span></span>

:::image type="content" source="../../images/scenario-task-reminders-adaptive-card.png" alt-text="Адаптивная карта в Teams, запрашиваемая сотрудником для обновления состояния.":::

### <a name="after-running-the-flow"></a><span data-ttu-id="3484c-178">После запуска потока</span><span class="sxs-lookup"><span data-stu-id="3484c-178">After running the flow</span></span>

:::image type="content" source="../../images/scenario-task-reminders-spreadsheet-after.png" alt-text="Таблица с отчетом о состоянии с записью состояния, заполненной в настоящее время.":::
