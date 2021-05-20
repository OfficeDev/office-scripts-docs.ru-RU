---
title: 'Office Сценарий пример сценария: Автоматизированные напоминания о задачах'
description: Образец, в Power Automate и адаптивных карт, автоматизирует напоминания о задачах в таблице управления проектами.
ms.date: 11/30/2020
localization_priority: Normal
ms.openlocfilehash: c254a627da8442c0974263908a41275182740b6e
ms.sourcegitcommit: 4687693f02fc90a57ba30c461f35046e02e6f5fb
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 05/19/2021
ms.locfileid: "52545610"
---
# <a name="office-scripts-sample-scenario-automated-task-reminders"></a><span data-ttu-id="a9142-103">Office Сценарий пример сценария: Автоматизированные напоминания о задачах</span><span class="sxs-lookup"><span data-stu-id="a9142-103">Office Scripts sample scenario: Automated task reminders</span></span>

<span data-ttu-id="a9142-104">В этом сценарии вы управляете проектом.</span><span class="sxs-lookup"><span data-stu-id="a9142-104">In this scenario you're managing a project.</span></span> <span data-ttu-id="a9142-105">Вы используете Excel лист для отслеживания статуса ваших сотрудников каждый месяц.</span><span class="sxs-lookup"><span data-stu-id="a9142-105">You use an Excel worksheet to track your employees' status every month.</span></span> <span data-ttu-id="a9142-106">Часто нужно напоминать людям, чтобы заполнить их статус, поэтому вы решили автоматизировать этот процесс напоминания.</span><span class="sxs-lookup"><span data-stu-id="a9142-106">You often need to remind people to fill out their status, so you've decided to automate that reminder process.</span></span>

<span data-ttu-id="a9142-107">Вы создадите поток сообщений Power Automate недостающих полях статусов и примените их ответы к электронной таблице.</span><span class="sxs-lookup"><span data-stu-id="a9142-107">You'll create a Power Automate flow to message people with missing status fields and apply their responses to the spreadsheet.</span></span> <span data-ttu-id="a9142-108">Для этого вы разработаете пару скриптов для работы с рабочей книгой.</span><span class="sxs-lookup"><span data-stu-id="a9142-108">To do this, you'll develop a pair of scripts to handle the working with the workbook.</span></span> <span data-ttu-id="a9142-109">Первый скрипт получает список людей с пустыми состояниями, а второй скрипт добавляет строку статуса в правую строку.</span><span class="sxs-lookup"><span data-stu-id="a9142-109">The first script gets a list of people with blank statuses and the second script adds a status string to the right row.</span></span> <span data-ttu-id="a9142-110">Вы также будете использовать эти Teams [карты, чтобы](/microsoftteams/platform/task-modules-and-cards/what-are-cards) сотрудники ввести свой статус непосредственно из уведомления.</span><span class="sxs-lookup"><span data-stu-id="a9142-110">You'll also make use of [Teams Adaptive Cards](/microsoftteams/platform/task-modules-and-cards/what-are-cards) to have employees enter their status directly from the notification.</span></span>

## <a name="scripting-skills-covered"></a><span data-ttu-id="a9142-111">Навыки сценариев охвачены</span><span class="sxs-lookup"><span data-stu-id="a9142-111">Scripting skills covered</span></span>

- <span data-ttu-id="a9142-112">Создание потоков в Power Automate</span><span class="sxs-lookup"><span data-stu-id="a9142-112">Create flows in Power Automate</span></span>
- <span data-ttu-id="a9142-113">Перейти данные к скриптам</span><span class="sxs-lookup"><span data-stu-id="a9142-113">Pass data to scripts</span></span>
- <span data-ttu-id="a9142-114">Возврат данных из скриптов</span><span class="sxs-lookup"><span data-stu-id="a9142-114">Return data from scripts</span></span>
- <span data-ttu-id="a9142-115">Teams Адаптивные карты</span><span class="sxs-lookup"><span data-stu-id="a9142-115">Teams Adaptive Cards</span></span>
- <span data-ttu-id="a9142-116">Таблицы</span><span class="sxs-lookup"><span data-stu-id="a9142-116">Tables</span></span>

## <a name="prerequisites"></a><span data-ttu-id="a9142-117">Предварительные требования</span><span class="sxs-lookup"><span data-stu-id="a9142-117">Prerequisites</span></span>

<span data-ttu-id="a9142-118">В этом сценарии [используются Power Automate](https://flow.microsoft.com) [и Microsoft Teams.](https://www.microsoft.com/microsoft-365/microsoft-teams/group-chat-software)</span><span class="sxs-lookup"><span data-stu-id="a9142-118">This scenario uses [Power Automate](https://flow.microsoft.com) and [Microsoft Teams](https://www.microsoft.com/microsoft-365/microsoft-teams/group-chat-software).</span></span> <span data-ttu-id="a9142-119">Вам понадобится как связанный с учетной записью, которую вы используете для разработки Office скриптов.</span><span class="sxs-lookup"><span data-stu-id="a9142-119">You will need both associated with the account that you use for developing Office Scripts.</span></span> <span data-ttu-id="a9142-120">Для получения бесплатного доступа к подписке разработчика Майкрософт, чтобы узнать об этих приложениях и работать с [ними, рассмотрите возможность присоединения Microsoft 365 программе разработчиков.](https://developer.microsoft.com/microsoft-365/dev-program)</span><span class="sxs-lookup"><span data-stu-id="a9142-120">For free access to a Microsoft Developer subscription to learn about and work with these applications, consider joining the [Microsoft 365 Developer Program](https://developer.microsoft.com/microsoft-365/dev-program).</span></span>

## <a name="setup-instructions"></a><span data-ttu-id="a9142-121">Инструкции по настройке</span><span class="sxs-lookup"><span data-stu-id="a9142-121">Setup instructions</span></span>

1. <span data-ttu-id="a9142-122">Загрузите <a href="task-reminders.xlsx">task-reminders.xlsx</a> в свой OneDrive.</span><span class="sxs-lookup"><span data-stu-id="a9142-122">Download <a href="task-reminders.xlsx">task-reminders.xlsx</a> to your OneDrive.</span></span>

2. <span data-ttu-id="a9142-123">Откройте трудовую книжку в Excel в Интернете.</span><span class="sxs-lookup"><span data-stu-id="a9142-123">Open the workbook in Excel on the web.</span></span>

3. <span data-ttu-id="a9142-124">Под **вкладкой Автоматизировать,** открыть **все сценарии**.</span><span class="sxs-lookup"><span data-stu-id="a9142-124">Under the **Automate** tab, open **All Scripts**.</span></span>

4. <span data-ttu-id="a9142-125">Во-первых, нам нужен скрипт, чтобы получить все сотрудники с отчетами о состоянии, которые отсутствуют в таблице.</span><span class="sxs-lookup"><span data-stu-id="a9142-125">First, we need a script to get all the employees with status reports that are missing from the spreadsheet.</span></span> <span data-ttu-id="a9142-126">В **панели задач редактора** кода нажмите **New Script и** вставьте следующий скрипт в редактор.</span><span class="sxs-lookup"><span data-stu-id="a9142-126">In the **Code Editor** task pane, press **New Script** and paste the following script into the editor.</span></span>

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

5. <span data-ttu-id="a9142-127">Сохраните скрипт с именем **Get People**.</span><span class="sxs-lookup"><span data-stu-id="a9142-127">Save the script with the name **Get People**.</span></span>

6. <span data-ttu-id="a9142-128">Далее нам нужен второй скрипт для обработки карт отчета о состоянии и вовся новой информации в электронную таблицу.</span><span class="sxs-lookup"><span data-stu-id="a9142-128">Next, we need a second script to process the status report cards and put the new information in the spreadsheet.</span></span> <span data-ttu-id="a9142-129">В **панели задач редактора** кода нажмите **New Script и** вставьте следующий скрипт в редактор.</span><span class="sxs-lookup"><span data-stu-id="a9142-129">In the **Code Editor** task pane, press **New Script** and paste the following script into the editor.</span></span>

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

7. <span data-ttu-id="a9142-130">Сохраните скрипт с именем **Сохранить статус**.</span><span class="sxs-lookup"><span data-stu-id="a9142-130">Save the script with the name **Save Status**.</span></span>

8. <span data-ttu-id="a9142-131">Теперь нам нужно создать поток.</span><span class="sxs-lookup"><span data-stu-id="a9142-131">Now, we need to create the flow.</span></span> <span data-ttu-id="a9142-132">Открытый [Power Automate](https://flow.microsoft.com/).</span><span class="sxs-lookup"><span data-stu-id="a9142-132">Open [Power Automate](https://flow.microsoft.com/).</span></span>

    > [!TIP]
    > <span data-ttu-id="a9142-133">Если вы еще не создали поток раньше, пожалуйста, проверьте наш [учебник Начните с помощью скриптов Power Automate,](../../tutorials/excel-power-automate-manual.md) чтобы узнать основы.</span><span class="sxs-lookup"><span data-stu-id="a9142-133">If you haven't created a flow before, please check out our tutorial [Start using scripts with Power Automate](../../tutorials/excel-power-automate-manual.md) to learn the basics.</span></span>

9. <span data-ttu-id="a9142-134">Создайте новый **мгновенный поток.**</span><span class="sxs-lookup"><span data-stu-id="a9142-134">Create a new **Instant flow**.</span></span>

10. <span data-ttu-id="a9142-135">Выберите **Вручную вызвать поток из** вариантов и нажмите **Создать**.</span><span class="sxs-lookup"><span data-stu-id="a9142-135">Choose **Manually trigger a flow** from the options and press **Create**.</span></span>

11. <span data-ttu-id="a9142-136">Поток должен вызвать скрипт **Get People, чтобы** получить всех сотрудников с пустыми полями состояния.</span><span class="sxs-lookup"><span data-stu-id="a9142-136">The flow needs to call the **Get People** script to get all the employees with empty status fields.</span></span> <span data-ttu-id="a9142-137">Нажмите **новый** шаг и **выберите Excel Интернет (Бизнес)**.</span><span class="sxs-lookup"><span data-stu-id="a9142-137">Press **New step** and select **Excel Online (Business)**.</span></span> <span data-ttu-id="a9142-138">Под **действия,** выберите **сценарий Run**.</span><span class="sxs-lookup"><span data-stu-id="a9142-138">Under **Actions**, select **Run script**.</span></span> <span data-ttu-id="a9142-139">Предоставьте следующие записи для шага потока:</span><span class="sxs-lookup"><span data-stu-id="a9142-139">Provide the following entries for the flow step:</span></span>

    - <span data-ttu-id="a9142-140">**Расположение**: OneDrive для бизнеса</span><span class="sxs-lookup"><span data-stu-id="a9142-140">**Location**: OneDrive for Business</span></span>
    - <span data-ttu-id="a9142-141">**Библиотека документов**: OneDrive</span><span class="sxs-lookup"><span data-stu-id="a9142-141">**Document Library**: OneDrive</span></span>
    - <span data-ttu-id="a9142-142">**Файл**: task-reminders.xlsx *(Выбранный через файл браузера)*</span><span class="sxs-lookup"><span data-stu-id="a9142-142">**File**: task-reminders.xlsx *(Chosen through the file browser)*</span></span>
    - <span data-ttu-id="a9142-143">**Сценарий**: Получить люди</span><span class="sxs-lookup"><span data-stu-id="a9142-143">**Script**: Get People</span></span>

    :::image type="content" source="../../images/scenario-task-reminders-first-flow-step.png" alt-text="Поток Power Automate, показывающий первый шаг потока скрипта Run":::

12. <span data-ttu-id="a9142-145">Далее поток должен обрабатывать каждого сотрудника в массиве, возвращенный скриптом.</span><span class="sxs-lookup"><span data-stu-id="a9142-145">Next, the flow needs to process each Employee in the array returned by the script.</span></span> <span data-ttu-id="a9142-146">Нажмите **новый** шаг и **выберите Опубликовать адаптивную карту Teams пользователя и ждать ответа**.</span><span class="sxs-lookup"><span data-stu-id="a9142-146">Press **New step** and select **Post an Adaptive Card to a Teams user and wait for a response**.</span></span>

13. <span data-ttu-id="a9142-147">Для поля **Получателя** добавьте **электронную** почту из динамического содержимого (выбор будет иметь Excel логотипом).</span><span class="sxs-lookup"><span data-stu-id="a9142-147">For the **Recipient** field, add **email** from the dynamic content (the selection will have the Excel logo by it).</span></span> <span data-ttu-id="a9142-148">Добавление **электронной** почты приводит к тому, что шаг потока окружен **применить к каждому блоку.**</span><span class="sxs-lookup"><span data-stu-id="a9142-148">Adding **email** causes the flow step to be surrounded by an **Apply to each** block.</span></span> <span data-ttu-id="a9142-149">Это означает, что массив будет итерирован Power Automate.</span><span class="sxs-lookup"><span data-stu-id="a9142-149">That means the array will be iterated over by Power Automate.</span></span>

14. <span data-ttu-id="a9142-150">Отправка адаптивной карты требует, чтобы JSON карты был предоставлен в качестве **Сообщения.**</span><span class="sxs-lookup"><span data-stu-id="a9142-150">Sending an Adaptive Card requires the card's JSON to be provided as the **Message**.</span></span> <span data-ttu-id="a9142-151">Вы можете использовать [адаптивную карту Дизайнер для](https://adaptivecards.io/designer/) создания пользовательских карт.</span><span class="sxs-lookup"><span data-stu-id="a9142-151">You can use the [Adaptive Card Designer](https://adaptivecards.io/designer/) to create custom cards.</span></span> <span data-ttu-id="a9142-152">Для этого образца используйте следующий JSON.</span><span class="sxs-lookup"><span data-stu-id="a9142-152">For this sample, use the following JSON.</span></span>  

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

15. <span data-ttu-id="a9142-153">Заполните оставшиеся поля следующим образом:</span><span class="sxs-lookup"><span data-stu-id="a9142-153">Fill out the remaining fields as follows:</span></span>

    - <span data-ttu-id="a9142-154">**Сообщение об** обновлении : Спасибо за отправку отчета о состоянии.</span><span class="sxs-lookup"><span data-stu-id="a9142-154">**Update message**: Thank you for submitting your status report.</span></span> <span data-ttu-id="a9142-155">Ваш ответ был успешно добавлен в электронную таблицу.</span><span class="sxs-lookup"><span data-stu-id="a9142-155">Your response has been successfully added to the spreadsheet.</span></span>
    - <span data-ttu-id="a9142-156">**Если обновить карту**: Да</span><span class="sxs-lookup"><span data-stu-id="a9142-156">**Should update card**: Yes</span></span>

16. <span data-ttu-id="a9142-157">В Применить **к каждому** блоку, после поста адаптивной карты для Teams пользователя и **ждать ответа,** **нажмите Добавить действие**.</span><span class="sxs-lookup"><span data-stu-id="a9142-157">In the **Apply to each** block, following the **Post an Adaptive Card to a Teams user and wait for a response**, press **Add an action**.</span></span> <span data-ttu-id="a9142-158">Выберите **Excel Интернет (Бизнес)**.</span><span class="sxs-lookup"><span data-stu-id="a9142-158">Select **Excel Online (Business)**.</span></span> <span data-ttu-id="a9142-159">Под **действия,** выберите **сценарий Run**.</span><span class="sxs-lookup"><span data-stu-id="a9142-159">Under **Actions**, select **Run script**.</span></span> <span data-ttu-id="a9142-160">Предоставьте следующие записи для шага потока:</span><span class="sxs-lookup"><span data-stu-id="a9142-160">Provide the following entries for the flow step:</span></span>

    - <span data-ttu-id="a9142-161">**Расположение**: OneDrive для бизнеса</span><span class="sxs-lookup"><span data-stu-id="a9142-161">**Location**: OneDrive for Business</span></span>
    - <span data-ttu-id="a9142-162">**Библиотека документов**: OneDrive</span><span class="sxs-lookup"><span data-stu-id="a9142-162">**Document Library**: OneDrive</span></span>
    - <span data-ttu-id="a9142-163">**Файл**: task-reminders.xlsx *(Выбранный через файл браузера)*</span><span class="sxs-lookup"><span data-stu-id="a9142-163">**File**: task-reminders.xlsx *(Chosen through the file browser)*</span></span>
    - <span data-ttu-id="a9142-164">**Сценарий**: Сохранить статус</span><span class="sxs-lookup"><span data-stu-id="a9142-164">**Script**: Save Status</span></span>
    - <span data-ttu-id="a9142-165">**senderEmail**: электронная *почта (динамическое содержание от Excel)*</span><span class="sxs-lookup"><span data-stu-id="a9142-165">**senderEmail**: email *(dynamic content from Excel)*</span></span>
    - <span data-ttu-id="a9142-166">**statusReportResponse**: ответ *(динамический контент от Teams)*</span><span class="sxs-lookup"><span data-stu-id="a9142-166">**statusReportResponse**: response *(dynamic content from Teams)*</span></span>

    :::image type="content" source="../../images/scenario-task-reminders-last-flow-step.png" alt-text="Поток Power Automate, показывающий применение к каждому шагу":::

17. <span data-ttu-id="a9142-168">Сохранить поток.</span><span class="sxs-lookup"><span data-stu-id="a9142-168">Save the flow.</span></span>

## <a name="running-the-flow"></a><span data-ttu-id="a9142-169">Запуск потока</span><span class="sxs-lookup"><span data-stu-id="a9142-169">Running the flow</span></span>

<span data-ttu-id="a9142-170">Чтобы проверить поток, убедитесь, что любые строки таблицы с пустым статусом используют адрес электронной почты, привязанный к учетной записи Teams (вы, вероятно, должны использовать свой собственный адрес электронной почты во время тестирования).</span><span class="sxs-lookup"><span data-stu-id="a9142-170">To test the flow, make sure any table rows with blank status use an email address tied to a Teams account (you should probably use your own email address while testing).</span></span>

<span data-ttu-id="a9142-171">Вы можете выбрать **тест** от конструктора потока или запустить поток со страницы **My flows.**</span><span class="sxs-lookup"><span data-stu-id="a9142-171">You can either select **Test** from the flow designer, or run the flow from the **My flows** page.</span></span> <span data-ttu-id="a9142-172">После запуска потока и принятия использования необходимых соединений, вы должны получить адаптивную карту от Power Automate через Teams.</span><span class="sxs-lookup"><span data-stu-id="a9142-172">After starting the flow and accepting the use of the required connections, you should receive an Adaptive Card from Power Automate through Teams.</span></span> <span data-ttu-id="a9142-173">Как только вы заполните поле состояния в карте, поток будет продолжаться и обновлять электронную таблицу со статусом, который вы предоставляете.</span><span class="sxs-lookup"><span data-stu-id="a9142-173">Once you fill out the status field in the card, the flow will continue and update the spreadsheet with the status you provide.</span></span>

### <a name="before-running-the-flow"></a><span data-ttu-id="a9142-174">Перед запуском потока</span><span class="sxs-lookup"><span data-stu-id="a9142-174">Before running the flow</span></span>

:::image type="content" source="../../images/scenario-task-reminders-spreadsheet-before.png" alt-text="Лист с отчетом о состоянии, содержащим одну недостающую запись статуса":::

### <a name="receiving-the-adaptive-card"></a><span data-ttu-id="a9142-176">Получение адаптивной карты</span><span class="sxs-lookup"><span data-stu-id="a9142-176">Receiving the Adaptive Card</span></span>

:::image type="content" source="../../images/scenario-task-reminders-adaptive-card.png" alt-text="Адаптивная карта в Teams просит сотрудника об обновлении статуса":::

### <a name="after-running-the-flow"></a><span data-ttu-id="a9142-178">После запуска потока</span><span class="sxs-lookup"><span data-stu-id="a9142-178">After running the flow</span></span>

:::image type="content" source="../../images/scenario-task-reminders-spreadsheet-after.png" alt-text="Лист с отчетом о состоянии с заполненной в настоящее время записью статуса":::
