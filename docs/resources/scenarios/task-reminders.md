---
title: 'Сценарий примера сценариев Office: автоматические напоминания о задачах'
description: Пример, в котором используются автоматизированные и адаптивные карточки, автоматизирующие напоминания о задачах в электронной таблице управления проектом.
ms.date: 11/30/2020
localization_priority: Normal
ms.openlocfilehash: 90769eca0201e450e25778db0eb5c62284b9feb0
ms.sourcegitcommit: af487756dffea0f8f0cd62710c586842cb08073c
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 12/04/2020
ms.locfileid: "49571453"
---
# <a name="office-scripts-sample-scenario-automated-task-reminders"></a><span data-ttu-id="b4bef-103">Сценарий примера сценариев Office: автоматические напоминания о задачах</span><span class="sxs-lookup"><span data-stu-id="b4bef-103">Office Scripts sample scenario: Automated task reminders</span></span>

<span data-ttu-id="b4bef-104">В этом сценарии управление проектом осуществляется с помощью.</span><span class="sxs-lookup"><span data-stu-id="b4bef-104">In this scenario you're managing a project.</span></span> <span data-ttu-id="b4bef-105">Лист Excel используется для отслеживания состояния сотрудников каждый месяц.</span><span class="sxs-lookup"><span data-stu-id="b4bef-105">You use an Excel worksheet to track your employees' status every month.</span></span> <span data-ttu-id="b4bef-106">Часто приходится напоминать пользователям о своем состоянии, поэтому вы решили автоматизировать этот процесс напоминания.</span><span class="sxs-lookup"><span data-stu-id="b4bef-106">You often need to remind people to fill out their status, so you've decided to automate that reminder process.</span></span>

<span data-ttu-id="b4bef-107">Вы создадите блок автоматизации Power, чтобы отправлять сообщения людям с отсутствующими состояниями и применить их ответы к электронной таблице.</span><span class="sxs-lookup"><span data-stu-id="b4bef-107">You'll create a Power Automate flow to message people with missing status fields and apply their responses to the spreadsheet.</span></span> <span data-ttu-id="b4bef-108">Для этого вы разрабатываете сочетание сценариев для работы с книгой.</span><span class="sxs-lookup"><span data-stu-id="b4bef-108">To do this, you'll develop a pair of scripts to handle the working with the workbook.</span></span> <span data-ttu-id="b4bef-109">Первый сценарий получает список людей с пустыми состояниями, а второй добавляет строку состояния в правую строку.</span><span class="sxs-lookup"><span data-stu-id="b4bef-109">The first script gets a list of people with blank statuses and the second script adds a status string to the right row.</span></span> <span data-ttu-id="b4bef-110">Вы также можете использовать [адаптивные карты Teams](/microsoftteams/platform/task-modules-and-cards/what-are-cards) , чтобы сотрудники вводили свое состояние непосредственно из уведомления.</span><span class="sxs-lookup"><span data-stu-id="b4bef-110">You'll also make use of [Teams Adaptive Cards](/microsoftteams/platform/task-modules-and-cards/what-are-cards) to have employees enter their status directly from the notification.</span></span>

## <a name="scripting-skills-covered"></a><span data-ttu-id="b4bef-111">Охваченные навыки работы со сценариями</span><span class="sxs-lookup"><span data-stu-id="b4bef-111">Scripting skills covered</span></span>

- <span data-ttu-id="b4bef-112">Создание потоков в автоматизации управления питанием</span><span class="sxs-lookup"><span data-stu-id="b4bef-112">Create flows in Power Automate</span></span>
- <span data-ttu-id="b4bef-113">Передача данных в скрипты</span><span class="sxs-lookup"><span data-stu-id="b4bef-113">Pass data to scripts</span></span>
- <span data-ttu-id="b4bef-114">Возвращение данных из скриптов</span><span class="sxs-lookup"><span data-stu-id="b4bef-114">Return data from scripts</span></span>
- <span data-ttu-id="b4bef-115">Адаптивные карты Teams</span><span class="sxs-lookup"><span data-stu-id="b4bef-115">Teams Adaptive Cards</span></span>
- <span data-ttu-id="b4bef-116">Таблицы</span><span class="sxs-lookup"><span data-stu-id="b4bef-116">Tables</span></span>

## <a name="prerequisites"></a><span data-ttu-id="b4bef-117">Необходимые компоненты</span><span class="sxs-lookup"><span data-stu-id="b4bef-117">Prerequisites</span></span>

<span data-ttu-id="b4bef-118">В этом сценарии используется [Power автоматизировать](https://flow.microsoft.com) и [Microsoft Teams](https://www.microsoft.com/microsoft-365/microsoft-teams/group-chat-software).</span><span class="sxs-lookup"><span data-stu-id="b4bef-118">This scenario uses [Power Automate](https://flow.microsoft.com) and [Microsoft Teams](https://www.microsoft.com/microsoft-365/microsoft-teams/group-chat-software).</span></span> <span data-ttu-id="b4bef-119">Вам потребуется связать с учетной записью, используемой для разработки сценариев Office.</span><span class="sxs-lookup"><span data-stu-id="b4bef-119">You will need both associated with the account that you use for developing Office Scripts.</span></span> <span data-ttu-id="b4bef-120">Чтобы получить бесплатный доступ к подписке Майкрософт для изучения и работы с этими приложениями, рассмотрите возможность присоединения к [программе для разработчиков microsoft 365](https://developer.microsoft.com/microsoft-365/dev-program).</span><span class="sxs-lookup"><span data-stu-id="b4bef-120">For free access to a Microsoft Developer subscription to learn about and work with these applications, consider joining the [Microsoft 365 Developer Program](https://developer.microsoft.com/microsoft-365/dev-program).</span></span>

## <a name="setup-instructions"></a><span data-ttu-id="b4bef-121">Инструкции по настройке</span><span class="sxs-lookup"><span data-stu-id="b4bef-121">Setup instructions</span></span>

1. <span data-ttu-id="b4bef-122">Скачайте <a href="task-reminders.xlsx">task-reminders.xlsx</a> в OneDrive.</span><span class="sxs-lookup"><span data-stu-id="b4bef-122">Download <a href="task-reminders.xlsx">task-reminders.xlsx</a> to your OneDrive.</span></span>

2. <span data-ttu-id="b4bef-123">Откройте книгу в Excel в Интернете.</span><span class="sxs-lookup"><span data-stu-id="b4bef-123">Open the workbook in Excel on the web.</span></span>

3. <span data-ttu-id="b4bef-124">На вкладке **Автоматизация** откройте **Редактор кода**.</span><span class="sxs-lookup"><span data-stu-id="b4bef-124">Under the **Automate** tab, open the **Code Editor**.</span></span>

4. <span data-ttu-id="b4bef-125">Для начала нам нужен сценарий для получения всех сотрудников с отчетами о состоянии, отсутствующими в электронной таблице.</span><span class="sxs-lookup"><span data-stu-id="b4bef-125">First, we need a script to get all the employees with status reports that are missing from the spreadsheet.</span></span> <span data-ttu-id="b4bef-126">В области задач **Редактор кода** нажмите **новый скрипт** и вставьте следующий скрипт в редактор.</span><span class="sxs-lookup"><span data-stu-id="b4bef-126">In the **Code Editor** task pane, press **New Script** and paste the following script into the editor.</span></span>

    ```typescript
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

5. <span data-ttu-id="b4bef-127">Сохраните сценарий с именем **получить людей**.</span><span class="sxs-lookup"><span data-stu-id="b4bef-127">Save the script with the name **Get People**.</span></span>

6. <span data-ttu-id="b4bef-128">Далее нам нужен второй скрипт для обработки карточек отчетов о состоянии и размещения новых данных в электронной таблице.</span><span class="sxs-lookup"><span data-stu-id="b4bef-128">Next, we need a second script to process the status report cards and put the new information in the spreadsheet.</span></span> <span data-ttu-id="b4bef-129">В области задач **Редактор кода** нажмите **новый скрипт** и вставьте следующий скрипт в редактор.</span><span class="sxs-lookup"><span data-stu-id="b4bef-129">In the **Code Editor** task pane, press **New Script** and paste the following script into the editor.</span></span>

    ```typescript
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

7. <span data-ttu-id="b4bef-130">Сохраните скрипт с именем **Сохранение состояния**.</span><span class="sxs-lookup"><span data-stu-id="b4bef-130">Save the script with the name **Save Status**.</span></span>

8. <span data-ttu-id="b4bef-131">Теперь необходимо создать блок.</span><span class="sxs-lookup"><span data-stu-id="b4bef-131">Now, we need to create the flow.</span></span> <span data-ttu-id="b4bef-132">Откройте компонент " [Автоматизация Power](https://flow.microsoft.com/)".</span><span class="sxs-lookup"><span data-stu-id="b4bef-132">Open [Power Automate](https://flow.microsoft.com/).</span></span>

    > [!TIP]
    > <span data-ttu-id="b4bef-133">Если вы еще не создали потоки, ознакомьтесь со статьей начало работы с [помощью сценариев в Power Автоматизация](../../tutorials/excel-power-automate-manual.md) , чтобы изучить основные принципы.</span><span class="sxs-lookup"><span data-stu-id="b4bef-133">If you haven't created a flow before, please check out our tutorial [Start using scripts with Power Automate](../../tutorials/excel-power-automate-manual.md) to learn the basics.</span></span>

9. <span data-ttu-id="b4bef-134">Создайте новый **мгновенный процесс**.</span><span class="sxs-lookup"><span data-stu-id="b4bef-134">Create a new **Instant flow**.</span></span>

10. <span data-ttu-id="b4bef-135">Выберите **вручную активировать потоки** из параметров и нажмите кнопку **создать**.</span><span class="sxs-lookup"><span data-stu-id="b4bef-135">Choose **Manually trigger a flow** from the options and press **Create**.</span></span>

11. <span data-ttu-id="b4bef-136">Для этого необходимо вызвать сценарий **Get люди** , чтобы получить все сотрудники с пустыми полями состояния.</span><span class="sxs-lookup"><span data-stu-id="b4bef-136">The flow needs to call the **Get People** script to get all the employees with empty status fields.</span></span> <span data-ttu-id="b4bef-137">Нажмите кнопку **создать шаг** и выберите **Excel Online (бизнес)**.</span><span class="sxs-lookup"><span data-stu-id="b4bef-137">Press **New step** and select **Excel Online (Business)**.</span></span> <span data-ttu-id="b4bef-138">В разделе **Действия** выберите **Запустить сценарий (предварительная версия)**.</span><span class="sxs-lookup"><span data-stu-id="b4bef-138">Under **Actions**, select **Run script (preview)**.</span></span> <span data-ttu-id="b4bef-139">Введите следующие записи для шага процесса:</span><span class="sxs-lookup"><span data-stu-id="b4bef-139">Provide the following entries for the flow step:</span></span>

    - <span data-ttu-id="b4bef-140">**Расположение**: OneDrive для бизнеса</span><span class="sxs-lookup"><span data-stu-id="b4bef-140">**Location**: OneDrive for Business</span></span>
    - <span data-ttu-id="b4bef-141">**Библиотека документов**: OneDrive</span><span class="sxs-lookup"><span data-stu-id="b4bef-141">**Document Library**: OneDrive</span></span>
    - <span data-ttu-id="b4bef-142">**Файл**: task-reminders.xlsx *(выбрано через браузер файлов)*</span><span class="sxs-lookup"><span data-stu-id="b4bef-142">**File**: task-reminders.xlsx *(Chosen through the file browser)*</span></span>
    - <span data-ttu-id="b4bef-143">**Сценарий**: получение людей</span><span class="sxs-lookup"><span data-stu-id="b4bef-143">**Script**: Get People</span></span>

    ![Шаг сценария первого запуска.](../../images/scenario-task-reminders-first-flow-step.png)

12. <span data-ttu-id="b4bef-145">Затем необходимо обработать каждого сотрудника в массиве, возвращенном сценарием.</span><span class="sxs-lookup"><span data-stu-id="b4bef-145">Next, the flow needs to process each Employee in the array returned by the script.</span></span> <span data-ttu-id="b4bef-146">Нажмите кнопку **создать шаг** и выберите **Отправить адаптивную карту пользователю Teams и дождитесь ответа**.</span><span class="sxs-lookup"><span data-stu-id="b4bef-146">Press **New step** and select **Post an Adaptive Card to a Teams user and wait for a response**.</span></span>

13. <span data-ttu-id="b4bef-147">В поле **получатель** добавьте **электронную почту** из динамического содержимого (выделенный фрагмент будет содержать логотип Excel).</span><span class="sxs-lookup"><span data-stu-id="b4bef-147">For the **Recipient** field, add **email** from the dynamic content (the selection will have the Excel logo by it).</span></span> <span data-ttu-id="b4bef-148">Добавление **электронной почты** приводит к тому, что этап процесса будет отключаться от **применения к каждому** блоку.</span><span class="sxs-lookup"><span data-stu-id="b4bef-148">Adding **email** causes the flow step to be surrounded by an **Apply to each** block.</span></span> <span data-ttu-id="b4bef-149">Это означает, что массив будет перебираться по автоматизации управления питанием.</span><span class="sxs-lookup"><span data-stu-id="b4bef-149">That means the array will be iterated over by Power Automate.</span></span>

14. <span data-ttu-id="b4bef-150">Для отправки адаптивной карточки необходимо, чтобы в качестве **сообщения** было предоставлено значение JSON карты.</span><span class="sxs-lookup"><span data-stu-id="b4bef-150">Sending an Adaptive Card requires the card's JSON to be provided as the **Message**.</span></span> <span data-ttu-id="b4bef-151">Для создания настраиваемых карточек можно использовать [адаптивный конструктор карточек](https://adaptivecards.io/designer/) .</span><span class="sxs-lookup"><span data-stu-id="b4bef-151">You can use the [Adaptive Card Designer](https://adaptivecards.io/designer/) to create custom cards.</span></span> <span data-ttu-id="b4bef-152">Для этого примера используйте приведенный ниже код JSON.</span><span class="sxs-lookup"><span data-stu-id="b4bef-152">For this sample, use the following JSON.</span></span>  

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

15. <span data-ttu-id="b4bef-153">Заполните оставшиеся поля следующим образом:</span><span class="sxs-lookup"><span data-stu-id="b4bef-153">Fill out the remaining fields as follows:</span></span>

    - <span data-ttu-id="b4bef-154">**Сообщение об обновлении**: Благодарим вас за отправку отчета о состоянии.</span><span class="sxs-lookup"><span data-stu-id="b4bef-154">**Update message**: Thank you for submitting your status report.</span></span> <span data-ttu-id="b4bef-155">Ваш ответ успешно добавлен в электронную таблицу.</span><span class="sxs-lookup"><span data-stu-id="b4bef-155">Your response has been successfully added to the spreadsheet.</span></span>
    - <span data-ttu-id="b4bef-156">**Обновите карточку**: Да</span><span class="sxs-lookup"><span data-stu-id="b4bef-156">**Should update card**: Yes</span></span>

16. <span data-ttu-id="b4bef-157">В разделе **Применить к каждому** блоку **после отправки адаптивной карты пользователю Teams и ожидания ответа** нажмите **Добавить действие**.</span><span class="sxs-lookup"><span data-stu-id="b4bef-157">In the **Apply to each** block, following the **Post an Adaptive Card to a Teams user and wait for a response**, press **Add an action**.</span></span> <span data-ttu-id="b4bef-158">Выберите **Excel Online (бизнес)**.</span><span class="sxs-lookup"><span data-stu-id="b4bef-158">Select **Excel Online (Business)**.</span></span> <span data-ttu-id="b4bef-159">В разделе **Действия** выберите **Запустить сценарий (предварительная версия)**.</span><span class="sxs-lookup"><span data-stu-id="b4bef-159">Under **Actions**, select **Run script (preview)**.</span></span> <span data-ttu-id="b4bef-160">Введите следующие записи для шага процесса:</span><span class="sxs-lookup"><span data-stu-id="b4bef-160">Provide the following entries for the flow step:</span></span>

    - <span data-ttu-id="b4bef-161">**Расположение**: OneDrive для бизнеса</span><span class="sxs-lookup"><span data-stu-id="b4bef-161">**Location**: OneDrive for Business</span></span>
    - <span data-ttu-id="b4bef-162">**Библиотека документов**: OneDrive</span><span class="sxs-lookup"><span data-stu-id="b4bef-162">**Document Library**: OneDrive</span></span>
    - <span data-ttu-id="b4bef-163">**Файл**: task-reminders.xlsx *(выбрано через браузер файлов)*</span><span class="sxs-lookup"><span data-stu-id="b4bef-163">**File**: task-reminders.xlsx *(Chosen through the file browser)*</span></span>
    - <span data-ttu-id="b4bef-164">**Сценарий**: сохранение состояния</span><span class="sxs-lookup"><span data-stu-id="b4bef-164">**Script**: Save Status</span></span>
    - <span data-ttu-id="b4bef-165">**сендеремаил**: Электронная почта *(динамическое содержимое из Excel)*</span><span class="sxs-lookup"><span data-stu-id="b4bef-165">**senderEmail**: email *(dynamic content from Excel)*</span></span>
    - <span data-ttu-id="b4bef-166">**статусрепортреспонсе**: отклик *(динамический контент из Teams)*</span><span class="sxs-lookup"><span data-stu-id="b4bef-166">**statusReportResponse**: response *(dynamic content from Teams)*</span></span>

    ![Шаг "Применить ко всем процессам".](../../images/scenario-task-reminders-last-flow-step.png)

17. <span data-ttu-id="b4bef-168">Сохраните ход выполнения.</span><span class="sxs-lookup"><span data-stu-id="b4bef-168">Save the flow.</span></span>

## <a name="running-the-flow"></a><span data-ttu-id="b4bef-169">Выполнение процесса</span><span class="sxs-lookup"><span data-stu-id="b4bef-169">Running the flow</span></span>

<span data-ttu-id="b4bef-170">Чтобы протестировать потоки, убедитесь, что все строки таблицы с пустым состоянием используют адрес электронной почты, связанный с учетной записью Teams (при тестировании вам, возможно, потребуется использовать свой собственный адрес электронной почты).</span><span class="sxs-lookup"><span data-stu-id="b4bef-170">To test the flow, make sure any table rows with blank status use an email address tied to a Teams account (you should probably use your own email address while testing).</span></span>

<span data-ttu-id="b4bef-171">Вы можете выбрать **тест** из конструктора потоков или запустить поток из страницы " **мои потоки** ".</span><span class="sxs-lookup"><span data-stu-id="b4bef-171">You can either select **Test** from the flow designer, or run the flow from the **My flows** page.</span></span> <span data-ttu-id="b4bef-172">После запуска процесса и приема использования необходимых подключений необходимо получить адаптивную карту Power автоматизировать через Teams.</span><span class="sxs-lookup"><span data-stu-id="b4bef-172">After starting the flow and accepting the use of the required connections, you should receive an Adaptive Card from Power Automate through Teams.</span></span> <span data-ttu-id="b4bef-173">Когда вы заполните поле Status на карточке, этот процесс продолжится и обновит электронную таблицу, указав предоставленное вами состояние.</span><span class="sxs-lookup"><span data-stu-id="b4bef-173">Once you fill out the status field in the card, the flow will continue and update the spreadsheet with the status you provide.</span></span>

### <a name="before-running-the-flow"></a><span data-ttu-id="b4bef-174">Перед запуском процесса</span><span class="sxs-lookup"><span data-stu-id="b4bef-174">Before running the flow</span></span>

![Лист с отчетом о состоянии, содержащий одну отсутствующую запись о состоянии.](../../images/scenario-task-reminders-spreadsheet-before.png)

### <a name="receiving-the-adaptive-card"></a><span data-ttu-id="b4bef-176">Получение адаптивной карточки</span><span class="sxs-lookup"><span data-stu-id="b4bef-176">Receiving the Adaptive Card</span></span>

![Адаптивная карта в Teams запрашивает у сотрудника обновление состояния.](../../images/scenario-task-reminders-adaptive-card.png)

### <a name="after-running-the-flow"></a><span data-ttu-id="b4bef-178">После выполнения последовательности</span><span class="sxs-lookup"><span data-stu-id="b4bef-178">After running the flow</span></span>

![Лист с отчетом о состоянии с записью состояния, заполненной теперь.](../../images/scenario-task-reminders-spreadsheet-after.png)
