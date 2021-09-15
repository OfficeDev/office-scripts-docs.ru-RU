---
title: 'Office Пример сценария: автоматические напоминания о задачах'
description: Пример, использующий Power Automate и адаптивные карты, автоматизирует напоминания о задачах в таблице управления проектами.
ms.date: 06/29/2021
ms.localizationpriority: medium
ms.openlocfilehash: a2d4701fb7a42953de669c84dbb93104d199d5b8
ms.sourcegitcommit: d3ed4bdeeba805d97c930394e172e8306a0cf484
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/15/2021
ms.locfileid: "59330050"
---
# <a name="office-scripts-sample-scenario-automated-task-reminders"></a>Office Пример сценария: автоматические напоминания о задачах

В этом сценарии вы управляете проектом. Вы используете таблицу Excel для отслеживания состояния сотрудников каждый месяц. Часто необходимо напоминать людям о том, как заполнить их состояние, поэтому вы решили автоматизировать процесс напоминания.

Вы создайте поток Power Automate для сообщений людей с отсутствующих полями состояния и применить их ответы к таблице. Для этого вы разработает пару скриптов для обработки работы с книгой. Первый скрипт получает список людей с пустыми состояниями, а второй сценарий добавляет строку состояния в правой строке. Вы также будете использовать Teams [адаптивные](/microsoftteams/platform/task-modules-and-cards/what-are-cards) карты, чтобы сотрудники ввести свой статус непосредственно из уведомления.

## <a name="scripting-skills-covered"></a>Навыки скриптов, охватываемых

- Создание потоков в Power Automate
- Передачу данных скриптам
- Возвращение данных из скриптов
- Teams Адаптивные карты
- Таблицы

## <a name="prerequisites"></a>Предварительные условия

В этом [сценарии используются Power Automate](https://flow.microsoft.com) и [Microsoft Teams](https://www.microsoft.com/microsoft-365/microsoft-teams/group-chat-software). Вам потребуется как связанное с учетной записью, используемой для разработки Office скриптов. Чтобы получить бесплатный доступ к подписке microsoft Developer, чтобы узнать об этих приложениях и работать с ними, рассмотрите возможность присоединения [к программе Microsoft 365 разработчика.](https://developer.microsoft.com/microsoft-365/dev-program)

## <a name="setup-instructions"></a>Инструкции по настройке

1. Скачайте <a href="task-reminders.xlsx">task-reminders.xlsx</a> в OneDrive.

1. Откройте книгу в Excel в Интернете.

1. Сначала нам нужен сценарий для получения всех сотрудников с отчетами о состоянии, которые отсутствуют в таблице. В **вкладке Automate** выберите **Новый скрипт** и вклеите следующий скрипт в редактор.

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

1. Сохраните сценарий с именем **Get People**.

1. Далее нам нужен второй скрипт для обработки карт отчетов о состоянии и вложения новых сведений в таблицу. В области задач редактора кода выберите **Новый скрипт** и вклеите следующий скрипт в редактор.

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

1. Сохраните сценарий с именем **Сохранить состояние**.

1. Теперь необходимо создать поток. Откройте [Power Automate](https://flow.microsoft.com/).

    > [!TIP]
    > Если вы еще не создали поток раньше, ознакомьтесь с нашим учебником Начните использовать сценарии с Power Automate, [чтобы](../../tutorials/excel-power-automate-manual.md) изучить основы.

1. Создайте новый **мгновенный поток.**

1. Выберите **вручную вызвать поток** из параметров и выберите **Создать**.

1. Потоку необходимо вызвать скрипт **Get People,** чтобы получить всех сотрудников с пустыми полями состояния. Выберите **новый шаг,** а **затем выберите Excel Online (Бизнес).** В разделе **Действия** выберите **Запуск скрипта**. Предоставление следующих записей для шага потока:

    - **Расположение**: OneDrive для бизнеса
    - **Библиотека документов**: OneDrive
    - **Файл**: task-reminders.xlsx *(выбранный через браузер файлов)*
    - **Сценарий**: Get People

    :::image type="content" source="../../images/scenario-task-reminders-first-flow-step.png" alt-text="Поток Power Automate, показывающий первый шаг потока скрипта Run.":::

1. Далее поток должен обрабатывать каждого сотрудника в массиве, возвращаемом скриптом. Выберите **новый шаг,** затем выберите сообщение адаптивной карты Teams пользователю и **дождись ответа.**

1. В поле **Получатель** добавьте **электронную** почту из динамического контента (в выборе будет Excel логотип). Добавление **электронной** почты вызывает, что шаг потока будет окружен применить **к каждому блоку.** Это означает, что массив будет итерирован Power Automate.

1. Отправка адаптивной карты требует, чтобы JSON карты предоставлялись в качестве **сообщения.** Для создания пользовательских карт можно использовать [конструктор адаптивных](https://adaptivecards.io/designer/) карт. В этом примере используйте следующий JSON.  

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

1. Заполните оставшиеся поля следующим образом:

    - **Сообщение об обновлении.** Спасибо за отправку отчета о состоянии. Ваш ответ успешно добавлен в таблицу.
    - **Должна обновить карточку**: Да

1. В **пункте Применить** к каждому блоку после публикации адаптивной карты Teams пользователю и дождаться **ответа,** выберите **Добавить действие.** Выберите **Excel Online (Бизнес).** В разделе **Действия** выберите **Запуск скрипта**. Предоставление следующих записей для шага потока:

    - **Расположение**: OneDrive для бизнеса
    - **Библиотека документов**: OneDrive
    - **Файл**: task-reminders.xlsx *(выбранный через браузер файлов)*
    - **Сценарий:** сохранение состояния
    - **senderEmail:** email *(динамическое содержимое из Excel)*
    - **statusReportResponse**: response *(динамическое содержимое из Teams)*

    :::image type="content" source="../../images/scenario-task-reminders-last-flow-step.png" alt-text="Поток Power Automate, показывающий каждый шаг apply-to-each.":::

1. Сохраните поток.

## <a name="running-the-flow"></a>Запуск потока

Чтобы проверить поток, убедитесь, что в таблицах с пустым состоянием используется адрес электронной почты, привязанный к учетной записи Teams (возможно, при тестировании следует использовать собственный адрес электронной почты). Используйте **кнопку Test** на странице редактора потока или запустите поток через вкладку **Мои потоки.** Не забудьте разрешить доступ при запросе.

Вы должны получать адаптивную карту с Power Automate до Teams. После заполнения поля состояния в карточке поток будет продолжаться и обновлять таблицу со статусом, который вы предоставляете.

### <a name="before-running-the-flow"></a>Перед запуском потока

:::image type="content" source="../../images/scenario-task-reminders-spreadsheet-before.png" alt-text="Таблица с отчетом о состоянии, содержащим одну отсутствующую запись состояния.":::

### <a name="receiving-the-adaptive-card"></a>Получение адаптивной карты

:::image type="content" source="../../images/scenario-task-reminders-adaptive-card.png" alt-text="Адаптивная карта в Teams запрашивает у сотрудника обновление состояния.":::

### <a name="after-running-the-flow"></a>После запуска потока

:::image type="content" source="../../images/scenario-task-reminders-spreadsheet-after.png" alt-text="Таблица с отчетом о состоянии с записью состояния, заполненной в настоящее время.":::
