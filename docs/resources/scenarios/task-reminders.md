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
# <a name="office-scripts-sample-scenario-automated-task-reminders"></a>Office Сценарий пример сценария: Автоматизированные напоминания о задачах

В этом сценарии вы управляете проектом. Вы используете Excel лист для отслеживания статуса ваших сотрудников каждый месяц. Часто нужно напоминать людям, чтобы заполнить их статус, поэтому вы решили автоматизировать этот процесс напоминания.

Вы создадите поток сообщений Power Automate недостающих полях статусов и примените их ответы к электронной таблице. Для этого вы разработаете пару скриптов для работы с рабочей книгой. Первый скрипт получает список людей с пустыми состояниями, а второй скрипт добавляет строку статуса в правую строку. Вы также будете использовать эти Teams [карты, чтобы](/microsoftteams/platform/task-modules-and-cards/what-are-cards) сотрудники ввести свой статус непосредственно из уведомления.

## <a name="scripting-skills-covered"></a>Навыки сценариев охвачены

- Создание потоков в Power Automate
- Перейти данные к скриптам
- Возврат данных из скриптов
- Teams Адаптивные карты
- Таблицы

## <a name="prerequisites"></a>Предварительные требования

В этом сценарии [используются Power Automate](https://flow.microsoft.com) [и Microsoft Teams.](https://www.microsoft.com/microsoft-365/microsoft-teams/group-chat-software) Вам понадобится как связанный с учетной записью, которую вы используете для разработки Office скриптов. Для получения бесплатного доступа к подписке разработчика Майкрософт, чтобы узнать об этих приложениях и работать с [ними, рассмотрите возможность присоединения Microsoft 365 программе разработчиков.](https://developer.microsoft.com/microsoft-365/dev-program)

## <a name="setup-instructions"></a>Инструкции по настройке

1. Загрузите <a href="task-reminders.xlsx">task-reminders.xlsx</a> в свой OneDrive.

2. Откройте трудовую книжку в Excel в Интернете.

3. Под **вкладкой Автоматизировать,** открыть **все сценарии**.

4. Во-первых, нам нужен скрипт, чтобы получить все сотрудники с отчетами о состоянии, которые отсутствуют в таблице. В **панели задач редактора** кода нажмите **New Script и** вставьте следующий скрипт в редактор.

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

5. Сохраните скрипт с именем **Get People**.

6. Далее нам нужен второй скрипт для обработки карт отчета о состоянии и вовся новой информации в электронную таблицу. В **панели задач редактора** кода нажмите **New Script и** вставьте следующий скрипт в редактор.

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

7. Сохраните скрипт с именем **Сохранить статус**.

8. Теперь нам нужно создать поток. Открытый [Power Automate](https://flow.microsoft.com/).

    > [!TIP]
    > Если вы еще не создали поток раньше, пожалуйста, проверьте наш [учебник Начните с помощью скриптов Power Automate,](../../tutorials/excel-power-automate-manual.md) чтобы узнать основы.

9. Создайте новый **мгновенный поток.**

10. Выберите **Вручную вызвать поток из** вариантов и нажмите **Создать**.

11. Поток должен вызвать скрипт **Get People, чтобы** получить всех сотрудников с пустыми полями состояния. Нажмите **новый** шаг и **выберите Excel Интернет (Бизнес)**. Под **действия,** выберите **сценарий Run**. Предоставьте следующие записи для шага потока:

    - **Расположение**: OneDrive для бизнеса
    - **Библиотека документов**: OneDrive
    - **Файл**: task-reminders.xlsx *(Выбранный через файл браузера)*
    - **Сценарий**: Получить люди

    :::image type="content" source="../../images/scenario-task-reminders-first-flow-step.png" alt-text="Поток Power Automate, показывающий первый шаг потока скрипта Run":::

12. Далее поток должен обрабатывать каждого сотрудника в массиве, возвращенный скриптом. Нажмите **новый** шаг и **выберите Опубликовать адаптивную карту Teams пользователя и ждать ответа**.

13. Для поля **Получателя** добавьте **электронную** почту из динамического содержимого (выбор будет иметь Excel логотипом). Добавление **электронной** почты приводит к тому, что шаг потока окружен **применить к каждому блоку.** Это означает, что массив будет итерирован Power Automate.

14. Отправка адаптивной карты требует, чтобы JSON карты был предоставлен в качестве **Сообщения.** Вы можете использовать [адаптивную карту Дизайнер для](https://adaptivecards.io/designer/) создания пользовательских карт. Для этого образца используйте следующий JSON.  

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

15. Заполните оставшиеся поля следующим образом:

    - **Сообщение об** обновлении : Спасибо за отправку отчета о состоянии. Ваш ответ был успешно добавлен в электронную таблицу.
    - **Если обновить карту**: Да

16. В Применить **к каждому** блоку, после поста адаптивной карты для Teams пользователя и **ждать ответа,** **нажмите Добавить действие**. Выберите **Excel Интернет (Бизнес)**. Под **действия,** выберите **сценарий Run**. Предоставьте следующие записи для шага потока:

    - **Расположение**: OneDrive для бизнеса
    - **Библиотека документов**: OneDrive
    - **Файл**: task-reminders.xlsx *(Выбранный через файл браузера)*
    - **Сценарий**: Сохранить статус
    - **senderEmail**: электронная *почта (динамическое содержание от Excel)*
    - **statusReportResponse**: ответ *(динамический контент от Teams)*

    :::image type="content" source="../../images/scenario-task-reminders-last-flow-step.png" alt-text="Поток Power Automate, показывающий применение к каждому шагу":::

17. Сохранить поток.

## <a name="running-the-flow"></a>Запуск потока

Чтобы проверить поток, убедитесь, что любые строки таблицы с пустым статусом используют адрес электронной почты, привязанный к учетной записи Teams (вы, вероятно, должны использовать свой собственный адрес электронной почты во время тестирования).

Вы можете выбрать **тест** от конструктора потока или запустить поток со страницы **My flows.** После запуска потока и принятия использования необходимых соединений, вы должны получить адаптивную карту от Power Automate через Teams. Как только вы заполните поле состояния в карте, поток будет продолжаться и обновлять электронную таблицу со статусом, который вы предоставляете.

### <a name="before-running-the-flow"></a>Перед запуском потока

:::image type="content" source="../../images/scenario-task-reminders-spreadsheet-before.png" alt-text="Лист с отчетом о состоянии, содержащим одну недостающую запись статуса":::

### <a name="receiving-the-adaptive-card"></a>Получение адаптивной карты

:::image type="content" source="../../images/scenario-task-reminders-adaptive-card.png" alt-text="Адаптивная карта в Teams просит сотрудника об обновлении статуса":::

### <a name="after-running-the-flow"></a>После запуска потока

:::image type="content" source="../../images/scenario-task-reminders-spreadsheet-after.png" alt-text="Лист с отчетом о состоянии с заполненной в настоящее время записью статуса":::
