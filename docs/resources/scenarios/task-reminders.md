---
title: 'Office сценарии сценариев: автоматические напоминания о задачах'
description: Пример, в котором Power Automate и адаптивных карточек автоматизируют напоминания о задачах в электронной таблице управления проектами.
ms.date: 06/29/2021
ms.localizationpriority: medium
ms.openlocfilehash: 08f3713210e83162f86d38bc8eb33d76bf8a7288
ms.sourcegitcommit: dd01979d34b3499360d2f79a56f8a8f24f480eed
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 06/15/2022
ms.locfileid: "66088115"
---
# <a name="office-scripts-sample-scenario-automated-task-reminders"></a>Office сценарии сценариев: автоматические напоминания о задачах

В этом сценарии вы управляете проектом. Вы используете Excel для отслеживания состояния сотрудников каждый месяц. Часто необходимо напоминать людям о том, что нужно заполнить их состояние, поэтому вы решили автоматизировать этот процесс напоминания.

Вы создадите поток Power Automate сообщения людям с отсутствующими полями состояния и примените их ответы к электронной таблице. Для этого вы разработаете пару скриптов для обработки работы с книгой. Первый скрипт получает список людей с пустыми состояниями, а второй скрипт добавляет строку состояния в правую строку. Вы также будете использовать адаптивные [карточки Teams](/microsoftteams/platform/task-modules-and-cards/what-are-cards), чтобы сотрудники могли вводить свое состояние непосредственно из уведомления.

## <a name="scripting-skills-covered"></a>Рассматриваются навыки навыков на написание скриптов

- Создание потоков в Power Automate
- Передача данных в скрипты
- Возврат данных из скриптов
- Teams адаптивных карточек
- Таблицы

## <a name="prerequisites"></a>Предварительные требования

В этом [сценарии используются Power Automate](https://flow.microsoft.com) и [Microsoft Teams](https://www.microsoft.com/microsoft-365/microsoft-teams/group-chat-software). Вам потребуется связать их с учетной записью, используемой для разработки Office сценариев. Чтобы получить бесплатный доступ к подписке разработчика Майкрософт, чтобы узнать об этих приложениях и работать с этими приложениями, рассмотрите возможность присоединения к Microsoft 365 [разработчика](https://developer.microsoft.com/microsoft-365/dev-program).

## <a name="setup-instructions"></a>Инструкции по настройке

1. <a href="task-reminders.xlsx"> Скачайтеtask-reminders.xlsx</a> на OneDrive.

1. Откройте книгу в Excel в Интернете.

1. Во-первых, нам нужен сценарий для получения всех сотрудников с отчетами о состоянии, отсутствующих в электронной таблице. На **вкладке "Автоматизация** " выберите " **Новый скрипт** " и вставьте следующий скрипт в редактор.

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

1. Сохраните скрипт с именем **"Получить людей"**.

1. Далее нам нужен второй сценарий для обработки карт отчетов о состоянии и размещения новых сведений в электронной таблице. В области задач редактора кода **выберите "Новый** скрипт" и вставьте следующий скрипт в редактор.

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

1. Сохраните скрипт с именем **"Сохранить состояние"**.

1. Теперь необходимо создать поток. Откройте [Power Automate](https://flow.microsoft.com/).

    > [!TIP]
    > Если вы еще не создали поток, ознакомьтесь с нашим руководством. Начните [использовать](../../tutorials/excel-power-automate-manual.md) скрипты Power Automate, чтобы изучить основные сведения.

1. Создайте новый **мгновенный поток**.

1. Выберите **вручную активировать поток из** параметров и нажмите кнопку **"Создать"**.

1. Поток должен вызвать сценарий **get People** , чтобы получить всех сотрудников с пустыми полями состояния. Выберите **новый шаг**, а затем Excel **Online (Business)**. В разделе **Действия** выберите **Запуск скрипта**. Укажите следующие записи для шага потока:

    - **Расположение**: OneDrive для бизнеса
    - **Библиотека документов**: OneDrive
    - **Файл**: task-reminders.xlsx *(выбирается через браузер файлов)*
    - **Сценарий**: получение людей

    :::image type="content" source="../../images/scenario-task-reminders-first-flow-step.png" alt-text="Поток Power Automate, показывающий первый шаг потока выполнения скрипта.":::

1. Далее поток должен обработать каждый сотрудник в массиве, возвращаемом скриптом. Выберите **"Новый шаг**", а затем нажмите кнопку "Опубликовать адаптивную карточку **" Teams пользователя и дождитесь ответа**.

1. В поле **"Получатель**" добавьте **сообщение электронной** почты из динамического содержимого (выбранный фрагмент будет содержать Excel логотип). Добавление **электронной** почты приводит к тому, что шаг потока будет заключен в действие **"Применить к каждому блоку** ". Это означает, что массив будет переопределяться Power Automate.

1. Отправка адаптивной карточки требует, чтобы в качестве сообщения был указан [JSON](https://www.w3schools.com/whatis/whatis_json.asp) **карточки**. Конструктор адаптивных [карточек можно использовать](https://adaptivecards.io/designer/) для создания пользовательских карточек. В этом примере используйте следующий код JSON.  

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

1. Заполните остальные поля следующим образом:

    - **Сообщение об обновлении**: благодарим вас за отправку отчета о состоянии. Ваш ответ успешно добавлен в электронную таблицу.
    - **Карточка обновления**: Да

1. В разделе **"Применить к каждому** блоку" после публикации адаптивной карточки Teams пользователя и ожидания **ответа нажмите** кнопку **"Добавить действие"**. Выберите **Excel Online (Business)**. В разделе **Действия** выберите **Запуск скрипта**. Укажите следующие записи для шага потока:

    - **Расположение**: OneDrive для бизнеса
    - **Библиотека документов**: OneDrive
    - **Файл**: task-reminders.xlsx *(выбирается через браузер файлов)*
    - **Сценарий**: сохранение состояния
    - **senderEmail**: электронная *почта (динамическое содержимое из Excel)*
    - **statusReportResponse**: ответ *(динамическое содержимое из Teams)*

    :::image type="content" source="../../images/scenario-task-reminders-last-flow-step.png" alt-text="Поток Power Automate, в котором показан шаг &quot;Применить к каждому&quot;.":::

1. Сохраните поток.

## <a name="running-the-flow"></a>Выполнение потока

Чтобы протестировать поток, убедитесь, что все строки таблицы с пустым состоянием используют адрес электронной почты, привязанный к учетной записи Teams (возможно, при тестировании следует использовать собственный адрес электронной почты). Нажмите **кнопку "** Тест" на странице редактора потоков или запустите поток на **вкладке "Мои потоки** ". Не забудьте разрешить доступ при появлении запроса.

Вы должны получить адаптивную карточку от Power Automate до Teams. После заполнения поля состояния в карточке поток продолжит работу и обно службы электронной таблицы с указанным состоянием.

### <a name="before-running-the-flow"></a>Перед выполнением потока

:::image type="content" source="../../images/scenario-task-reminders-spreadsheet-before.png" alt-text="Лист с отчетом о состоянии, содержащим одну отсутствующую запись состояния.":::

### <a name="receiving-the-adaptive-card"></a>Получение адаптивной карточки

:::image type="content" source="../../images/scenario-task-reminders-adaptive-card.png" alt-text="Адаптивная карточка в Teams запрашивает у сотрудника обновление состояния.":::

### <a name="after-running-the-flow"></a>После выполнения потока

:::image type="content" source="../../images/scenario-task-reminders-spreadsheet-after.png" alt-text="Лист с отчетом о состоянии с заполняемой записью состояния.":::
