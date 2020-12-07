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
# <a name="pass-data-to-scripts-in-an-automatically-run-power-automate-flow-preview"></a>Передача данных сценариям в автоматически запускаемых рабочих процессах Power Automate (предварительная версия)

В этом руководстве объясняется, как использовать сценарий Office для Excel в Интернете с помощью автоматизированных рабочих процессов [Power Automate](https://flow.microsoft.com). Сценарий будет автоматически выполняться каждый раз при получении электронной почты. Данные из сообщений электронной почты будут записываться в книгу Excel. Возможность передавать данные из других приложений в сценарии Office предоставляет вам значительную гибкость и свободу в автоматизированных процессах.

> [!TIP]
> Если вы только приступили к работе со сценариями Office, рекомендуем начать с учебника [Запись, редактирование и создание сценариев Office в Excel в Интернете](excel-tutorial.md). Если вы впервые используете Power Automate, рекомендуем начать с учебника [Вызов сценариев из активированного вручную потока Power Automate](excel-power-automate-manual.md). [Сценарии Office используют TypeScript](../overview/code-editor-environment.md), и этот учебник предназначен для пользователей с начальным и средним уровнем знаний по JavaScript или TypeScript. Если вы впервые работаете с JavaScript, рекомендуем начать с [учебника Mozilla по JavaScript](https://developer.mozilla.org/docs/Web/JavaScript/Guide/Introduction).

## <a name="prerequisites"></a>Предварительные условия

[!INCLUDE [Tutorial prerequisites](../includes/power-automate-tutorial-prerequisites.md)]

## <a name="prepare-the-workbook"></a>Подготовка книги

Power Automate не может использовать [относительные ссылки](../develop/power-automate-integration.md#avoid-using-relative-references), такие как `Workbook.getActiveWorksheet`, для доступа к компонентам книги. Поэтому нужно, чтобы в книге и в таблице были согласованные имена, на которые сможет ссылаться Power Automate.

1. Создайте новую книгу с именем **MyWorkbook**.

2. Перейдите на вкладку **Автоматизация** и выберите **Редактор кода**.

3. Выберите **Новый сценарий**.

4. Замените имеющийся код на следующий и нажмите кнопку **Выполнить**. При том будет создана книга с нужными именами листа, таблицы и сводной таблицы.

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

## <a name="create-an-office-script"></a>Создание сценария Office

Создадим сценарий, записывающий информацию из электронной почты. Предположим, что нужно узнать, в какие дни недели мы получаем больше всего почты, и сколько уникальных отправителей отправляют ее. В нашей книге содержится таблица со столбцами **Дата**, **День недели**, **Адрес электронной почты** и **Тема**. Кроме того, в книге содержится сводная таблица, содержащая **День недели** и **Адрес электронной почты** (это иерархии строк). Количество уникальных **тем** — это отображаемая объединенная информация (иерархия данных). Наш сценарий будет обновлять эту сводную таблицу после обновления таблицы электронной почты.

1. В окне **Редактор кода** выберите **Создать сценарий**.

2. Поток, который мы создадим на более позднем этапе, будет отправлять данные о каждом полученном сообщении электронной почты. Сценарий должен обращаться к этим входным данным с помощью параметров в функции `main`. Замените сценарий по умолчанию следующим сценарием.

    ```TypeScript
    function main(
      workbook: ExcelScript.Workbook,
      from: string,
      dateReceived: string,
      subject: string) {

    }
    ```

3. Этому сценарию требуется доступ к таблице книги и к сводной таблице. Добавьте следующий код в текст сценария после открывающего символа `{`:

    ```TypeScript
    // Get the email table.
    let emailWorksheet = workbook.getWorksheet("Emails");
    let table = emailWorksheet.getTable("EmailTable");
  
    // Get the PivotTable.
    let pivotTableWorksheet = workbook.getWorksheet("Subjects");
    let pivotTable = pivotTableWorksheet.getPivotTable("Pivot");
    ```

4. Параметр `dateReceived` относится к типу `string`. Преобразуем его в объекту [`Date`](../develop/javascript-objects.md#date), чтобы можно было удобно получать день недели. После этого нужно будет сопоставить значение номера дня с более читаемой версией. Добавьте следующий код в конце сценария перед закрывающим символом `}`

    ```TypeScript
      // Parse the received date string to determine the day of the week.
      let emailDate = new Date(dateReceived);
      let dayName = emailDate.toLocaleDateString("en-US", { weekday: 'long' });
    ```

5. Строка `subject` может включать тег ответа "RE:". Давайте удалим этот тег из строки, чтобы у сообщений электронной почте в одной и той же беседе была одинаковая тема для таблицы. Добавьте следующий код в конце сценария перед закрывающим символом `}`

    ```TypeScript
    // Remove the reply tag from the email subject to group emails on the same thread.
    let subjectText = subject.replace("Re: ", "");
    subjectText = subjectText.replace("RE: ", "");
    ```

6. Теперь, когда данные электронной почты отформатированы по нашему желанию, добавим строку в таблицу электронной почты. Добавьте следующий код в конце сценария перед закрывающим символом `}`

    ```TypeScript
    // Add the parsed text to the table.
    table.addRow(-1, [dateReceived, dayName, from, subjectText]);
    ```

7. Теперь нужно обновить сводную таблицу. Добавьте следующий код в конце сценария перед закрывающим символом `}`

    ```TypeScript
    // Refresh the PivotTable to include the new row.
    pivotTable.refresh();
    ```

8. Переименуйте сценарий в **Запись электронной почты** и нажмите кнопку **Сохранить сценарий**.

Теперь сценарий готов для рабочего процесса Power Automate. Сценарий должен выглядеть примерно так:

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

## <a name="create-an-automated-workflow-with-power-automate"></a>Создание автоматизированного рабочего процесса с помощью Power Automate

1. Войдите на [сайт Power Automate](https://flow.microsoft.com).

2. В меню в левой части экрана выберите **Создать**. При этом откроется список способов создания новых рабочих процессов.

    ![Кнопка "Создать" в Power Automate.](../images/power-automate-tutorial-1.png)

3. В разделе **Начать с пустого** выберите **Автоматизированный рабочий процесс**. В этом случае создается рабочий процесс, запускаемый каким-либо событием, например получением сообщения электронной почты.

    ![Функция "Автоматизированный рабочий процесс" в Power Automate.](../images/power-automate-params-tutorial-1.png)

4. В появившемся диалоговом окне введите имя рабочего процесса в текстовом поле **Имя рабочего процесса**. Затем выберите **При получении новой электронной почты** в списке параметров **Выберите триггер рабочего процесса**. Может потребоваться найти этот параметр с помощью поля поиска. Затем нажмите кнопку **Создать**.

    ![Часть окна "Создание автоматизированного рабочего процесса" в Power Automate с параметром "получение нового сообщения электронной почты".](../images/power-automate-params-tutorial-2.png)

    > [!NOTE]
    > В этом учебном руководстве используется Outlook. Можно использовать любую предпочитаемую вами службу электронной почты, хотя в этом случае некоторые параметры могут отличаться.

5. Нажмите кнопку **Новый шаг**.

6. Перейдите на вкладку **Стандартные** и выберите **Excel Online (бизнес)**.

    ![Функция Power Automate для Excel Online (бизнес).](../images/power-automate-tutorial-4.png)

7. В разделе **Действия** выберите **Запустить сценарий (предварительная версия)**.

    ![Вариант действия Power Automate "Запуск сценария" (предварительная версия).](../images/power-automate-tutorial-5.png)

8. Затем выберите книгу, сценарий и исходные аргументы сценария для использования на следующем шаге. В этом учебнике вы будете использовать книгу, созданную в OneDrive, но вы можете воспользоваться любой книгой в OneDrive или на сайте SharePoint. Укажите следующие параметры для соединителя **Запуск сценария**.

    - **Расположение**: OneDrive для бизнеса
    - **Библиотека документов**: OneDrive
    - **Файл**: MyWorkbook.xlsx *(выбран с помощью браузера файлов)*
    - **Сценарий**: запись электронной почты
    - **от**: От *(динамическое содержимое из Outlook)*
    - **dateReceived**: Время получения *(динамическое содержимое из Outlook)*
    - **тема**: Тема *(динамическое содержимое из Outlook)*

    *Обратите внимание, что эти параметры сценария будут отображаться только после выбора сценария.*

    ![Параметры варианта действия Power Automate "Запуск сценария" (предварительная версия).](../images/power-automate-params-tutorial-3.png)

9. Нажмите кнопку **Сохранить**.

Теперь рабочий процесс включен. Он будет автоматически выполнять сценарий каждый раз при получении сообщения электронной почты через Outlook.

## <a name="manage-the-script-in-power-automate"></a>Управление сценарием в Power Automate

1. На главной странице Power Automate выберите **Мои рабочие процессы**.

    ![Кнопка "Мои рабочие процессы" в Power Automate.](../images/power-automate-tutorial-7.png)

2. Выберите рабочий процесс. Здесь можно просмотреть историю запусков. Можно обновить страницу или нажать кнопку обновления **всех запусков**, чтобы обновить историю. Рабочий процесс запустится вскоре после получения сообщения электронной почты. Проверьте рабочий процесс, отправив себе сообщение электронной почты.

При срабатывании рабочего процесса и успешном выполнении сценария должна обновляться таблица книги и сводная таблица.

![Таблица электронной почты после нескольких выполнений рабочего процесса.](../images/power-automate-params-tutorial-4.png)

![Сводная таблица после нескольких выполнений рабочего процесса.](../images/power-automate-params-tutorial-5.png)

## <a name="next-steps"></a>Дальнейшие действия

Посетите страницу [Запуск сценариев Office с помощью Power Automate](../develop/power-automate-integration.md) для получения дополнительных сведений о подключениях сценариев Office с помощью Power Automate.

Кроме того, прочтите статью [Образец сценария автоматизированных напоминаний о задачах](../resources/scenarios/task-reminders.md), чтобы узнать, как использовать сценарии Office и Power Automate с адаптивными карточками Teams.
