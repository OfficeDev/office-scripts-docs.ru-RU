---
title: Передача данных сценариям в автоматически запускаемых рабочих процессах Power Automate
description: Учебное руководство, посвященное запуску сценариев Office для Excel в Интернете с помощью Power Automate при получении электронной почты с дальнейшей передачей данных рабочего процесса в сценарий.
ms.date: 06/10/2022
ms.localizationpriority: high
ms.openlocfilehash: 73a551df09eadba1f6e75de35e17e1c5a93498e9
ms.sourcegitcommit: dd01979d34b3499360d2f79a56f8a8f24f480eed
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 06/15/2022
ms.locfileid: "66088141"
---
# <a name="pass-data-to-scripts-in-an-automatically-run-power-automate-flow"></a>Передача данных сценариям в автоматически запускаемых рабочих процессах Power Automate

В этом руководстве объясняется, как использовать сценарий Office для Excel в Интернете с помощью автоматизированных рабочих процессов [Power Automate](https://flow.microsoft.com). Сценарий будет автоматически выполняться каждый раз при получении электронной почты. Данные из сообщений электронной почты будут записываться в книгу Excel. Возможность передавать данные из других приложений в сценарии Office предоставляет вам значительную гибкость и свободу в автоматизированных процессах.

> [!TIP]
> Если вы только приступили к работе со сценариями Office, рекомендуем начать с учебника [Запись, редактирование и создание сценариев Office в Excel в Интернете](excel-tutorial.md). Если вы впервые используете Power Automate, рекомендуем начать с учебника [Вызов сценариев из активированного вручную потока Power Automate](excel-power-automate-manual.md). [Сценарии Office используют TypeScript](../overview/code-editor-environment.md), и этот учебник предназначен для пользователей с начальным и средним уровнем знаний по JavaScript или TypeScript. Если вы впервые работаете с JavaScript, рекомендуем начать с [учебника Mozilla по JavaScript](https://developer.mozilla.org/docs/Web/JavaScript/Guide/Introduction).

## <a name="prerequisites"></a>Предварительные условия

[!INCLUDE [Tutorial prerequisites](../includes/power-automate-tutorial-prerequisites.md)]

## <a name="prepare-the-workbook"></a>Подготовка книги

Power Automate не должен использовать [относительные ссылки](../testing/power-automate-troubleshooting.md#avoid-relative-references), такие как `Workbook.getActiveWorksheet`, для доступа к компонентам книги. Поэтому нужно, чтобы в книге и в таблице были согласованные имена, на которые сможет ссылаться Power Automate.

1. Создайте новую книгу с именем **MyWorkbook**.

2. Перейдите на вкладку **Автоматизация** и выберите **Все сценарии**.

3. Выберите **Создать сценарий**.

4. Замените имеющийся код на следующий и нажмите **Запустить**. При том будет создана книга с нужными именами листа, таблицы и сводной таблицы.

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

1. В области задач "Редактор кода" выберите **Создать сценарий**.

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

8. Переименуйте сценарий в **Запись электронной почты** и нажмите **Сохранить сценарий**.

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

    :::image type="content" source="../images/power-automate-tutorial-1.png" alt-text="Кнопка &quot;Создать&quot; в Power Automate.":::

3. В разделе **Начать с пустого** выберите **Автоматизированный рабочий процесс**. В этом случае создается рабочий процесс, запускаемый каким-либо событием, например получением сообщения электронной почты.

    :::image type="content" source="../images/power-automate-params-tutorial-1.png" alt-text="Параметр &quot;Автоматизированный поток&quot; в Power Automate.":::

4. В появившемся диалоговом окне введите имя рабочего процесса в текстовом поле **Имя рабочего процесса**. Затем выберите **При получении новой электронной почты** в списке параметров **Выберите триггер рабочего процесса**. Может потребоваться найти этот параметр с помощью поля поиска. По завершении нажмите **Создать**.

    :::image type="content" source="../images/power-automate-params-tutorial-2.png" alt-text="Часть потока Power Automate с указанием параметров &quot;Имя потока&quot; и &quot;Выберите триггер потока&quot;. Имя потока — &quot;Поток записи электронной почты&quot;, а триггер — &quot;При поступления нового сообщения в Outlook&quot;.":::

    > [!NOTE]
    > В этом руководстве используется Outlook. Можно использовать любую предпочитаемую вами службу электронной почты, хотя в этом случае некоторые параметры могут отличаться.

5. Выберите **Новый шаг**.

6. Перейдите на вкладку **Стандартные** и выберите **Excel Online (бизнес)**.

    :::image type="content" source="../images/power-automate-tutorial-4.png" alt-text="Параметр Excel Online (бизнес) в Power Automate.":::

7. В разделе **Действия** выберите **Запуск скрипта**.

    :::image type="content" source="../images/power-automate-tutorial-5.png" alt-text="Вариант действия &quot;Запуск скрипта&quot; в Power Automate.":::

8. Затем выберите книгу, сценарий и исходные аргументы сценария для использования на следующем шаге. В этом учебнике вы будете использовать книгу, созданную в OneDrive, но вы можете воспользоваться любой книгой в OneDrive или на сайте SharePoint. Укажите следующие параметры для соединителя **Запуск сценария**.

    - **Расположение**: OneDrive для бизнеса
    - **Библиотека документов**: OneDrive
    - **Файл**: MyWorkbook.xlsx *(выбран с помощью браузера файлов)*
    - **Сценарий**: запись электронной почты
    - **от**: От *(динамическое содержимое из Outlook)*
    - **dateReceived**: Время получения *(динамическое содержимое из Outlook)*
    - **тема**: Тема *(динамическое содержимое из Outlook)*

    *Обратите внимание, что эти параметры сценария будут отображаться только после выбора сценария.*

    :::image type="content" source="../images/power-automate-params-tutorial-3.png" alt-text="Действие запуска сценария Power Automate с параметрами, которые отображаются после выбора сценария.":::

9. Нажмите **Сохранить**.

Теперь рабочий процесс включен. Он будет автоматически выполнять сценарий каждый раз при получении сообщения электронной почты через Outlook.

## <a name="manage-the-script-in-power-automate"></a>Управление сценарием в Power Automate

1. На главной странице Power Automate выберите **Мои рабочие процессы**.

    :::image type="content" source="../images/power-automate-tutorial-7.png" alt-text="Кнопка &quot;Мои потоки&quot; в Power Automate.":::

2. Выберите рабочий процесс. Здесь можно просмотреть журнал запусков. Можно обновить страницу или нажать кнопку обновления **Все запуски**, чтобы обновить журнал. Рабочий процесс запустится вскоре после получения сообщения электронной почты. Проверьте рабочий процесс, отправив себе сообщение электронной почты.

При срабатывании рабочего процесса и успешном выполнении сценария должна обновляться таблица книги и сводная таблица.

:::image type="content" source="../images/power-automate-params-tutorial-4.png" alt-text="Лист с таблицей электронной почты после трех запусков потока.":::

:::image type="content" source="../images/power-automate-params-tutorial-5.png" alt-text="Лист со сводной таблицей после трех запусков потока.":::

## <a name="troubleshooting"></a>Устранение неполадок

Одновременное получение нескольких сообщений электронной почты может привести к конфликтам слияния в Excel. Этот риск устраняется путем настройки соединителя электронной почты для выполнения действий только с одним сообщением электронной почты за раз. Выполните следующие действия:

1. Нажмите копку **Меню (...)** в соединителе электронной почты, а затем выберите пункт **Параметры**.

    :::image type="content" source="../images/outlook-connector-settings-1.png" alt-text="Вариант параметра, выделенный в меню соединителя.":::

1. Во всплывающих вариантах выбора **Параметры** переведите элемент управления **Параллелизм** в положение **Включено**. Затем для параметра **Степень параллелизма** установите значение **1**.

    :::image type="content" source="../images/outlook-connector-settings-2.png" alt-text="Параметры параллелизма в меню параметров.":::

## <a name="next-steps"></a>Дальнейшие действия

Прочитайте руководство [Возвращение данных из сценария в автоматически запускаемый поток Power Automate](excel-power-automate-returns.md). Из него вы узнаете, как вернуть данные из сценария в поток.

Кроме того, прочтите статью [Образец сценария автоматизированных напоминаний о задачах](../resources/scenarios/task-reminders.md), чтобы узнать, как использовать сценарии Office и Power Automate с адаптивными карточками Teams.
