---
title: Возвращение данных из сценария в автоматически запускаемый поток Power Automate
description: Руководство по отправке напоминаний по электронной почте путем запуска сценариев Office для Excel в Интернете с помощью Power Automate.
ms.date: 12/15/2020
localization_priority: Priority
ms.openlocfilehash: 1925a95938837707eacddff6832180b12cd2011c
ms.sourcegitcommit: 5f79e5ba9935edb8a890012f2cde3b89fe80faa0
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 12/22/2020
ms.locfileid: "49727086"
---
# <a name="return-data-from-a-script-to-an-automatically-run-power-automate-flow-preview"></a>Возвращение данных из сценария в автоматически запускаемый поток Power Automate (предварительная версия)

В этом руководстве объясняется, как возвращать сведения из сценария Office для Excel в Интернете в рамках автоматизированного рабочего процесса [Power Automate](https://flow.microsoft.com). Создайте сценарий, который выполняется по расписанию и работает с потоком для отправки напоминаний по электронной почте. Этот поток будет запускаться по расписанию и предоставлять напоминания от вашего имени.

> [!TIP]
> Если вы только приступили к работе со сценариями Office, рекомендуем начать с учебника [Запись, редактирование и создание сценариев Office в Excel в Интернете](excel-tutorial.md).
>
> Если вы впервые используете Power Automate, рекомендуем начать с учебников [Вызов сценариев из активированного вручную потока Power Automate](excel-power-automate-manual.md) и [Передача данных в сценарии в автоматически запускаемом потоке Power Automate](excel-power-automate-trigger.md).
>
> [Сценарии Office используют TypeScript](../overview/code-editor-environment.md), и этот учебник предназначен для пользователей с начальным и средним уровнем знаний по JavaScript или TypeScript. Если вы впервые работаете с JavaScript, рекомендуем начать с [учебника Mozilla по JavaScript](https://developer.mozilla.org/docs/Web/JavaScript/Guide/Introduction).

## <a name="prerequisites"></a>Предварительные условия

[!INCLUDE [Tutorial prerequisites](../includes/power-automate-tutorial-prerequisites.md)]

## <a name="prepare-the-workbook"></a>Подготовка книги

1. Загрузите книгу <a href="on-call-rotation.xlsx">on-call-rotation.xlsx</a> в OneDrive.

1. Откройте **on-call-rotation.xlsx** в Excel в Интернете.

1. Добавьте в таблицу строку со своим именем, адресом электронной почты и датами начала и окончания, которые перекрываются с текущей датой.

    > [!IMPORTANT]
    > Сценарий, который вы создаете, использует первую соответствующую запись в таблице, поэтому убедитесь, что ваше имя расположено выше строки с текущей неделей.

    ![Снимок экрана: таблица сменных дежурных на листе Excel](../images/power-automate-return-tutorial-1.png)

## <a name="create-an-office-script"></a>Создание сценария Office

1. Перейдите на вкладку **Автоматизация** и выберите **Все сценарии**.

1. Выберите **Новый сценарий**.

1. Назовите сценарий **Получение дежурного**.

1. Сейчас у вас должен быть пустой сценарий. Нам нужно использовать его для получения адреса электронной почты с листа. Измените функцию `main`, чтобы вернуть строку наподобие этой:

    ```typescript
    function main(workbook: ExcelScript.Workbook) : string {
    }
    ```

1. Затем нам нужно получить все данные из таблицы. Это позволит нам просмотреть каждую строку с помощью сценария. Добавьте следующий код в функцию `main`.

    ```typescript
    // Get the H1 worksheet.
    let worksheet = workbook.getWorksheet("H1");

    // Get the first (and only) table in the worksheet.
    let table = worksheet.getTables()[0];

    // Get the data from the table.
    let tableValues = table.getRangeBetweenHeaderAndTotal().getValues();
    ```

1. Даты в таблице хранятся в виде [порядковых номеров в Excel](https://support.microsoft.com/office/date-systems-in-excel-e7fe7167-48a9-4b96-bb53-5612a800b487). Необходимо преобразовать эти даты в даты JavaScript для сравнения. Добавим вспомогательную функцию в наш сценарий. Добавьте следующий код вне функции `main`:

    ```typescript
    // Convert the Excel date to a JavaScript Date object.
    function convertDate(excelDateValue: number) {
        let javaScriptDate = new Date(Math.round((excelDateValue - 25569) * 86400 * 1000));
        return javaScriptDate;
    }
    ```

1. Теперь необходимо определить, кто на дежурстве сейчас. В строке с его именем дата начала будет предшествовать текущей дате, а дата окончания — следовать за ней. Создавая сценарий, предположим, что сотрудники дежурят по одному. Сценарии могут возвращать массивы для обработки нескольких значений, но в этот раз будет возвращен первый соответствующий адрес электронной почты. Добавьте ниже указанный код в конец функции `main`.

    ```typescript
    // Look for the first row where today's date is between the row's start and end dates.
    let currentDate = new Date();
    for (let row = 0; row < tableValues.length; row++) {
        let startDate = convertDate(tableValues[row][2] as number);
        let endDate = convertDate(tableValues[row][3] as number);
        if (startDate <= currentDate && endDate >= currentDate) {
            // Return the first matching email address.
            return tableValues[row][1].toString();
        }
    }
    ```

1. Окончательный вариант сценария должен выглядеть так:

    ```typescript
    function main(workbook: ExcelScript.Workbook) : string {
        // Get the H1 worksheet.
        let worksheet = workbook.getWorksheet("H1");

        // Get the first (and only) table in the worksheet.
        let table = worksheet.getTables()[0];
    
        // Get the data from the table.
        let tableValues = table.getRangeBetweenHeaderAndTotal().getValues();
    
        // Look for the first row where today's date is between the row's start and end dates.
        let currentDate = new Date();
        for (let row = 0; row < tableValues.length; row++) {
            let startDate = convertDate(tableValues[row][2] as number);
            let endDate = convertDate(tableValues[row][3] as number);
            if (startDate <= currentDate && endDate >= currentDate) {
                // Return the first matching email address.
                return tableValues[row][1].toString();
            }
        }
    }

    // Convert the Excel date to a JavaScript Date object.
    function convertDate(excelDateValue: number) {
        let javaScriptDate = new Date(Math.round((excelDateValue - 25569) * 86400 * 1000));
        return javaScriptDate;
    }
    ```

## <a name="create-an-automated-workflow-with-power-automate"></a>Создание автоматизированного рабочего процесса с помощью Power Automate

1. Войдите на [сайт Power Automate](https://flow.microsoft.com).

1. В меню в левой части экрана выберите **Создать**. При этом откроется список способов создания новых рабочих процессов.

    ![Кнопка "Создать" в Power Automate](../images/power-automate-tutorial-1.png)

1. В разделе **Начать с пустого** выберите **Запланированный облачный поток**.

    ![Кнопка "Запланированный облачный поток" в Power Automate](../images/power-automate-return-tutorial-2.png)

1. Теперь необходимо задать расписание для этого потока. На нашем листе назначение новых дежурных начинается каждый понедельник в первой половине 2021 года. Настроим запуск потока в первую очередь утром в понедельник. Используйте приведенные ниже параметры, чтобы настроить запуск потока каждый понедельник.

    - **Имя потока**: Уведомление дежурного
    - **Начало**: 04.01.2021 в 01:00
    - **Частота повтора**: 1 неделя
    - **В такие дни**: Пн

    ![Окно с указанными параметрами для запланированного потока](../images/power-automate-return-tutorial-3.png)

1. Нажмите кнопку **Создать**.

1. Нажмите кнопку **Новый шаг**.

1. Перейдите на вкладку **Стандартные** и выберите **Excel Online (бизнес)**.

    ![Параметр Excel Online (бизнес) в Power Automate](../images/power-automate-tutorial-4.png)

1. В разделе **Действия** выберите **Запуск сценария (предварительная версия)**.

    ![Вариант действия "Запуск сценария" (предварительная версия) в Power Automate](../images/power-automate-tutorial-5.png)

1. Затем выберите книгу и сценарий для использования на следующем шаге. Используйте книгу **on-call-rotation.xlsx**, созданную в OneDrive. Укажите следующие параметры для соединителя **Запуск сценария**.

    - **Расположение**: OneDrive для бизнеса
    - **Библиотека документов**: OneDrive
    - **Файл**: on-call-rotation.xlsx *(выбран с помощью браузера файлов)*
    - **Сценарий**: Получение дежурного

    ![Параметры соединителя для запуска сценария в Power Automate](../images/power-automate-return-tutorial-4.png)

1. Нажмите кнопку **Новый шаг**.

1. Завершим поток отправкой сообщения с напоминанием. Выберите **Отправить сообщение (V2)** с помощью панели поиска соединителя. Чтобы добавить адрес электронной почты, возвращенный сценарием, используйте элемент управления **Добавить динамическое содержимое**. Он будет помечен как **результат** и значком Excel. Можно использовать любую тему и основной текст.

    ![Параметры соединителя для отправки сообщения в Power Automate](../images/power-automate-return-tutorial-5.png)

    > [!NOTE]
    > В этом учебном руководстве используется Outlook. Можно использовать любую предпочитаемую вами службу электронной почты, хотя в этом случае некоторые параметры могут отличаться.

1. Нажмите кнопку **Сохранить**.

## <a name="test-the-script-in-power-automate"></a>Тестирование сценария в Power Automate

Ваш поток будет запускаться каждый понедельник утром. Вы можете проверить сценарий, нажав кнопку **Проверить** в правом верхнем углу экрана. Выберите **Вручную** и нажмите **Запустить тест**, чтобы запустить поток и проверить поведение. Возможно, вам понадобится предоставить разрешения для Excel и Outlook, чтобы продолжить.

![Кнопка "Проверить" в Power Automate](../images/power-automate-return-tutorial-6.png)

> [!TIP]
> Если поток не сможет отправить сообщение, еще раз проверьте, чтобы на листе был указан действительный адрес электронной почты для текущего диапазона дат в верхней части страницы.

## <a name="next-steps"></a>Дальнейшие действия

Посетите страницу [Запуск сценариев Office с помощью Power Automate](../develop/power-automate-integration.md) для получения дополнительных сведений о подключениях сценариев Office с помощью Power Automate.

Кроме того, прочтите статью [Образец сценария автоматизированных напоминаний о задачах](../resources/scenarios/task-reminders.md), чтобы узнать, как использовать сценарии Office и Power Automate с адаптивными карточками Teams.
