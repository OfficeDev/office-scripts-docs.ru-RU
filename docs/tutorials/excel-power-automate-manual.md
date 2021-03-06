---
title: Вызов сценариев из активированного вручную потока Power Automate
description: В этом руководстве рассказывается об использовании сценариев Office в Power Automate с помощью триггера с ручным срабатыванием.
ms.date: 06/29/2021
localization_priority: Priority
ms.openlocfilehash: 1a8b9659ec6f6354d583496ba0f3e94d4a13c01b
ms.sourcegitcommit: 211c157ca746e266eeb079f5fa1925a1e35ab702
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/07/2021
ms.locfileid: "53313997"
---
# <a name="call-scripts-from-a-manual-power-automate-flow"></a>Вызов сценариев из активированного вручную потока Power Automate

В этом руководстве объясняется, как запускать сценарий Office для Excel в Интернете с помощью [Power Automate](https://flow.microsoft.com). Вы создадите сценарий, обновляющий значения двух ячеек текущим временем. После этого вы подключите этот сценарий к запускаемому вручную потоку Power Automate, чтобы сценарий выполнялся при каждом нажатии кнопки Power Automate. После знакомства с базовым шаблоном вы можете расширить поток, чтобы включить другие приложения и автоматизировать дополнительные повседневные рабочие процессы.

> [!TIP]
> Если вы только приступили к работе со сценариями Office, рекомендуем начать с учебника [Запись, редактирование и создание сценариев Office в Excel в Интернете](excel-tutorial.md). [Сценарии Office используют TypeScript](../overview/code-editor-environment.md), и этот учебник предназначен для пользователей с начальным и средним уровнем знаний по JavaScript или TypeScript. Если вы впервые работаете с JavaScript, рекомендуем начать с [учебника Mozilla по JavaScript](https://developer.mozilla.org/docs/Web/JavaScript/Guide/Introduction).

## <a name="prerequisites"></a>Предварительные условия

[!INCLUDE [Tutorial prerequisites](../includes/power-automate-tutorial-prerequisites.md)]

## <a name="prepare-the-workbook"></a>Подготовка книги

Power Automate не должен использовать [относительные ссылки](../testing/power-automate-troubleshooting.md#avoid-relative-references), такие как `Workbook.getActiveWorksheet`, для доступа к компонентам книги. Поэтому нужно использовать книгу и лист с именами, на которые может ссылаться Power Automate.

1. Создайте новую книгу под названием **MyWorkbook**.

2. В книге **MyWorkbook** создайте лист под названием **TutorialWorksheet**.

## <a name="create-an-office-script"></a>Создание сценария Office

1. Перейдите на вкладку **Автоматизация** и выберите **Все сценарии**.

2. Выберите **Новый сценарий**.

3. Замените сценарий по умолчанию следующим сценарием. Этот сценарий добавляет текущую дату и время в первые две ячейки листа **TutorialWorksheet**.

    ```TypeScript
    function main(workbook: ExcelScript.Workbook) {
      // Get the "TutorialWorksheet" worksheet from the workbook.
      let worksheet = workbook.getWorksheet("TutorialWorksheet");

      // Get the cells at A1 and B1.
      let dateRange = worksheet.getRange("A1");
      let timeRange = worksheet.getRange("B1");

      // Get the current date and time using the JavaScript Date object.
      let date = new Date(Date.now());

      // Add the date string to A1.
      dateRange.setValue(date.toLocaleDateString());

      // Add the time string to B1.
      timeRange.setValue(date.toLocaleTimeString());
    }
    ```

4. Переименуйте сценарий в **Установка даты и времени**. Выберите имя сценария, чтобы изменить его.

5. Сохраните сценарий, нажав **Сохранить сценарий**.

## <a name="create-an-automated-workflow-with-power-automate"></a>Создание автоматизированного рабочего процесса с помощью Power Automate

1. Войдите на [сайт Power Automate](https://flow.microsoft.com).

2. В меню в левой части экрана выберите **Создать**. При этом откроется список способов создания новых рабочих процессов.

    :::image type="content" source="../images/power-automate-tutorial-1.png" alt-text="Кнопка &quot;Создать&quot; в Power Automate.":::

3. В разделе **Создание нового** выберите пункт **Мгновенный поток**. В результате будет создан активированный вручную рабочий процесс.

    :::image type="content" source="../images/power-automate-tutorial-2.png" alt-text="Параметр &quot;Мгновенный поток&quot; в Power Automate для создания потока.":::

4. В открывшемся диалоговом окне введите имя для своего потока в поле **Имя потока**, выберите **Активировать поток вручную** из списка вариантов в разделе **Выберите способ запуска потока** и нажмите **Создать**.

    :::image type="content" source="../images/power-automate-tutorial-3.png" alt-text="Параметр &quot;Активировать поток вручную&quot; в Power Automate.":::

    Обратите внимание: запускаемый вручную поток — это лишь один из многих типов потоков. В следующем руководстве описывается создание потока, который будет выполняться автоматически при получении вами сообщения электронной почты.

5. Выберите **Новый шаг**.

6. Перейдите на вкладку **Стандартные** и выберите **Excel Online (бизнес)**.

    :::image type="content" source="../images/power-automate-tutorial-4.png" alt-text="Параметр Excel Online (бизнес) в Power Automate.":::

7. В разделе **Действия** выберите **Запуск скрипта**.

    :::image type="content" source="../images/power-automate-tutorial-5.png" alt-text="Вариант действия &quot;Запуск скрипта&quot; в Power Automate.":::

8. Затем выберите книгу и сценарий для использования на следующем шаге. В этом учебнике вы будете использовать книгу, созданную в OneDrive, но вы можете воспользоваться любой книгой в OneDrive или на сайте SharePoint. Укажите следующие параметры для соединителя **Запуск сценария**.

    - **Расположение**: OneDrive для бизнеса
    - **Библиотека документов**: OneDrive
    - **Файл**: MyWorkbook.xlsx *(выбран с помощью браузера файлов)*
    - **Сценарий**: установка даты и времени

    :::image type="content" source="../images/power-automate-tutorial-6.png" alt-text="Параметры соединителя Power Automate для запуска сценария.":::

9. Нажмите **Сохранить**.

Теперь ваш поток готов к запуску с помощью Power Automate. Вы можете проверить его с помощью кнопки **Тест** в редакторе потока или выполнить остальные действия согласно руководству, чтобы запустить поток из вашей коллекции потоков.

## <a name="run-the-script-through-power-automate"></a>Запуск сценария с помощью Power Automate

1. На главной странице Power Automate выберите **Мои потоки**.

    :::image type="content" source="../images/power-automate-tutorial-7.png" alt-text="Кнопка &quot;Мои потоки&quot; в Power Automate.":::

2. Выберите **Мой учебный поток** из списка во вкладке **Мои потоки**. При этом будут показаны подробные сведения о потоке, который мы создали ранее.

3. Нажмите **Запустить**.

    :::image type="content" source="../images/power-automate-tutorial-8.png" alt-text="Кнопка &quot;Запустить&quot; в Power Automate.":::

4. Появится панель задач для запуска потока. Когда будет предложено выполнить **Вход** в Excel Online, нажмите **Продолжить**.

5. Выберите **Запустить поток**. При этом запустится поток, выполняющий связанный сценарий Office.

6. Нажмите **Готово**. Вы можете заметить, что раздел **Запуски** соответствующим образом обновлен.

7. Обновите страницу, чтобы увидеть результаты работы Power Automate. После этого перейдите в книгу, где должны отобразиться обновленные ячейки. В случае неудачи проверьте параметры этого потока и запустите его еще раз.

    :::image type="content" source="../images/power-automate-tutorial-9.png" alt-text="В результатах работы Power Automate показано успешное выполнение потока.":::

## <a name="next-steps"></a>Дальнейшие действия

Прочитайте раздел руководства [Передача данных сценариям в автоматически запускаемом потоке Power Automate](excel-power-automate-trigger.md). В нем рассказывается о том, как передать данные из службы рабочего процесса в ваш сценарий Office и запустить поток Power Automate при возникновении определенных событий.
