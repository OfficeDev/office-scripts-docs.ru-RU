---
title: Запуск скриптов Office с помощью Power Automate
description: Как получить сценарии Office для Excel в Интернете с рабочим процессом Power Automate.
ms.date: 06/22/2022
ms.localizationpriority: medium
ms.openlocfilehash: 61e51861bd2c987c25d40e9ac6d2247122256918
ms.sourcegitcommit: c5ffe0a95b962936ee92e7ffe17388bef6d4fad8
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 06/29/2022
ms.locfileid: "66241858"
---
# <a name="run-office-scripts-with-power-automate"></a>Запуск скриптов Office с помощью Power Automate

[Power Automate](https://flow.microsoft.com) позволяет добавлять скрипты Office в более крупный автоматизированный рабочий процесс. Power Automate может выполнять такие действия, как добавление содержимого сообщения электронной почты в таблицу листа или создание действий в средствах управления проектами на основе комментариев к книге.

## <a name="get-started"></a>Начало работы

Если вы еще не используете Power Automate, рекомендуем посетить сайт ["Начало работы с Power Automate"](/power-automate/getting-started). Здесь вы можете узнать больше обо всех доступных возможностях автоматизации. В этом документе основное внимание уделяется работе сценариев Office с Power Automate и их улучшению в Excel.

### <a name="step-by-step-tutorials"></a>Пошаговые руководства

Существует три пошаговые руководства по Power Automate и сценариям Office. В них показано, как объединить службы автоматизации и передать данные между книгой и потоком.

- [Вызов сценариев из активированного вручную потока Power Automate](../tutorials/excel-power-automate-manual.md)
- [Передача данных сценариям в автоматически запускаемых рабочих процессах Power Automate](../tutorials/excel-power-automate-trigger.md)
- [Возвращение данных из сценария в автоматически запускаемый поток Power Automate](../tutorials//excel-power-automate-returns.md)

### <a name="create-a-flow-from-excel"></a>Создание потока из Excel

Вы можете приступить к работе с Power Automate в Excel с помощью различных шаблонов потоков. На **вкладке "Автоматизация** " выберите **"Автоматизация задачи"**.

:::image type="content" source="../images/automate-a-task-button.png" alt-text="Кнопка &quot;Автоматизация задачи&quot; на ленте.":::

Откроется область задач с несколькими вариантами подключения сценариев Office к более крупным автоматизированным решениям. Выберите любой вариант для начала. Поток поставляется с текущей книгой.

:::image type="content" source="../images/automate-a-task-choices.png" alt-text="Область задач с параметрами шаблона потока, такими как &quot;Запланировать выполнение скрипта Office в Excel, а затем отправить сообщение электронной почты&quot; и &quot;Выполнение скрипта Office в Excel при получении Microsoft Forms ответа&quot;.":::

> [!TIP]
> Вы также можете приступить к выполнению потока из меню " **Дополнительные параметры" (...)** в отдельном скрипте.

## <a name="excel-online-business-connector"></a>Соединитель Excel Online (бизнес)

[Соединители](/connectors/connectors) — это мосты между Power Automate и приложениями. [Соединитель Excel Online (business)](/connectors/excelonlinebusiness) предоставляет потокам доступ к книгам Excel. Действие "Запуск скрипта" позволяет вызывать любой сценарий Office, доступный через выбранную книгу. Вы также можете предоставить входные параметры скриптов, чтобы потоком можно было предоставить данные, или получить сведения о возврате скрипта для последующих шагов в потоке.

> [!IMPORTANT]
> Действие "Запуск скрипта" предоставляет пользователям, использующим соединитель Excel, значительный доступ к книге и ее данным. Кроме того, существуют риски безопасности со скриптами, которые выполняют внешние вызовы API, как описано во внешних вызовах [Power Automate](external-calls.md). Если администратора беспокоит раскрытие конфиденциальных данных, он может либо отключить соединитель Excel Online, либо ограничить доступ к скриптам Office с помощью элементов управления администратора сценариев [Office](/microsoft-365/admin/manage/manage-office-scripts-settings).

> [!IMPORTANT]
> Сейчас Power Automate **не поддерживает** сценарии, хранящиеся в SharePoint.

## <a name="data-transfer-in-flows-for-scripts"></a>Передача данных в потоках для сценариев

Power Automate позволяет передавать фрагменты данных между шагами потока. Скрипты можно настроить так, чтобы они принимали любые необходимые типы информации и возвращали все данные из книги, которые вам нужны в потоке. Входные данные для скрипта задается путем добавления параметров `main` в функцию (в дополнение к).`workbook: ExcelScript.Workbook` Выходные данные скрипта объявляются путем добавления типа возвращаемого значения в `main`.

> [!NOTE]
> При создании блока Run Script в потоке заполняются принятые параметры и возвращаемые типы. Если вы измените параметры или типы возвращаемых значений скрипта, вам потребуется повторить блок "Выполнить сценарий" потока. Это гарантирует, что данные анализируются правильно.

В следующих разделах рассматриваются сведения о входных и выходных данных для сценариев, используемых в Power Automate. Если вы хотите использовать практический подход к изучению этого раздела, попробуйте передать данные в скрипты в руководстве по автоматическому запуску [потока Power Automate](../tutorials/excel-power-automate-trigger.md) или изучите пример сценария [](../resources/scenarios/task-reminders.md) автоматического напоминания о задачах.

### <a name="main-parameters-pass-data-to-a-script"></a>`main` Параметры: передача данных в скрипт

Все входные данные скрипта указаны в качестве дополнительных параметров для функции `main` . Например, если вы хотите `string` , чтобы скрипт принял имя, представляющее имя в качестве входных данных, необходимо `main` `function main(workbook: ExcelScript.Workbook, name: string)`изменить подпись на .

При настройке потока в Power Automate можно указать входные данные скрипта в виде статических значений, [выражений](/power-automate/use-expressions-in-conditions) или динамического содержимого. Подробные сведения о соединителе отдельной службы см. в документации [по Соединителю Power Automate](/connectors/).

#### <a name="type-restrictions"></a>Ограничения типов

При добавлении входных параметров в функцию скрипта `main` учитывайте следующие квоты и ограничения. Они также применяются к типу возвращаемого значения скрипта.

1. Первый параметр должен иметь тип `ExcelScript.Workbook`. Имя параметра не имеет значения.

1. `string`Типы , `number`, `boolean`, `unknown`, и `object``undefined` поддерживаются.

1. Поддерживаются массивы (и `[]` `Array<T>` стили) перечисленных выше типов. Также поддерживаются вложенные массивы.

1. Типы объединения допускаются, если они являются объединением литералов, принадлежащих одному типу (например`"Left" | "Right"`, не ).`"Left", 5` Также поддерживаются объединения поддерживаемого типа с неопределенным типом (например, `string | undefined`).

1. Типы объектов допускаются, если они содержат `string`свойства типа, `number`поддерживаемые `boolean`массивы или другие поддерживаемые объекты. В следующем примере показаны вложенные объекты, которые поддерживаются в качестве типов параметров.

    ```TypeScript
    // The Employee object is supported because Position is also composed of supported types.
    interface Employee {
        name: string;
        job: Position;
    }

    interface Position {
        id: number;
        title: string;
    }
    ```

1. Объекты должны иметь определение интерфейса или класса, определенное в скрипте. Объект также можно определить анонимно, как показано в следующем примере.

    ```TypeScript
    function main(workbook: ExcelScript.Workbook): {name: string, email: string}
    ```

#### <a name="optional-and-default-parameters"></a>Необязательные параметры и параметры по умолчанию

1. Необязательные параметры разрешены и обозначаются необязательным модификатором `?` (например, `function main(workbook: ExcelScript.Workbook, Name?: string)`).

1. Допустимы значения параметров по умолчанию (например,`function main(workbook: ExcelScript.Workbook, Name: string = 'Jane Doe')`

### <a name="return-data-from-a-script"></a>Возврат данных из скрипта

Скрипты могут возвращать данные из книги, которые будут использоваться в качестве динамического содержимого в потоке Power Automate. Те [же ограничения типа, которые были указаны ранее](#type-restrictions) , применяются к типу возвращаемого значения. Чтобы вернуть объект, добавьте в функцию синтаксис типа возвращаемого `main` значения. Например, если вы хотите вернуть значение `string` из скрипта, ваша подпись `main` будет выглядеть следующим образом `function main(workbook: ExcelScript.Workbook): string`.

## <a name="example"></a>Пример

На следующем снимке экрана показан поток Power Automate, который активируется при каждой назначенной проблеме [GitHub](https://github.com/) . Поток запускает сценарий, который добавляет проблему в таблицу в книге Excel. Если в этой таблице есть пять или более проблем, поток отправляет напоминание по электронной почте.

:::image type="content" source="../images/power-automate-parameter-return-sample.png" alt-text="Редактор потока Power Automate, показывающий пример потока.":::

Функция `main` скрипта указывает идентификатор проблемы и заголовок выдачи в качестве входных параметров, а скрипт возвращает количество строк в таблице проблем.

```TypeScript
function main(
  workbook: ExcelScript.Workbook,
  issueId: string,
  issueTitle: string): number {
  // Get the "GitHub" worksheet.
  let worksheet = workbook.getWorksheet("GitHub");

  // Get the first table in this worksheet, which contains the table of GitHub issues.
  let issueTable = worksheet.getTables()[0];

  // Add the issue ID and issue title as a row.
  issueTable.addRow(-1, [issueId, issueTitle]);

  // Return the number of rows in the table, which represents how many issues are assigned to this user.
  return issueTable.getRangeBetweenHeaderAndTotal().getRowCount();
}
```

## <a name="see-also"></a>См. также

- [Вызов сценариев из активированного вручную потока Power Automate](../tutorials/excel-power-automate-manual.md)
- [Передача данных сценариям в автоматически запускаемых рабочих процессах Power Automate](../tutorials/excel-power-automate-trigger.md)
- [Возвращение данных из сценария в автоматически запускаемый поток Power Automate](../tutorials/excel-power-automate-returns.md)
- [Сведения об устранении неполадок в Power Automate с помощью сценариев Office](../testing/power-automate-troubleshooting.md)
- [Начало работы с Power Automate](/power-automate/getting-started)
- [Справочная документация по соединителю Excel Online (business)](/connectors/excelonlinebusiness/)
