---
title: Выполнение Office с помощью Power Automate
description: Как получить Office скрипты для Excel в Интернете с рабочим Power Automate рабочим процессом.
ms.date: 03/08/2022
ms.localizationpriority: medium
ms.openlocfilehash: 67e48d297a8db16661ce394a11f2e425bc0a33be
ms.sourcegitcommit: 34c7740c9bff0e4c7426e01029f967724bfee566
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 05/13/2022
ms.locfileid: "65393679"
---
# <a name="run-office-scripts-with-power-automate"></a>Выполнение Office с помощью Power Automate

[Power Automate](https://flow.microsoft.com) позволяет добавлять Office в более крупный автоматизированный рабочий процесс. Вы можете Power Automate такие действия, как добавление содержимого сообщения электронной почты в таблицу листа или создание действий в средствах управления проектами на основе комментариев к книге.

## <a name="get-started"></a>Начало работы

Если вы еще не Power Automate, рекомендуем посетить начало работы [с Power Automate](/power-automate/getting-started). Здесь вы можете узнать больше обо всех доступных возможностях автоматизации. В этом документе основное внимание уделяется работе Office сценариев с Power Automate и как это может помочь улучшить Excel взаимодействия.

Чтобы начать объединение Power Automate и Office сценариев, следуйте указаниям в руководстве по началу работы с [Power Automate.](../tutorials/excel-power-automate-manual.md) Вы научитесь создавать поток, который вызывает простой сценарий. После завершения работы с этим руководством и передачи данных в скрипты в руководстве по автоматическому Power Automate [flow](../tutorials/excel-power-automate-trigger.md) вернитесь сюда, чтобы получить подробные сведения о подключении Office скриптов к Power Automate потокам.

## <a name="excel-online-business-connector"></a>Excel Online (Business) connector

[Соединители](/connectors/connectors) — это мосты между Power Automate приложениями. [Соединитель Excel Online (Business)](/connectors/excelonlinebusiness) предоставляет потокам доступ к Excel книг. Действие "Запуск скрипта" позволяет вызвать любой Office, доступный через выбранную книгу. Вы также можете предоставить входные параметры скриптов, чтобы потоком можно было предоставить данные, или получить сведения о возврате скрипта для последующих шагов в потоке.

> [!IMPORTANT]
> Действие "Запуск скрипта" предоставляет пользователям, использующим Excel соединителя, значительный доступ к книге и ее данным. Кроме того, существуют риски безопасности со скриптами, которые выполняют внешние вызовы API, как описано во внешних вызовах [Power Automate](external-calls.md). Если администратора беспокоит раскрытие конфиденциальных данных, он может либо отключить соединитель Excel Online, либо ограничить доступ к Office Scripts с помощью элементов управления [Office Scripts](/microsoft-365/admin/manage/manage-office-scripts-settings).

## <a name="data-transfer-in-flows-for-scripts"></a>Передача данных в потоках для сценариев

Power Automate позволяет передавать фрагменты данных между шагами потока. Скрипты можно настроить так, чтобы они принимали любые необходимые типы информации и возвращали все данные из книги, которые вам нужны в потоке. Входные данные для скрипта задается путем добавления параметров `main` в функцию (в дополнение к).`workbook: ExcelScript.Workbook` Выходные данные скрипта объявляются путем добавления типа возвращаемого значения в `main`.

> [!NOTE]
> При создании блока Run Script в потоке заполняются принятые параметры и возвращаемые типы. Если вы измените параметры или типы возвращаемых значений скрипта, вам потребуется повторить блок "Выполнить сценарий" потока. Это гарантирует, что данные анализируются правильно.

В следующих разделах рассматриваются сведения о входных и выходных данных для скриптов, используемых в Power Automate. Если вы хотите использовать практический подход к изучению этого раздела, попробуйте передать данные в скрипты в руководстве по автоматическому запуску Power Automate [потока](../tutorials/excel-power-automate-trigger.md) или изучите пример сценария автоматического напоминания о задачах.[](../resources/scenarios/task-reminders.md)

### <a name="main-parameters-pass-data-to-a-script"></a>`main` Параметры: передача данных в скрипт

Все входные данные скрипта указаны в качестве дополнительных параметров для функции `main` . Например, если вы хотите `string` , чтобы скрипт принял имя, представляющее имя в качестве входных данных, необходимо `main` `function main(workbook: ExcelScript.Workbook, name: string)`изменить подпись на .

При настройке потока в Power Automate можно указать входные данные скрипта в виде статических значений, [выражений](/power-automate/use-expressions-in-conditions) или динамического содержимого. Подробные сведения о соединителе отдельной службы см. в документации [Power Automate connector](/connectors/).

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

Скрипты могут возвращать данные из книги, которые будут использоваться в качестве динамического содержимого в Power Automate потоке. Те [же ограничения типа, которые были указаны ранее](#type-restrictions) , применяются к типу возвращаемого значения. Чтобы вернуть объект, добавьте в функцию синтаксис типа возвращаемого `main` значения. Например, если вы хотите вернуть значение `string` из скрипта, ваша подпись `main` будет выглядеть следующим образом `function main(workbook: ExcelScript.Workbook): string`.

## <a name="example"></a>Пример

На следующем снимке экрана показан Power Automate, который активируется при каждом GitHub проблемы.[](https://github.com/) Поток запускает сценарий, который добавляет проблему в таблицу в Excel книге. Если в этой таблице есть пять или более проблем, поток отправляет напоминание по электронной почте.

:::image type="content" source="../images/power-automate-parameter-return-sample.png" alt-text="Редактор Power Automate потока с примером потока.":::

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
- [Сведения об устранении неполадок Power Automate с помощью Office сценариев](../testing/power-automate-troubleshooting.md)
- [Начало работы с Power Automate](/power-automate/getting-started)
- [Excel справки по соединителю Online (Business)](/connectors/excelonlinebusiness/)
