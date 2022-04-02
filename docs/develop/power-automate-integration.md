---
title: Запустите Office скрипты с Power Automate
description: Как получить Office скрипты для Excel в Интернете работы с рабочим Power Automate рабочим процессом.
ms.date: 03/08/2022
ms.localizationpriority: medium
ms.openlocfilehash: dbf65086e564b20ca0fc3a4dc1c527188540be6b
ms.sourcegitcommit: 7023b9e23499806901a5ecf8ebc460b76887cca6
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/31/2022
ms.locfileid: "64585753"
---
# <a name="run-office-scripts-with-power-automate"></a>Запустите Office скрипты с Power Automate

[Power Automate](https://flow.microsoft.com) позволяет добавлять Office скрипты в более крупный автоматизированный рабочий процесс. Можно использовать Power Automate, например добавить содержимое электронной почты в таблицу таблицы или создать действия в средствах управления проектами на основе комментариев к книгам.

## <a name="get-started"></a>Начало работы

Если вы не Power Automate, рекомендуем посетить начало работы [с Power Automate](/power-automate/getting-started). Здесь вы можете узнать больше обо всех доступных вам возможностях автоматизации. В этих документах основное внимание уделяется работе Office скриптов с Power Automate и как это может помочь улучшить Excel работу.

Чтобы приступить к Power Automate и Office скриптов, следуйте учебнику Начните использовать скрипты с [Power Automate](../tutorials/excel-power-automate-manual.md). Это научит вас создавать поток, который вызывает простой сценарий. После завершения этого учебника и передачи данных [](../tutorials/excel-power-automate-trigger.md) скриптам в автоматическом руководстве Power Automate потока возвращайтесь сюда для получения подробных сведений о подключении скриптов Office к Power Automate потокам.

## <a name="excel-online-business-connector"></a>Excel соединители Online (Бизнес)

[Соединители](/connectors/connectors) — это мосты между Power Automate приложениями. [Соединитель Excel Online (Бизнес)](/connectors/excelonlinebusiness) предоставляет вашим потокам доступ к Excel книгам. Действие "Сценарий запуска" позволяет вызывать любой Office, доступный через выбранную книгу. Вы также можете предоставить параметры ввода сценариев, чтобы данные могли быть предоставлены потоком, или иметь сведения о возвращении сценария для последующих действий в потоке.

> [!IMPORTANT]
> Действие "Сценарий запуска" предоставляет людям, которые используют Excel, значительный доступ к вашей книге и ее данным. Кроме того, существуют риски безопасности со скриптами, которые делают внешние вызовы API, как объясняется [в внешних звонках из Power Automate](external-calls.md). Если администратор обеспокоен воздействием высокочувствительных данных, он может отключить соединители Excel Online или ограничить доступ к Office скриптам с помощью элементов управления администратором [Office скриптов](/microsoft-365/admin/manage/manage-office-scripts-settings).

## <a name="data-transfer-in-flows-for-scripts"></a>Передача данных в потоках для скриптов

Power Automate позволяет передавать фрагменты данных между шагами потока. Скрипты можно настроить, чтобы принимать все необходимые типы информации и возвращать все, что требуется в вашей книге. Ввод для скрипта определяется добавлением параметров `main` в функцию (в дополнение к `workbook: ExcelScript.Workbook`). Выход из скрипта объявляется путем добавления типа возврата к `main`.

> [!NOTE]
> При создании блока "Сценарий запуска" в потоке заполняются принятые параметры и возвращаемые типы. Если вы измените параметры или типы возвращаемого скрипта, вам потребуется переоценить блок потока "Запустить сценарий". Это обеспечивает правильную обработку данных.

В следующих разделах подробно освещается ввод и выход сценариев, используемых в Power Automate. Если вы хотите практический подход к изучению этой темы, попробуйте передать данные в [](../tutorials/excel-power-automate-trigger.md) скрипты в автоматическом руководстве по потоку Power Automate или ознакомьтесь с примером сценария автоматизированных [](../resources/scenarios/task-reminders.md) напоминаний задач.

### <a name="main-parameters-pass-data-to-a-script"></a>`main` Параметры: передай данные в скрипт

Все входные данные скрипта указаны в качестве дополнительных параметров для функции `main` . Например, если вы хотите, `string` чтобы сценарий принял имя в качестве ввода, необходимо изменить подпись `main` `function main(workbook: ExcelScript.Workbook, name: string)`на .

При настройке потока в Power Automate можно указать ввод сценария в качестве статических значений, [выражений](/power-automate/use-expressions-in-conditions) или динамического контента. Сведения о соединители отдельной службы можно найти в документации [Power Automate connector](/connectors/).

#### <a name="type-restrictions"></a>Ограничения типа

При добавлении параметров ввода в функцию скрипта `main` рассмотрите следующие ограничения и надбавки. Они также применяются к типу возврата скрипта.

1. Первый параметр должен быть типа `ExcelScript.Workbook`. Его имя параметра не имеет значения.

1. `string`Типы , `number`, `boolean`, `unknown`, и `object``undefined` поддерживаются.

1. Поддерживаются массивы `[]` `Array<T>` (как и стили) перечисленных ранее типов. Вложенные массивы также поддерживаются.

1. Типы Union разрешены, если они являются союзом литералов, принадлежащих к одному типу ( `"Left" | "Right"`например, нет `"Left", 5`). Поддерживаются также союзы поддерживаемого типа с неопределенным (например `string | undefined`).

1. Типы объектов разрешены, если они содержат свойства типа `string`, `number`поддерживаемых `boolean`массивов или других поддерживаемых объектов. В следующем примере показаны вложенные объекты, поддерживаемые в качестве типов параметров.

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

1. Объекты должны иметь свой интерфейс или определение класса, определенные в сценарии. Объект также можно определить анонимно, как в следующем примере.

    ```TypeScript
    function main(workbook: ExcelScript.Workbook): {name: string, email: string}
    ```

#### <a name="optional-and-default-parameters"></a>Необязательные параметры и параметры по умолчанию

1. Необязательные параметры разрешены и обозначаются необязательным модификатором `?` (например, `function main(workbook: ExcelScript.Workbook, Name?: string)`).

1. Разрешены значения параметров по умолчанию (например `function main(workbook: ExcelScript.Workbook, Name: string = 'Jane Doe')`.

### <a name="return-data-from-a-script"></a>Возвращение данных из скрипта

Скрипты могут возвращать данные из книги, которая будет использоваться в качестве динамического контента в потоке Power Automate. Те [же ограничения типа, перечисленные ранее](#type-restrictions) , применяются к типу возврата. Чтобы вернуть объект, добавьте синтаксис типа возврата в функцию `main` . Например, если вы хотите вернуть значение `string` из скрипта, ваша подпись `main` будет `function main(workbook: ExcelScript.Workbook): string`.

## <a name="example"></a>Пример

На следующем снимке экрана Power Automate поток, который запускается при GitHub назначена вам проблема[](https://github.com/). Поток запускает сценарий, который добавляет проблему в таблицу в Excel книге. Если в этой таблице имеется пять или более проблем, поток отправляет напоминание по электронной почте.

:::image type="content" source="../images/power-automate-parameter-return-sample.png" alt-text="Редактор Power Automate потока, показывающий поток примера.":::

Функция `main` скрипта указывает ID проблемы и название выпуска в качестве параметров ввода, и скрипт возвращает количество строк в таблице вопросов.

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

- [Запустите Office скрипты в Excel в Интернете с Power Automate](../tutorials/excel-power-automate-manual.md)
- [Передача данных сценариям в автоматически запускаемых рабочих процессах Power Automate](../tutorials/excel-power-automate-trigger.md)
- [Возвращение данных из сценария в автоматически запускаемый поток Power Automate](../tutorials/excel-power-automate-returns.md)
- [Сведения об устранении неполадок для Power Automate с Office скриптами](../testing/power-automate-troubleshooting.md)
- [Начало работы с Power Automate](/power-automate/getting-started)
- [Excel справочная документация по соединители Online (Business)](/connectors/excelonlinebusiness/)
