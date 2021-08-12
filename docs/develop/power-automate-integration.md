---
title: Запустите Office скрипты с Power Automate
description: Как получить Office скрипты для Excel в Интернете работы с рабочим Power Automate рабочим процессом.
ms.date: 05/17/2021
localization_priority: Normal
ms.openlocfilehash: 61b43904cbc46b97a0102230c9c87c1051edd1516668f42fbded63c53c958de9
ms.sourcegitcommit: 75f7ed8c2d23a104acc293f8ce29ea580b4fcdc5
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 08/11/2021
ms.locfileid: "57846521"
---
# <a name="run-office-scripts-with-power-automate"></a>Запустите Office скрипты с Power Automate

[Power Automate](https://flow.microsoft.com) позволяет добавлять Office скрипты в более крупный автоматизированный рабочий процесс. Вы можете Power Automate что-то вроде добавления содержимого электронной почты в таблицу таблицы или создания действий в средствах управления проектами на основе комментариев к книгам.

## <a name="get-started"></a>Начало работы

Если вы не Power Automate, рекомендуем посетить Начало работы [с Power Automate](/power-automate/getting-started). Здесь вы можете узнать больше обо всех доступных вам возможностях автоматизации. В этих документах основное внимание уделяется работе Office скриптов с Power Automate и как это может помочь улучшить Excel работу.

Чтобы приступить к Power Automate и Office скриптов, следуйте учебнику [Начните](../tutorials/excel-power-automate-manual.md)использовать скрипты с Power Automate . Это научит вас создавать поток, который вызывает простой сценарий. После завершения этого учебника [](../tutorials/excel-power-automate-trigger.md) и передачи данных скриптам в руководстве Power Automate потока автоматически возвращайтесь сюда для получения подробных сведений о подключении Office скриптов к Power Automate потокам.

## <a name="excel-online-business-connector"></a>Excel Соединитетор Online (Business)

[Соединители](/connectors/connectors) — это мосты между Power Automate и приложениями. [Соединитель Excel Online (Бизнес)](/connectors/excelonlinebusiness) предоставляет вашим потокам доступ к Excel книгам. Действие "Сценарий запуска" позволяет вызывать любой Office, доступный через выбранную книгу. Вы также можете предоставить параметры ввода сценариев, чтобы данные могли быть предоставлены потоком, или иметь сведения о возвращении сценария для последующих действий в потоке.

> [!IMPORTANT]
> Действие "Запуск скрипта" предоставляет людям, Excel соединитетелем, значительный доступ к вашей книге и ее данным. Кроме того, существуют риски безопасности со скриптами, которые делают внешние вызовы API, как объясняется в внешних звонках [из Power Automate.](external-calls.md) Если администратор обеспокоен воздействием высокочувствительных данных, он может отключить соединители Excel Online или ограничить доступ к скриптам Office с помощью элементов управления Office [скриптов.](/microsoft-365/admin/manage/manage-office-scripts-settings)

## <a name="data-transfer-in-flows-for-scripts"></a>Передача данных в потоках для скриптов

Power Automate позволяет передавать фрагменты данных между шагами потока. Скрипты можно настроить, чтобы принимать все необходимые типы информации и возвращать все, что требуется в вашей книге. Ввод для скрипта определяется добавлением параметров в `main` функцию (в дополнение `workbook: ExcelScript.Workbook` к). Выход из скрипта объявляется путем добавления типа возврата `main` к .

> [!NOTE]
> При создании блока "Сценарий запуска" в потоке заполняются принятые параметры и возвращаемые типы. Если вы измените параметры или типы возвращаемого скрипта, вам потребуется переоценить блок потока "Запустить сценарий". Это обеспечивает правильную обработку данных.

В следующих разделах подробно освещается ввод и выход сценариев, используемых в Power Automate. Если вы хотите практический подход к изучению этой [](../tutorials/excel-power-automate-trigger.md) темы, попробуйте передать данные в скрипты в автоматическом [](../resources/scenarios/task-reminders.md) руководстве по потоку Power Automate или изучите пример примера автоматических напоминаний задач.

### <a name="main-parameters-pass-data-to-a-script"></a>`main` Параметры: передай данные в скрипт

Все входные данные скрипта указаны в качестве дополнительных параметров для `main` функции. Например, если вы хотите, чтобы сценарий принял имя в качестве ввода, необходимо изменить подпись на `string` `main` `function main(workbook: ExcelScript.Workbook, name: string)` .

При настройке потока в Power Automate можно указать ввод сценария в качестве [](/power-automate/use-expressions-in-conditions)статических значений, выражений или динамического контента. Сведения о соединители отдельной службы можно найти в документации [Power Automate connector.](/connectors/)

При добавлении параметров ввода в функцию скрипта рассмотрите следующие ограничения и `main` надбавки.

1. Первый параметр должен быть типа `ExcelScript.Workbook` . Его имя параметра не имеет значения.

2. Каждый параметр должен иметь тип (например, `string` или `number` ).

3. Основные типы `string` , , , , и `number` `boolean` `unknown` `object` `undefined` поддерживаются.

4. Поддерживаются массивы перечисленных ранее базовых типов.

5. Вложенные массивы поддерживаются в качестве параметров (но не как типы возврата).

6. Типы Union разрешены, если они являются союзом литералов, принадлежащих к одному типу `"Left" | "Right"` (например). Поддерживаются также союзы поддерживаемого типа с неопределенным `string | undefined` (например).

7. Типы объектов разрешены, если они содержат свойства типа `string` `number` , поддерживаемых `boolean` массивов или других поддерживаемых объектов. В следующем примере показаны вложенные объекты, поддерживаемые в качестве типов параметров:

    ```TypeScript
    // Office Scripts can return an Employee object because Position only contains strings and numbers.
    interface Employee {
        name: string;
        job: Position;
    }

    interface Position {
        id: number;
        title: string;
    }
    ```

8. Объекты должны иметь свой интерфейс или определение класса, определенные в сценарии. Объект также можно определить анонимно, как в следующем примере:

    ```TypeScript
    function main(workbook: ExcelScript.Workbook): {name: string, email: string}
    ```

9. Необязательные параметры разрешены и могут быть обозначаться как таковые с помощью дополнительного модификатора `?` (например, `function main(workbook: ExcelScript.Workbook, Name?: string)` ).

10. Разрешены значения параметров по умолчанию `async function main(workbook: ExcelScript.Workbook, Name: string = 'Jane Doe')` (например.

### <a name="return-data-from-a-script"></a>Возвращение данных из скрипта

Скрипты могут возвращать данные из книги, которая будет использоваться в качестве динамического контента в потоке Power Automate. Как и в отношении параметров ввода, Power Automate вводит некоторые ограничения для возвращаемого типа.

1. Основные типы `string` `number` , , и `boolean` `void` `undefined` поддерживаются.

2. Типы Union, используемые в качестве типов возврата, следуют тем же ограничениям, что и при их использования в качестве параметров скрипта.

3. Типы массивов разрешены, если они имеют тип `string` `number` , или `boolean` . Они также разрешены, если тип является поддерживаемым или поддерживаемым литеральным типом.

4. Типы объектов, используемые в качестве типов возвращаемой, следуют тем же ограничениям, что и при их использования в качестве параметров скрипта.

5. Неявный ввод поддерживается, хотя он должен следовать тем же правилам, что и определенный тип.

## <a name="example"></a>Пример

На следующем скриншоте показан Power Automate, который запускается [](https://github.com/) при GitHub назначена вам проблема. Поток запускает сценарий, который добавляет проблему в таблицу в Excel книге. Если в этой таблице имеется пять или более проблем, поток отправляет напоминание по электронной почте.

:::image type="content" source="../images/power-automate-parameter-return-sample.png" alt-text="Редактор Power Automate потока, показывающий поток примера.":::

Функция скрипта указывает ID проблемы и название выпуска в качестве параметров ввода, и скрипт возвращает количество строк в `main` таблице вопросов.

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
- [Excel Справочная документация по соединители online (Business)](/connectors/excelonlinebusiness/)
