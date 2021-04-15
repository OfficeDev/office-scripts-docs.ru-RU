---
title: Запуск сценариев Office с помощью power Automate
description: Как получить скрипты Office для Excel в Интернете, работая с рабочим процессом Power Automate.
ms.date: 12/16/2020
localization_priority: Normal
ms.openlocfilehash: 1ca9aa14efe7cf2c91100a32fbc9a69054012f06
ms.sourcegitcommit: 45ffe3dbd2c834b78592ad35928cf8096f5e80bc
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 04/14/2021
ms.locfileid: "51755072"
---
# <a name="run-office-scripts-with-power-automate"></a>Запуск сценариев Office с помощью power Automate

[Power Automate](https://flow.microsoft.com) позволяет добавлять скрипты Office в более крупный автоматизированный рабочий процесс. Вы можете использовать Power Automate для таких действий, как добавление содержимого электронной почты в таблицу таблицы или создание действий в средствах управления проектами на основе комментариев к книгам.

## <a name="getting-started"></a>Начало работы

Если вы только что приступили к power Automate, рекомендуем посетить [Get started with Power Automate.](/power-automate/getting-started) Здесь вы можете узнать больше обо всех доступных вам возможностях автоматизации. В этих документах основное внимание уделяется работе сценариев Office с Power Automate и улучшению работы с Excel.

Чтобы приступить к объединению сценариев Power Automate и Office, следуйте учебнику [Начните с помощью скриптов с power Automate.](../tutorials/excel-power-automate-manual.md) Это научит вас создавать поток, который вызывает простой сценарий. После завершения этого учебника и передачи данных скриптам в руководстве по автоматическому запуску потока power Automate возвращайтесь сюда для получения подробных сведений о подключении скриптов Office к потокам Power [Automate.](../tutorials/excel-power-automate-trigger.md)

## <a name="excel-online-business-connector"></a>Соединитель Excel Online (Бизнес)

[Соединители](/connectors/connectors) — это мосты между Power Automate и приложениями. Соединитель [Excel Online (Бизнес)](/connectors/excelonlinebusiness) предоставляет вашим потокам доступ к книгам Excel. Действие "Сценарий запуска" позволяет вызывать любой скрипт Office, доступный через выбранную книгу. Вы также можете предоставить параметры ввода сценариев, чтобы данные могли быть предоставлены потоком, или иметь сведения о возвращении сценария для последующих действий в потоке.

> [!IMPORTANT]
> Действие "Сценарий запуска" предоставляет людям, которые используют соединитель Excel, значительный доступ к вашей книге и ее данным. Кроме того, существуют риски безопасности со скриптами, которые делают внешние вызовы API, как объясняется в внешних звонках [из Power Automate.](external-calls.md) Если администратор обеспокоен воздействием высокочувствительных данных, он может отключить соединитель Excel Online или ограничить доступ к скриптам Office с помощью элементов управления администратором [office Scripts.](/microsoft-365/admin/manage/manage-office-scripts-settings)

## <a name="data-transfer-in-flows-for-scripts"></a>Передача данных в потоках для скриптов

Power Automate позволяет передавать фрагменты данных между шагами потока. Скрипты можно настроить, чтобы принимать все необходимые типы информации и возвращать все, что требуется в вашей книге. Ввод для скрипта определяется добавлением параметров в `main` функцию (в дополнение `workbook: ExcelScript.Workbook` к). Выход из скрипта объявляется путем добавления типа возврата `main` к .

> [!NOTE]
> При создании блока "Сценарий запуска" в потоке заполняются принятые параметры и возвращаемые типы. Если вы измените параметры или типы возвращаемого скрипта, вам потребуется переоценить блок потока "Запустить сценарий". Это обеспечивает правильную обработку данных.

В следующих разделах подробно освещается ввод и выход сценариев, используемых в Power Automate. Если вы хотите практический подход к изучению этой темы, попробуйте передать данные в скрипты в автоматически [](../resources/scenarios/task-reminders.md) запускаемом руководстве по потоку [Power Automate](../tutorials/excel-power-automate-trigger.md) или ознакомьтесь с примером сценария автоматизированных напоминаний задач.

### <a name="main-parameters-passing-data-to-a-script"></a>`main` Параметры: передача данных в скрипт

Все входные данные скрипта указаны в качестве дополнительных параметров для `main` функции. Например, если вы хотите, чтобы сценарий принял имя в качестве ввода, необходимо изменить подпись на `string` `main` `function main(workbook: ExcelScript.Workbook, name: string)` .

При настройке потока в Power Automate можно указать ввод скрипта в качестве статических значений, [выражений](/power-automate/use-expressions-in-conditions)или динамического контента. Сведения о соединители отдельной службы можно найти в документации [power Automate Connector.](/connectors/)

При добавлении параметров ввода в функцию скрипта рассмотрите следующие ограничения и `main` надбавки.

1. Первый параметр должен быть типа `ExcelScript.Workbook` . Его имя параметра не имеет значения.

2. Каждый параметр должен иметь тип (например, `string` или `number` ).

3. Основные типы `string` , , , , , и `number` `boolean` `any` `unknown` `object` `undefined` поддерживаются.

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

### <a name="returning-data-from-a-script"></a>Возвращение данных из скрипта

Скрипты могут возвращать данные из книги, которая используется в качестве динамического контента в потоке Power Automate. Как и в отношении параметров ввода, Power Automate вводит некоторые ограничения для возвращаемого типа.

1. Основные типы `string` `number` , , и `boolean` `void` `undefined` поддерживаются.

2. Типы Union, используемые в качестве типов возврата, следуют тем же ограничениям, что и при их использования в качестве параметров скрипта.

3. Типы массивов разрешены, если они имеют тип `string` `number` , или `boolean` . Они также разрешены, если тип является поддерживаемым или поддерживаемым литеральным типом.

4. Типы объектов, используемые в качестве типов возвращаемой, следуют тем же ограничениям, что и при их использования в качестве параметров скрипта.

5. Неявный ввод поддерживается, хотя он должен следовать тем же правилам, что и определенный тип.

## <a name="example"></a>Пример

На следующем скриншоте показан поток power Automate, который запускается при назначенной вам проблеме [GitHub.](https://github.com/) Поток запускает сценарий, который добавляет проблему в таблицу в книге Excel. Если в этой таблице имеется пять или более проблем, поток отправляет напоминание по электронной почте.

:::image type="content" source="../images/power-automate-parameter-return-sample.png" alt-text="Редактор потока Power Automate, показывающий поток примера.":::

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

- [Запуск скриптов Office в Excel в Интернете с помощью power Automate](../tutorials/excel-power-automate-manual.md)
- [Передача данных сценариям в автоматически запускаемых рабочих процессах Power Automate](../tutorials/excel-power-automate-trigger.md)
- [Возвращение данных из сценария в автоматически запускаемый поток Power Automate](../tutorials/excel-power-automate-returns.md)
- [Сведения об устранении неполадок для power Automate с помощью скриптов Office](../testing/power-automate-troubleshooting.md)
- [Начало работы с Power Automate](/power-automate/getting-started)
- [Справочная документация по соединители Excel Online (Бизнес)](/connectors/excelonlinebusiness/)
