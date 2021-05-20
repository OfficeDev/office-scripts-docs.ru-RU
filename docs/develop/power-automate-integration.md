---
title: Вы запустите Office скрипты с Power Automate
description: Как получить Office для Excel в Интернете с рабочим Power Automate процесса.
ms.date: 05/17/2021
localization_priority: Normal
ms.openlocfilehash: 7562a2b2359cde67a9a47e0640515018fe23ac35
ms.sourcegitcommit: 4687693f02fc90a57ba30c461f35046e02e6f5fb
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 05/19/2021
ms.locfileid: "52545042"
---
# <a name="run-office-scripts-with-power-automate"></a>Вы запустите Office скрипты с Power Automate

[Power Automate](https://flow.microsoft.com) позволяет добавлять Office скрипты в более крупный автоматизированный рабочий процесс. Вы можете использовать Power Automate делать такие вещи, как добавление содержимого электронной почты в таблицу листа или создавать действия в инструментах управления проектами на основе комментариев к рабочей книге.

## <a name="get-started"></a>Начало работы

Если вы только что Power Automate, мы рекомендуем [посетить Начало работы с Power Automate](/power-automate/getting-started). Там вы можете узнать больше обо всех доступных вам возможностях автоматизации. Документы здесь сосредоточены на том, Office скрипты работают Power Automate и как это может помочь улучшить Excel опыт.

Чтобы начать Power Automate и Office, следуйте [учебнику Начните использовать скрипты с Power Automate.](../tutorials/excel-power-automate-manual.md) Это научит вас, как создать поток, который вызывает простой сценарий. После того как вы завершили этот учебник и [данные Pass для скриптов в автоматическом учебнике потока Power Automate,](../tutorials/excel-power-automate-trigger.md) вернитесь сюда для получения подробной информации о подключении Office scripts к Power Automate потокам.

## <a name="excel-online-business-connector"></a>Excel Онлайн (Бизнес) разъем

[Коннекторы](/connectors/connectors) являются мостами между Power Automate и приложениями. Разъем [Excel Online (Business) предоставляет](/connectors/excelonlinebusiness) вашим потокам доступ к Excel книгам. Действие "Run script" позволяет вызвать любую Office, доступную через выбранную трудовую книжку. Вы также можете предоставить параметры ввода скриптов, чтобы данные могли быть предоставлены потоком, или иметь информацию о возврате скрипта для более поздних шагов в потоке.

> [!IMPORTANT]
> Действие "Run script" дает людям, которые используют Excel разъем значительный доступ к вашей рабочей книге и ее данным. Кроме того, существуют риски безопасности со скриптами, которые делают внешние вызовы API, как это [объясняется во внешних звонках Power Automate.](external-calls.md) Если ваш администратор обеспокоен воздействием высокочувствительных данных, он может либо отключить разъем Excel Online, либо ограничить доступ к скриптам Office [через Office Scripts.](/microsoft-365/admin/manage/manage-office-scripts-settings)

## <a name="data-transfer-in-flows-for-scripts"></a>Передача данных в потоках для скриптов

Power Automate позволяет передавать фрагменты данных между шагами вашего потока. Скрипты могут быть настроены, чтобы принимать любые типы информации, вам нужно, и возвращать что-либо из вашей рабочей книги, что вы хотите в потоке. Ввод скрипта определяется путем добавления параметров к `main` функции (в дополнение `workbook: ExcelScript.Workbook` к). Выход из скрипта объявляется путем добавления типа возврата `main` к .

> [!NOTE]
> При создании блока "Run Script" в потоке заполняется принятые параметры и возвращенные типы. Если вы измените параметры или вернете типы скрипта, вам нужно будет переписать блок потока "Run script". Это гарантирует, что данные правильно разобраются.

Следующие разделы охватывают детали ввода и вывода скриптов, используемых в Power Automate. Если вы хотите практический подход к изучению этой темы, попробуйте [данные Pass для скриптов в автоматическом учебнике по потоку Power Automate](../tutorials/excel-power-automate-trigger.md) или [исследуйте пример автоматического напоминания](../resources/scenarios/task-reminders.md) о задачах.

### <a name="main-parameters-pass-data-to-a-script"></a>`main` Параметры: Переведать данные в скрипт

Все ввода скрипта указаны в качестве дополнительных параметров для `main` функции. Например, если вы хотите, чтобы скрипт принял `string` имя, представляющее его в качестве ввода, вы бы изменили `main` подпись на `function main(workbook: ExcelScript.Workbook, name: string)` .

При настройке потока в Power Automate можно указать ввод скрипта как статические значения, [выражения или](/power-automate/use-expressions-in-conditions)динамическое содержимое. Подробную информацию о разъеме отдельной службы можно найти [в Power Automate Connector.](/connectors/)

При добавлении параметров ввода в `main` функцию скрипта учитывайте следующие надбавки и ограничения.

1. Первый параметр должен быть `ExcelScript.Workbook` типа. Его название параметра не имеет значения.

2. Каждый параметр должен иметь тип (например, `string` или `number` ).

3. Основные типы `string` , , , и `number` `boolean` `unknown` `object` `undefined` поддерживаются.

4. Поддерживаются массивы ранее перечисленных базовых типов.

5. Вложенные массивы поддерживаются в качестве параметров (но не в качестве типов возврата).

6. Типы союзов допускаются, если они являются союзом букватов, принадлежащих к одному типу `"Left" | "Right"` (например). Поддерживаются также союзы поддерживаемого типа с неопределенными `string | undefined` (например).

7. Типы объектов допускаются, если они содержат свойства `string` `number` типа, `boolean` поддерживаемые массивы или другие поддерживаемые объекты. В следующем примере показаны вложенные объекты, которые поддерживаются в качестве типов параметров:

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

8. Объекты должны иметь свой интерфейс или определение класса, определенные в скрипте. Объект также может быть определен анонимно, как в следующем примере:

    ```TypeScript
    function main(workbook: ExcelScript.Workbook): {name: string, email: string}
    ```

9. Дополнительные параметры разрешены и могут быть обозначены как таковые с помощью дополнительного `?` модификатора (например, `function main(workbook: ExcelScript.Workbook, Name?: string)` ).

10. Разрешены значения параметров по умолчанию `async function main(workbook: ExcelScript.Workbook, Name: string = 'Jane Doe')` (например.

### <a name="return-data-from-a-script"></a>Возврат данных из скрипта

Скрипты могут возвращать данные из рабочей книги, которые будут использоваться в качестве динамического содержимого в Power Automate потоке. Как и в случае с параметрами Power Automate, он устанавливает некоторые ограничения на тип возврата.

1. Основные типы `string` , , , и `number` `boolean` `void` `undefined` поддерживаются.

2. Типы союзов, используемые в качестве типов возврата, следуют тем же ограничениям, что и при использовании в качестве параметров скрипта.

3. Типы массивов разрешены, если они `string` `number` типа, или `boolean` . Они также допускаются, если тип поддерживается союзом или поддерживается буквальным типом.

4. Типы объектов, используемые в качестве типов возврата, следуют тем же ограничениям, что и при использовании в качестве параметров скрипта.

5. Неявная ввод поддерживается, хотя она должна следовать тем же правилам, что и определенный тип.

## <a name="example"></a>Пример

На следующем скриншоте показан Power Automate поток, который срабатывает [всякий раз, GitHub](https://github.com/) вам назначается проблема с номером. Поток выполняет скрипт, который добавляет проблему в таблицу в Excel книги. Если в этой таблице есть пять или более проблем, поток отправляет напоминание по электронной почте.

:::image type="content" source="../images/power-automate-parameter-return-sample.png" alt-text="Редактор Power Automate потока, показывающий поток примера":::

Функция `main` скрипта определяет идентификатор проблемы и название вопроса в качестве входных параметров, а скрипт возвращает количество строк в таблице проблем.

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

- [Вы запустите Office сценарии в Excel в Интернете с Power Automate](../tutorials/excel-power-automate-manual.md)
- [Передача данных сценариям в автоматически запускаемых рабочих процессах Power Automate](../tutorials/excel-power-automate-trigger.md)
- [Возвращение данных из сценария в автоматически запускаемый поток Power Automate](../tutorials/excel-power-automate-returns.md)
- [Информация о устранении неполадок для Power Automate с помощью Office скриптов](../testing/power-automate-troubleshooting.md)
- [Начало работы с Power Automate](/power-automate/getting-started)
- [Excel Онлайн (Бизнес) разъем справочная документация](/connectors/excelonlinebusiness/)
