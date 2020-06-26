---
title: Интеграция сценариев Office с автоматизацией управления питанием
description: Как получить скрипты Office для Excel в Интернете, работая с рабочими процессами Power Автоматизация.
ms.date: 06/24/2020
localization_priority: Normal
ms.openlocfilehash: 977d9c88d75c8070eb729a443b4e8bc9a32e456d
ms.sourcegitcommit: aec3c971c6640429f89b6bb99d2c95ea06725599
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 06/25/2020
ms.locfileid: "44878850"
---
# <a name="integrate-office-scripts-with-power-automate"></a>Интеграция сценариев Office с автоматизацией управления питанием

[Power автоматизиру](https://flow.microsoft.com) интегрирует ваш сценарий в больший рабочий процесс. Вы можете использовать автоматизацию управления питанием, например добавить содержимое электронной почты в таблицу листа или создать действия в средствах управления проектами на основе комментариев к книгам. Если вы впервые используете автоматизированное управление питанием, рекомендуем [ознакомиться со статьей "начать автоматизацию](/power-automate/getting-started)". Здесь вы можете узнать больше об автоматизации рабочих процессов для нескольких служб.

> [!IMPORTANT]
> В настоящее время вы не можете запускать сценарии Office из [общего потока](/power-automate/share-buttons). Только пользователь, создавший сценарий, может запускать его, даже если вы автоматизируем Power.

## <a name="getting-started"></a>Начало работы

Чтобы приступить к объединению сценариев Power автоматизированного и Office, следуйте рекомендациям, описанным в разделе [starting Scripts with Power Автоматизация](../tutorials/excel-power-automate-manual.md). С его помощью вы узнаете, как создать последовательность, вызывающую простой сценарий. После выполнения этого руководства и [автоматического запуска сценариев с помощью руководства Power Автоматизация](../tutorials/excel-power-automate-trigger.md) вернитесь сюда, чтобы узнать подробности об интеграции платформы.

## <a name="excel-online-business-connector"></a>Соединитель Excel Online (Business)

[Соединители](/connectors/connectors) — это мосты между автоматизированной автоматизацией и приложениями. [Соединитель Excel Online (Business)](/connectors/excelonlinebusiness) предоставляет потокам доступ к книгам Excel. Действие "Запуск скрипта" позволяет вызывать любой сценарий Office, доступный через выбранную книгу. Вы не можете выполнять сценарии с помощью потока, вы можете передавать данные в книгу и из нее с помощью скриптов.

> [!IMPORTANT]
> Действие "Запуск скрипта" дает пользователям, использующим Microsoft Connector, значительный доступ к книге и ее данным. Кроме того, существуют риски, связанные с безопасностью, с помощью скриптов, которые выполняют внешние вызовы API, как описано во [внешних вызовах от автоматизации Powering](external-calls.md). Если администратор имеет дело с очень конфиденциальными данными, он может либо отключить Microsoft Excel Online Connector, либо ограничить доступ к сценариям Office с помощью [сценариев Office](https://support.microsoft.com/office/19d3c51a-6ca2-40ab-978d-60fa49554dcf).

## <a name="passing-data-from-power-automate-into-a-script"></a>Передача данных из Power автоматизировать в сценарий

Все входные данные сценария указываются как дополнительные параметры `main` функции. Например, если вы хотите, чтобы сценарий принимал объект `string` , представляющий имя в качестве входных данных, вы можете изменить `main` подпись на `function main(workbook: ExcelScript.Workbook, name: string)` .

Когда вы настраиваете потоки в Power Автоматизация, вы можете указать входные данные скрипта в виде статических значений, [выражений](/power-automate/use-expressions-in-conditions)или динамического содержимого. Подробные сведения о соединителе отдельных служб можно найти в [документации Power автоматизиру Connector](/connectors/).

При добавлении входных параметров в функцию сценария `main` учитывайте следующие ограничения и ограничения.

1. Первый параметр должен иметь тип `ExcelScript.Workbook` . Имя параметра не имеет значения.

2. Каждый параметр должен иметь тип.

3. Основные типы,,,,,, `string` `number` `boolean` `any` `unknown` `object` и `undefined` поддерживаются.

4. Массивы приведенных выше базовых типов поддерживаются.

5. Вложенные массивы поддерживаются в качестве параметров (но не как типы возвращаемого значения).

6. Типы Union разрешены, если они являются объединением литералов, принадлежащих одному типу ( `string` , `number` или `boolean` ). Также поддерживаются объединения поддерживаемого типа с неопределенными.

7. Типы объектов разрешены, если они содержат свойства типа `string` , `number` , `boolean` , поддерживаемых массивов или других поддерживаемых объектов. В следующем примере показаны вложенные объекты, которые поддерживаются как типы параметров:

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

8. Объекты должны иметь определение интерфейса или класса, определенное в сценарии. Объект также может быть определен анонимно, как показано в следующем примере:

    ```TypeScript
    function main(workbook: ExcelScript.Workbook): {name: string, email: string}
    ```

9. Необязательные параметры разрешены и могут быть отмечены с помощью необязательного модификатора `?` (например, `function main(workbook: ExcelScript.Workbook, Name?: string)` ).

10. Допустимые значения параметров по умолчанию (например `async function main(workbook: ExcelScript.Workbook, Name: string = 'Jane Doe')` ,.

## <a name="returning-data-from-a-script-back-to-power-automate"></a>Возврат данных из скрипта в Power Автоматизация

Скрипты могут возвращать данные из книги для использования в качестве динамического контента в автоматизированном блоке управления питанием. Как и в случае с входными параметрами, Автоматизация управления питанием применяет некоторые ограничения к типу возвращаемого значения.

1. Поддерживаются основные типы,,,, `string` `number` `boolean` `void` и `undefined` .

2. Типы объединения, используемые в качестве возвращаемых типов, действуют с теми же ограничениями, что и при использовании в качестве параметров сценария.

3. Типы массивов разрешены, если они имеют тип `string` , `number` или `boolean` . Они также разрешены, если тип является поддерживаемым объединением или поддерживаемым типом литерала.

4. Типы объектов, используемые в качестве возвращаемых типов, действуют с теми же ограничениями, что и при использовании в качестве параметров сценария.

5. Неявная типизация поддерживается, несмотря на то, что они должны следовать тем же правилам, что и определенный тип.

## <a name="avoid-using-relative-references"></a>Избегайте использования относительных ссылок

Power автоматизирует выполнение вашего сценария в выбранной книге Excel от вашего имени. В этом случае книга может быть закрыта. Любой API, зависящий от текущего состояния пользователя (например `Workbook.getActiveWorksheet` ,), не будет работать при использовании автоматизации Powering. При проектировании скриптов обязательно используйте абсолютные ссылки на листы и диапазоны.

Приведенные ниже функции вызовут ошибку и завершатся ошибкой при вызове из скрипта в блоке автоматизации Power.

- `Chart.activate`
- `Range.select`
- `Workbook.getActiveCell`
- `Workbook.getActiveChart`
- `Workbook.getActiveChartOrNullObject`
- `Workbook.getActiveSlicer`
- `Workbook.getActiveSlicerOrNullObject`
- `Workbook.getActiveWorksheet`
- `Workbook.getSelectedRange`
- `Workbook.getSelectedRanges`
- `Worksheet.activate`

## <a name="example"></a>Пример

На следующем снимке экрана показан процесс автоматизации Power, который срабатывает при назначении вопроса [GitHub](https://github.com/) . Поток выполняет сценарий, который добавляет ошибку в таблицу в книге Excel. Если в этой таблице имеется пять или более проблем, посылается напоминание по электронной почте.

![Пример процесса, показанный в редакторе автоматизации управления питанием.](../images/power-automate-parameter-return-sample.png)

`main`Функция скрипта ЗАДАЕТ идентификатор вопроса и заголовок вопроса в качестве входных параметров, а скрипт возвращает количество строк в таблице "ошибка".

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

- [Запуск сценариев Office в Excel в Интернете с помощью Power автоматизиру](../tutorials/excel-power-automate-manual.md)
- [Автоматический запуск сценариев с помощью автоматизации управления питанием](../tutorials/excel-power-automate-trigger.md)
- [Основы сценариев для сценариев Office в Excel в Интернете](scripting-fundamentals.md)
- [Начало работы с Power Automate](/power-automate/getting-started)
- [Справочная документация по Microsoft Online Connector (бизнес)](/connectors/excelonlinebusiness/)