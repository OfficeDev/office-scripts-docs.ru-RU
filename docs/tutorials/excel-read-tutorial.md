---
title: Чтение данных книги с помощью сценариев Office в Excel в Интернете
description: Учебник по сценариям Office о чтении данных из книг и их оценке в сценарии.
ms.date: 01/27/2020
localization_priority: Priority
ms.openlocfilehash: 42ed0fe5843a78692f9660b873211e3668702164
ms.sourcegitcommit: b075eed5a6f275274fbbf6d62633219eac416f26
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/10/2020
ms.locfileid: "42700324"
---
# <a name="read-workbook-data-with-office-scripts-in-excel-on-the-web"></a>Чтение данных книги с помощью сценариев Office в Excel в Интернете

В этом учебнике объясняется, как читать данные из книги с помощью сценария Office для Excel в Интернете. После этого вы сможете отредактировать прочитанные данные и вернуть их в книгу.

> [!TIP]
> Если вы только приступили к работе со сценариями Office, рекомендуем начать с учебника [Запись, редактирование и создание сценариев Office в Excel в Интернете](excel-tutorial.md).

## <a name="prerequisites"></a>Необходимые компоненты

[!INCLUDE [Preview note](../includes/preview-note.md)]

Перед началом работы с этим учебником у вас должен быть доступ к сценариям Office. Для этого требуется следующее:

- [Excel в Интернете](https://www.office.com/launch/excel).
- Попросите своего администратора [включить сценарии Office для организации](https://support.office.com/article/office-scripts-settings-in-m365-19d3c51a-6ca2-40ab-978d-60fa49554dcf), в результате чего на ленту добавится вкладка **Автоматизировать**.

> [!IMPORTANT]
> Этот учебник предназначен для пользователей с начальным и средним уровнем знаний по JavaScript или TypeScript. Если вы впервые работаете с JavaScript, рекомендуем прочесть [учебник Mozilla по JavaScript](https://developer.mozilla.org/docs/Web/JavaScript/Guide/Introduction). Чтобы получить дополнительные сведения о среде сценариев, ознакомьтесь со статьей [Сценарии Office в Excel в Интернете](../overview/excel.md).

## <a name="read-a-cell"></a>Чтение ячейки

Сценарии, созданные с помощью средства записи действий, могут только записывать информацию в книгу. С помощью редактора кода можно редактировать и создавать сценарии, которые также читают данные из книги.

Давайте создадим сценарий, читающий данные и действующий на основе прочитанного. Мы будем работать с примером банковской выписки. Эта выписка объединяет чековую выписку и выписку по кредиту. К сожалению, изменения баланса в них указываются по-разному. В чековой выписке доходы указываются как положительный кредит, а расходы — в виде отрицательного дебета. В выписке по кредиту все наоборот.

В остальной части учебника мы нормализуем эти данные с помощью сценария. Сначала научимся читать данные из книги.

1. Создайте лист в книге, которую вы использовали в остальной части учебника.
2. Скопируйте следующие данные и вставьте их на новый лист, начиная с ячейки **A1**.

    |Дата |Счет |Описание |Дебет |Кредит |
    |:--|:--|:--|:--|:--|
    |10.10.2019 |Чековый |Виноградник Coho |–20,05 | |
    |11.10.2019 |Кредитный |Телефонная компания |99,95 | |
    |13.10.2019 |Кредитный |Виноградник Coho |154,43 | |
    |15.10.2019 |Чековый |Внешний депозит | |1000 |
    |20.10.2019 |Кредитный |Виноградник Coho — возмещение | |–35,45 |
    |25.10.2019 |Чековый |Органическая компания "Лучшее для вас" | –85,64 | |
    |01.11.2019 |Чековый |Внешний депозит | |1000 |

3. Откройте **Редактор кода** и выберите **Создать сценарий**.
4. Давайте очистим форматирование. Это финансовый документ, поэтому изменим числовой формат в столбцах **Дебет** и **Кредит**, чтобы отобразить значения в долларах. Также настроим ширину столбца по данным.

    Замените содержимое сценария следующим кодом:

    ```TypeScript
    async function main(context: Excel.RequestContext) {
      // Get the current worksheet.
      let workbook = context.workbook;
      let worksheets = workbook.worksheets;
      let selectedSheet = worksheets.getActiveWorksheet();

      // Format the range to display numerical dollar amounts.
      selectedSheet.getRange("D2:E8").numberFormat = [["$#,##0.00"]];

      // Fit the width of all the used columns to the data.
      selectedSheet.getUsedRange().format.autofitColumns();
    }
    ```

5. Теперь прочитаем значение в одном из числовых столбцов. Добавьте следующий код в конце сценария:

    ```TypeScript
    // Get the value of cell D2.
    let range = selectedSheet.getRange("D2");
    range.load("values");
    await context.sync();
  
    // Print the value of D2.
    console.log(range.values);
    ```

    Обратите внимание на вызовы `load` и `sync`. Подробные сведения об этих методах можно найти в статье [Основные сведения о сценариях Office в Excel в Интернете](../develop/scripting-fundamentals.md#sync-and-load). Пока же учитывайте, что требуется запросить данные для чтения, а затем синхронизировать сценарий с книгой, чтобы прочесть эти данные.

6. Запустите сценарий.
7. Откройте консоль. Откройте меню **Многоточие** и нажмите **Журналы...**.
8. В консоли должно отображаться следующее: `[Array[1]]`. Это не число, так как диапазоны являются двухмерными массивами данных. Этот двухмерный диапазон напрямую регистрируется в консоли. К счастью, редактор кода позволяет просмотреть содержимое массива.
9. Когда двухмерный массив регистрируется в консоли, она группирует значения столбцов под каждой строкой. Разверните журнал массива, нажав синий треугольник.
10. Разверните второй уровень массива, нажав появившийся синий треугольник. Должно отобразиться следующее:

    ![Журнал консоли, отображающий результат "–20,05", размещенный под двумя массивами.](../images/tutorial-4.png)

## <a name="modify-the-value-of-a-cell"></a>Изменение значения ячейки

Теперь, когда мы можем читать данные, воспользуемся ими, чтобы изменить книгу. Мы сделаем значение ячейки **D2** положительным с помощью функции `Math.abs`. Объект [Math](https://developer.mozilla.org/docs/web/javascript/reference/global_objects/math) содержит множество функций, к которым имеют доступ сценарии. Дополнительные сведения о `Math` и других встроенных объектах можно найти в статье [Использование встроенных объектов JavaScript в сценариях Office](../develop/javascript-objects.md).

1. Добавьте следующий код в конце сценария:

    ```TypeScript
    // Run the `Math.abs` function with the value at D2 and apply that value back to D2.
    let positiveValue = Math.abs(range.values[0][0]);
    range.values = [[positiveValue]];
    ```

2. Значение ячейки **D2** теперь должно быть положительным.

## <a name="modify-the-values-of-a-column"></a>Изменение значений столбца

Теперь, когда вы знаете, как читать и записывать данные в одной ячейке, давайте обобщим сценарий для работы со всеми столбцами **Дебет** и **Кредит**.

1. Удалите код, влияющий только на одну ячейку (предыдущий код с абсолютным значением), чтобы ваш сценарий выглядел следующим образом:

    ```TypeScript
    async function main(context: Excel.RequestContext) {
      // Get the current worksheet.
      let workbook = context.workbook;
      let worksheets = workbook.worksheets;
      let selectedSheet = worksheets.getActiveWorksheet();

      // Format the range to display numerical dollar amounts.
      selectedSheet.getRange("D2:E8").numberFormat = [["$#,##0.00"]];

      // Fit the width of all the used columns to the data.
      selectedSheet.getUsedRange().format.autofitColumns();
    }
    ```

2. Добавьте цикл, выполняющий итерацию в строках двух последних столбцов. Для каждой ячейки сценарий устанавливает текущее абсолютное значение.

    Обратите внимание, что массив, определяющий расположения ячеек, отсчитывается от нуля. Это означает, что ячейка **A1** имеет значение `range[0][0]`.

    ```TypeScript
    // Get the values of the used range.
    let range = selectedSheet.getUsedRange();
    range.load("rowCount,values");
    await context.sync();

    // Iterate over the fourth and fifth columns and set their values to their absolute value.
    for (let i = 1; i < range.rowCount; i++) {
      // The column at index 3 is column "4" in the worksheet.
      if (range.values[i][3] != 0) {
        let positiveValue = Math.abs(range.values[i][3]);
        selectedSheet.getCell(i, 3).values = [[positiveValue]];
      }

      // The column at index 4 is column "5" in the worksheet.
      if (range.values[i][4] != 0) {
        let positiveValue = Math.abs(range.values[i][4]);
        selectedSheet.getCell(i, 4).values = [[positiveValue]];
      }
    }
    ```

    Эта часть сценария выполняет несколько важных задач. Сначала она загружает значения и количество строк используемого диапазона. Это позволяет просматривать значения и определять момент остановки. Затем выполняется итерация в используемом диапазоне с проверкой каждой ячейки в столбцах **Дебет** или **Кредит**. Наконец, если значение в ячейке не равно 0, оно заменяется абсолютным значением. Мы избегаем использования нулей, поэтому можно оставить пустые ячейки неизменными.

3. Запустите сценарий.

    Теперь банковская выписка должна выглядеть следующим образом:

    ![Банковская выписка в виде отформатированной таблицы только с положительными значениями.](../images/tutorial-5.png)

## <a name="next-steps"></a>Дальнейшие действия

Откройте редактор кода и попробуйте некоторые [примеры сценариев Office в Excel в Интернете](../resources/excel-samples.md). Дополнительные сведения о создании сценариев Office доступны в статье [Основные сведения о сценариях Office в Excel в Интернете](../develop/scripting-fundamentals.md).