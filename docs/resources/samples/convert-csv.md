---
title: Преобразование CSV-файлов в Excel книги
description: Узнайте, как использовать Office и Power Automate для создания .xlsx из .csv файлов.
ms.date: 03/28/2022
ms.localizationpriority: medium
ms.openlocfilehash: 52619c1867b654fae3fce1a383a612f81f80d868
ms.sourcegitcommit: 7023b9e23499806901a5ecf8ebc460b76887cca6
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/31/2022
ms.locfileid: "64585592"
---
# <a name="convert-csv-files-to-excel-workbooks"></a>Преобразование CSV-файлов в Excel книги

Многие службы экспортируют данные в качестве разделенных запятой файлов значения (CSV). Это решение автоматизирует процесс преобразования этих CSV-файлов в Excel книг в формате .xlsx файла. Он использует [поток Power Automate](https://flow.microsoft.com) для поиска файлов с расширением .csv в папке OneDrive и скрипта Office для копирования данных из файла .csv в новую книгу Excel.

## <a name="solution"></a>Решение

1. Храните .csv и пустой файл "Template" .xlsx в OneDrive папке.
1. Создайте Office для анализа данных CSV в диапазоне.
1. Создайте Power Automate для чтения .csv и передать их содержимое в скрипт.

## <a name="sample-files"></a>Примеры файлов

<a href="https://github.com/OfficeDev/office-scripts-docs/blob/master/docs/resources/samples/convert-csv-example.zip?raw=true"> Скачайтеconvert-csv-example.zip</a>, чтобы получить Template.xlsx и два примера .csv файлов. Извлечение файлов в папку в OneDrive. В этом примере предполагается, что папка называется "выход".

Добавьте следующий сценарий и создайте поток, используя шаги, которые даются для самостоятельной пробы!

## <a name="sample-code-insert-comma-separated-values-into-a-workbook"></a>Пример кода: Вставьте разделенные запятой значения в книгу

```TypeScript
/**
 * Convert incoming CSV data into a range and add it to the workbook.
 */
function main(workbook: ExcelScript.Workbook, csv: string) {
  let sheet = workbook.getWorksheet("Sheet1");

  // Remove any Windows \r characters.
  csv = csv.replace(/\r/g, "");

  // Split each line into a row.
  let rows = csv.split("\n");
  /*
   * For each row, match the comma-separated sections.
   * For more information on how to use regular expressions to parse CSV files,
   * see this Stack Overflow post: https://stackoverflow.com/a/48806378/9227753
   */
  const csvMatchRegex = /(?:,|\n|^)("(?:(?:"")*[^"]*)*"|[^",\n]*|(?:\n|$))/g
  rows.forEach((value, index) => {
    if (value.length > 0) {
        let row = value.match(csvMatchRegex);
    
        // Check for blanks at the start of the row.
        if (row[0].charAt(0) === ',') {
          row.unshift("");
        }
    
        // Remove the preceding comma.
        row.forEach((cell, index) => {
          row[index] = cell.indexOf(",") === 0 ? cell.substr(1) : cell;
        });
    
        // Create a 2D array with one row.
        let data: string[][] = [];
        data.push(row);
    
        // Put the data in the worksheet.
        let range = sheet.getRangeByIndexes(index, 0, 1, data[0].length);
        range.setValues(data);
    }
  });

  // Add any formatting or table creation that you want.
}
```

## <a name="power-automate-flow-create-new-xlsx-files"></a>Power Automate потока: создание новых .xlsx файлов

1. Вопишите [Power Automate](https://flow.microsoft.com) и создайте новый **поток запланированных облаков**.
1. Установите поток, чтобы **повторить каждый** "1" "День" и выберите **Создать**.
1. Получите файл Excel шаблона. Это основа для всех преобразованных .csv файлов. Добавьте новый **шаг,** использующий **соединителю OneDrive для бизнеса** и действие **контента Get.** Предокабъем путь к файлу "Template.xlsx".
    * **Файл**: /output/Template.xlsx
1. Переименуй шаг "Получить содержимое файла", переехав в меню Меню для получения контента **файла (...)** этого шага (в правом верхнем углу соединитетеля) и выбрав параметр **Переименование**. Измените имя шага на "Получить Excel шаблон".

     :::image type="content" source="../../images/convert-csv-flow-1.png" alt-text="Завершенный OneDrive для бизнеса в Power Automate, переименованный в шаблон Get Excel.":::
1. Получите все файлы в папке "выход". Добавьте новый **шаг,** использующий **соединителю OneDrive для бизнеса** и файлы **List в действии папки**. Предостереть путь папки, содержащий .csv файлы.
    * **Папка**: /выход

    :::image type="content" source="../../images/convert-csv-flow-2.png" alt-text="Завершенный OneDrive для бизнеса в Power Automate.":::
1. Добавьте условие, чтобы поток функционировал только на .csv файлах. Добавьте новый **шаг,** который является **управлением условием** . Используйте следующие значения для **условия**.
    * **Выберите значение**: *Name* (динамическое содержимое из **файлов Списка в папке**). Обратите внимание, что этот динамический контент имеет несколько результатов, поэтому **применение**  к каждому контролю значений окружает **условие**.
    * **заканчивается** (из списка отсев)
    * **Выберите значение**: .csv

    :::image type="content" source="../../images/convert-csv-flow-3.png" alt-text="Завершено управление состоянием с помощью применить к каждому из них.":::
1. Остальная часть потока находится в разделе **If Yes** , так как мы хотим действовать только .csv файлах. Получите отдельный .csv, добавив новый шаг с OneDrive для бизнеса соединитетелем и **действием контента Get**.  Используйте **Id из** динамического контента из **файлов List в папке**.
    * **Файл**: *Id* (динамическое содержимое из файлов **Списка в шаге папки** )
1. Переименуй новый шаг **Get file content** в "Get .csv файл". Это помогает отличать этот файл от Excel шаблона.
1. Сделайте новый .xlsx, используя шаблон Excel в качестве базового контента. Добавьте новый **шаг,** использующий **соединителю OneDrive для бизнеса** и действие **Create file**. Используйте следующие значения.
    * **Путь папки**: /выход
    * **Имя файла***. Имя* без расширения.xlsx (выберите имя без динамического  контента расширения из файлов List в  папке и вручную введите ".xlsx" после него)
    * **Содержимое файлов**: *содержимое файла* (динамическое содержимое из **шаблона Get Excel)**

     :::image type="content" source="../../images/convert-csv-flow-4.png" alt-text="Файл Get .csv и создание действий файла Power Automate потока.":::
1. Запустите скрипт для копирования данных в новую книгу. Добавьте **соединителю Excel Online (Бизнес)** с действием **сценария Run**. Используйте следующие значения для действия.
    * **Расположение**: OneDrive для бизнеса
    * **Библиотека документов**: OneDrive
    * **Файл**: *Id* (динамическое содержимое **из файла Create)**
    * **Сценарий**: преобразование CSV
    * **csv**. *Содержимое файла* (динамическое содержимое **из файла Get .csv**)

    :::image type="content" source="../../images/convert-csv-flow-5.png" alt-text="Завершенный соедините Excel Online (Бизнес) в Power Automate.":::
1. Сохраните поток. Используйте **кнопку Test** на странице редактора потока или запустите поток через вкладку **Мои потоки** . Не забудьте разрешить доступ при запросе.
1. Вы должны найти новые .xlsx файлы в папке "выход", а также исходные .csv файлы. Новые книги содержат те же данные, что и файлы CSV.

## <a name="troubleshooting"></a>Устранение неполадок

### <a name="script-testing"></a>Тестирование скриптов

Чтобы проверить сценарий без Power Automate, назначьте `csv` значение перед его использованием. Попробуйте добавить следующий код в качестве первой строки `main` функции и нажатия **run**.

```TypeScript
  csv = `1, 2, 3
         4, 5, 6
         7, 8, 9`;
```

### <a name="semicolon-separated-files-and-other-alternative-separators"></a>Разделенные за полуколоном файлы и другие альтернативные разделиторы

В некоторых регионах для раздельного (';') значения ячейки вместо запятых используются запятые. В этом случае необходимо изменить следующие строки в скрипте.

1. Замените запятые на запятые в обычном выражении. Это начинается с `let row = value.match`.

    ```TypeScript
    let row = value.match(/(?:;|\n|^)("(?:(?:"")*[^"]*)*"|[^";\n]*|(?:\n|$))/g);
    ```

1. Замените запятую полуколоном в чеке для пустой первой ячейки. Это начинается с `if (row[0].charAt(0)`.

    ```TypeScript
    if (row[0].charAt(0) === ';') {
    ```

1. Замените запятую полуколоном в строке, которая удаляет символ разделения из отображаемой строки. Это начинается с `row[index] = cell.indexOf`.

   ```TypeScript
      row[index] = cell.indexOf(";") === 0 ? cell.substr(1) : cell;
    ```

> [!NOTE]
> Если в файле используются вкладки или любой другой символ для раздельного использования значений, `;` `\t` замените вышеуказанные замены или любой используемый символ.

### <a name="large-csv-files"></a>Большие CSV-файлы

Если в вашем файле сотни тысяч ячеек, можно достичь Excel [передачи данных](../../testing/platform-limits.md#excel). Вам потребуется периодически принудить сценарий синхронизироваться с Excel. Самый простой способ сделать это — вызвать после `console.log` обработки пакета строк. Добавьте следующие строки кода, чтобы это произошло.

1. Прежде `rows.forEach((value, index) => {`чем добавить следующую строку.

    ```TypeScript
      let rowCount = 0;
    ```

1. После `range.setValues(data);`этого добавьте следующий код. Обратите внимание, что в зависимости от количества столбцов может потребоваться `5000` уменьшить число столбцов.

    ```TypeScript
      rowCount++;
      if (rowCount % 5000 === 0) {
        console.log("Syncing 5000 rows.");
      }
    ```

> [!WARNING]
> Если ваш CSV-файл очень большой, у вас могут возникнуть проблемы с [синхронизацией Power Automate](../../testing/platform-limits.md#power-automate). Необходимо разделить данные CSV на несколько файлов, прежде чем преобразовывать их в Excel книги.
