---
title: Преобразование CSV-файлов в Excel книги
description: Узнайте, как использовать Office и Power Automate для создания .xlsx из .csv файлов.
ms.date: 07/19/2021
ms.localizationpriority: medium
ms.openlocfilehash: 213c6caab1d1b20d566aa0e79630c1a9b50554f7
ms.sourcegitcommit: 5ec904cbb1f2cc00a301a5ba7ccb8ae303341267
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/18/2021
ms.locfileid: "59447480"
---
# <a name="convert-csv-files-to-excel-workbooks"></a>Преобразование CSV-файлов в Excel книги

Многие службы экспортируют данные в качестве разделенных запятой файлов значения (CSV). Это решение автоматизирует процесс преобразования этих CSV-файлов в Excel книги в формате .xlsx файла. Он использует [поток Power Automate](https://flow.microsoft.com) для поиска файлов с расширением .csv в папке OneDrive и скрипта Office для копирования данных из файла .csv в новую книгу Excel.

## <a name="solution"></a>Решение

1. Храните .csv и пустой файл "Template" .xlsx в OneDrive папке.
1. Создайте Office для анализа данных CSV в диапазоне.
1. Создайте Power Automate для чтения .csv и передать их содержимое в скрипт.

## <a name="sample-files"></a>Примеры файлов

Скачайте <a href="https://github.com/OfficeDev/office-scripts-docs/blob/master/docs/resources/samples/convert-csv-example.zip?raw=true">convert-csv-example.zip, </a> чтобы получить Template.xlsx и два примера .csv файлов. Извлечение файлов в папку в OneDrive. В этом примере предполагается, что папка называется "выход".

Добавьте следующий сценарий и создайте поток, используя шаги, которые даются для самостоятельной пробы!

## <a name="sample-code-insert-comma-separated-values-into-a-workbook"></a>Пример кода: Вставьте разделенные запятой значения в книгу

```TypeScript
function main(workbook: ExcelScript.Workbook, csv: string) {
  /* Convert the CSV data into a 2D array. */
  // Trim the trailing new line.
  csv = csv.trim();

  // Split each line into a row.
  let rows = csv.split("\r\n");
  let data : string[][] = [];
  rows.forEach((value) => {
    /*
     * For each row, match the comma-separated sections.
     * For more information on how to use regular expressions to parse CSV files,
     * see this Stack Overflow post: https://stackoverflow.com/a/48806378/9227753
     */
    let row = value.match(/(?:,|\n|^)("(?:(?:"")*[^"]*)*"|[^",\n]*|(?:\n|$))/g);
    
    // Remove the preceding comma.
    row.forEach((cell, index) => {
      row[index] = cell.indexOf(",") === 0 ? cell.substr(1) : cell;
    });
    data.push(row);
  });

  // Put the data in the worksheet.
  let sheet = workbook.getWorksheet("Sheet1");
  let range = sheet.getRangeByIndexes(0, 0, data.length, data[0].length);
  range.setValues(data);

  // Add any formatting or table creation that you want.
}
```

## <a name="power-automate-flow-create-new-xlsx-files"></a>Power Automate потока: создание новых .xlsx файлов

1. Вопишите [Power Automate](https://flow.microsoft.com) и создайте новый **поток запланированных облаков.**
1. Установите поток, чтобы **повторить каждый** "1" "День" и выберите **Создать**.
1. Получите файл Excel шаблона. Это основа для всех преобразованных .csv файлов. Добавьте новый **шаг,** использующий **соединителю OneDrive для бизнеса** и действие **контента Get.** Укай путь файла в файл "Template.xlsx".
    * **Файл**: /output/Template.xlsx
1. Переименуй шаг "Получить содержимое файла", переехав в меню Меню для получения контента **файла (...)** этого шага (в правом верхнем углу соединитетеля) и выбрав параметр  **Переименование.** Измените имя шага на "Получить Excel шаблон".

     :::image type="content" source="../../images/convert-csv-flow-1.png" alt-text="Завершенный OneDrive для бизнеса в Power Automate, переименованный в шаблон Get Excel.":::
1. Получите все файлы в папке "выход". Добавьте новый **шаг,** использующий **соединителю OneDrive для бизнеса** и файлы **List в действии папки.** Предостереть путь папки, содержащий .csv файлы.
    * **Папка**: /выход

    :::image type="content" source="../../images/convert-csv-flow-2.png" alt-text="Завершенный OneDrive для бизнеса в Power Automate.":::
1. Добавьте условие, чтобы поток функционировал только на .csv файлах. Добавьте новый **шаг,** который является **управлением условием.** Используйте следующие значения для **условия**.
    * **Выберите значение**: *Name* (динамическое содержимое из файлов Списка **в папке).** Обратите внимание, что этот динамический контент имеет несколько результатов, поэтому **применение**  к каждому контролю значений окружает **условие**.
    * **заканчивается** (из списка отсев)
    * **Выберите значение**: .csv

    :::image type="content" source="../../images/convert-csv-flow-3.png" alt-text="Завершено управление состоянием с помощью применить к каждому из них.":::
1. Остальная часть потока находится в разделе **If Yes,** так как мы хотим действовать только .csv файлах. Получите отдельный .csv, добавив новый  шаг, использующий **соединители** OneDrive для бизнеса и действие **контента Get.** Используйте **Id из** динамического контента из **файлов List в папке**.
    * **Файл**: *Id* (динамическое содержимое из файлов **Списка в шаге папки)**
1. Переименуй новый шаг **Get file content** в "Get .csv файл". Это помогает отличать этот файл от Excel шаблона.
1. Сделайте новый .xlsx файл, используя Excel в качестве базового контента. Добавьте новый **шаг,** использующий **соединителю OneDrive для бизнеса** и действие **Create file.** Используйте следующие значения.
    * **Путь папки:**/выход
    * **Имя файла.** Имя без *.xlsx* (выберите имя без динамического контента расширения из файлов List в папке и вручную введите ".xlsx" после него)  
    * **Содержимое файлов:** *содержимое файлов* (динамическое содержимое из **шаблона Get Excel)**

     :::image type="content" source="../../images/convert-csv-flow-4.png" alt-text="Файл Get .csv и создание действий файла Power Automate потока.":::
1. Запустите скрипт для копирования данных в новую книгу. Добавьте **соединителю Excel Online (Бизнес)** с действием **сценария Run.** Используйте следующие значения для действия.
    * **Расположение**: OneDrive для бизнеса
    * **Библиотека документов**: OneDrive
    * **Файл**: *Id* (динамическое содержимое **из файла Create)**
    * **Сценарий:** преобразование CSV
    * **csv**. *Содержимое файла* (динамическое содержимое из файла **Get .csv)**

    :::image type="content" source="../../images/convert-csv-flow-5.png" alt-text="Завершенный соедините Excel Online (Бизнес) в Power Automate.":::
1. Сохраните поток. Используйте **кнопку Test** на странице редактора потока или запустите поток через вкладку **Мои потоки.** Не забудьте разрешить доступ при запросе.
1. Необходимо найти новые .xlsx в папке "выход", а также исходные .csv файлы. Новые книги содержат те же данные, что и файлы CSV.
