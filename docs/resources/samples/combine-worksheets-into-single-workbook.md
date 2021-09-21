---
title: Объединение книг в одну книгу
description: Узнайте, как использовать Office и Power Automate для создания таблиц слияния из других книг в одну книгу.
ms.date: 09/17/2021
ms.localizationpriority: medium
ms.openlocfilehash: ffb0fd13cf587184aec87ade36e5e0e661043b94
ms.sourcegitcommit: c23816babcc628b52f6d8aaa4b6342e04e83a5bd
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/21/2021
ms.locfileid: "59460787"
---
# <a name="combine-worksheets-into-a-single-workbook"></a>Объединение нескольких книг в одну

В этом примере показано, как вытащить данные из нескольких книг в одну централизованную книгу. Он использует два сценария: один для получения сведений из книги, а другой для создания новых таблиц с этой информацией. Он объединяет скрипты в потоке Power Automate, который действует на всей OneDrive папке.

> [!IMPORTANT]
> Этот пример копирует только значения из других книг. Он не сохраняет форматирование, диаграммы, таблицы или другие объекты.

## <a name="scenario"></a>Сценарий

1. Создайте новый файл Excel в OneDrive и добавьте в него два сценария из этого примера.
1. Создайте папку в OneDrive и добавьте в нее одну или несколько книг с данными.
1. Создайте поток, чтобы получить все файлы этой папки.
1. Используйте скрипт **данных таблицы** Return для получения данных из каждого таблицы в каждой книге.
1. Используйте скрипт **Добавить таблицы** для создания нового таблицы в одной книге для каждого таблицы во всех остальных файлах.

## <a name="sample-code-return-worksheet-data"></a>Пример кода. Возвращаем данные таблицы

```TypeScript
/**
 * This script returns the values from the used ranges on each worksheet.
 */
function main(workbook: ExcelScript.Workbook): WorksheetData[]
{
  // Create an object to return the data from each worksheet.
  let worksheetInformation: WorksheetData[] = [];

  // Get the data from every worksheet, one at a time.
  workbook.getWorksheets().forEach((sheet) => {
    let values = sheet.getUsedRange()?.getValues();
    worksheetInformation.push({
       name: sheet.getName(),
       data: values as string[][]
    });
  });

  return worksheetInformation;
}

// An interface to pass the worksheet name and cell values through a flow.
interface WorksheetData {
  name: string;
  data: string[][];
}
```

## <a name="sample-code-add-worksheets"></a>Пример кода. Добавление таблиц

```TypeScript
/**
 * This script creates a new worksheet in the current workbook for each WorksheetData object provided.
 */
function main(workbook: ExcelScript.Workbook, workbookName: string, worksheetInformation: WorksheetData[])
{
  // Add each new worksheet.
  worksheetInformation.forEach((value) => {
    let sheet = workbook.addWorksheet(`${workbookName}.${value.name}`);

    // If there was any data in the worksheet, add it to a new range.
    if (value.data) {
      let range = sheet.getRangeByIndexes(0, 0, value.data.length, value.data[0].length);
      range.setValues(value.data);
    }
  });
}

// An interface to pass the worksheet name and cell values through a flow.
interface WorksheetData {
  name: string;
  data: string[][];
}
```

## <a name="power-automate-flow-combine-worksheets-into-a-single-workbook"></a>Power Automate потока: объединяйте таблицы в одну книгу

1. Вопишите [Power Automate](https://flow.microsoft.com) и создайте новый поток **мгновенных облаков.**
1. Выберите **вручную вызвать поток и** выберите **Создать**.
1. Получите все файлы в папке. В этом примере мы будем использовать папку с именем "выход". Добавьте новый **шаг,** использующий **соединителю OneDrive для бизнеса** и файлы **List в действии папки.** Предостереть путь папки, содержащий .csv файлы.
    * **Папка**: /выход

    :::image type="content" source="../../images/combine-worksheets-flow-1.png" alt-text="Завершенный OneDrive для бизнеса в Power Automate.":::
1. Запустите сценарий данных **таблицы** Return, чтобы получить все данные из каждой книги. Добавьте **соединителю Excel Online (Бизнес)** с действием **сценария Run.** Используйте следующие значения для действия. Обратите внимание, что при добавлении *Id* для файла Power Automate действие будет завернуться в apply **для** каждого управления, поэтому действие будет выполняться в каждом файле.
    * **Расположение**: OneDrive для бизнеса
    * **Библиотека документов**: OneDrive
    * **Файл**: *Id* (динамическое содержимое из **файлов списка в папке)**
    * **Сценарий.** Возвращаем данные таблицы
1. Запустите **сценарий Добавить таблицы** в созданном Excel файле. Это добавит данные из всех других книг. После предыдущего действия **скрипта Run** и внутри применяйте к каждому из них **соединителю Excel Online (Бизнес)** с действием **сценариев Run.**  Используйте следующие значения для действия.
    * **Расположение**: OneDrive для бизнеса
    * **Библиотека документов**: OneDrive
    * **Файл:** файл
    * **Сценарий:** Добавление таблиц
    * **workbookName**: *Name* (динамическое содержимое из **файлов списка в папке)**
    * **таблицаInformation** (после выбора кнопки **Switch to input entire array** см. примечание ниже следующего изображения): результат (динамическое содержимое из сценария **Run)** 

    :::image type="content" source="../../images/combine-worksheets-flow-2.png" alt-text="Два действия скрипта Run внутри apply to each control.":::
    > [!NOTE]
    > Выберите **кнопку Переключатель для** ввода всего массива, чтобы добавить объект массива непосредственно, а не отдельные элементы для массива.
    >
    > :::image type="content" source="../../images/combine-worksheets-flow-3.png" alt-text="Кнопка для перехода на ввод всего массива в поле ввода поля управления.":::
1. Сохраните поток. Используйте **кнопку Test** на странице редактора потока или запустите поток через вкладку **Мои потоки.** Не забудьте разрешить доступ при запросе.
1. Теперь Excel файл должен иметь новые таблицы.
