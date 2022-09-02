---
title: Запись повседневных изменений в Excel и создание отчетов о них с помощью потока Power Automate
description: Узнайте, как использовать сценарии Office и Power Automate для отслеживания изменений значений в книге
ms.date: 08/25/2022
ms.localizationpriority: medium
ms.openlocfilehash: 083ca08573db060aa4788aea58fc67e50d004a4b
ms.sourcegitcommit: a6504f8b0d6b717457c6e0b5306c35ad3900914e
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/02/2022
ms.locfileid: "67572662"
---
# <a name="record-day-to-day-changes-in-excel-and-report-them-with-a-power-automate-flow"></a>Запись повседневных изменений в Excel и создание отчетов о них с помощью потока Power Automate

Сценарии Power Automate и Office объединяются для обработки повторяющихся задач. В этом примере вам показано, как ежедневно записывать одно числовое чтение в книге и сообщать об изменениях со вчера. Вы создадите поток для чтения, записываете его в книгу и сообщаете об изменениях по электронной почте.

## <a name="sample-excel-file"></a>Пример файла Excel

[ Скачайтеdaily-readings.xlsx](daily-readings.xlsx) для готовой к использованию книги. Добавьте следующий скрипт, чтобы попробовать пример самостоятельно!

## <a name="sample-code-record-and-report-daily-readings"></a>Пример кода: запись ежедневных чтений и отчетов

```TypeScript
function main(workbook: ExcelScript.Workbook, newData: string): string {
  // Get the table by its name.
  const table = workbook.getTable("ReadingTable");

  // Read the current last entry in the Reading column.
  const readingColumn = table.getColumnByName("Reading");
  const readingColumnValues = readingColumn.getRange().getValues();
  const previousValue = readingColumnValues[readingColumnValues.length - 1][0] as number;

  // Add a row with the date, new value, and a formula calculating the difference.
  const currentDate = new Date(Date.now()).toLocaleDateString();
  const newRow = [currentDate, newData, "=[@Reading]-OFFSET([@Reading],-1,0)"];
  table.addRow(-1, newRow,);

  // Return the difference between the newData and the previous entry.
  const difference = Number.parseFloat(newData) - previousValue;
  console.log(difference);
  return difference;
}
```

## <a name="sample-flow-report-day-to-day-changes"></a>Пример потока: ежедневное создание отчетов об изменениях

Выполните следующие действия, чтобы создать [поток Power Automate](https://powerautomate.microsoft.com/) для примера.

1. Создайте новый **запланированный облачный поток**.
1. Запланируйте повторение потока каждые **1 день**.

    :::image type="content" source="../../images/day-to-day-changes-flow-1.png" alt-text="Шаг создания потока, показывающий, что он будет повторяться каждый день.":::
1. Нажмите **Создать**.
1. В реальном потоке вы добавите шаг, который получает данные. Данные могут поступать из другой книги, адаптивной карточки Teams или любого другого источника. Чтобы протестировать пример, сделайте тестовый номер. Добавьте новый шаг с действием **инициализации переменной** . Присвойте ему следующие значения.
    1. **Имя**: входные данные
    1. **Тип**: Целое число
    1. **Значение**: 190000

    :::image type="content" source="../../images/day-to-day-changes-flow-2.png" alt-text="Действие инициализации переменной с заданными значениями.":::
1. Добавьте новый шаг с помощью **соединителя Excel Online (business)** с действием **запуска скрипта** . Используйте следующие значения для действия.
    1. **Расположение**: OneDrive для бизнеса
    1. **Библиотека документов**: OneDrive
    1. **Файл**: daily-readings.xlsx *(выбирается через браузер файлов)*
    1. **Сценарий**: имя скрипта
    1. **newData**: *входные данные (динамическое содержимое)*

    :::image type="content" source="../../images/day-to-day-changes-flow-3.png" alt-text="Действие запуска скрипта с заданными значениями.":::
1. Скрипт возвращает разницу в ежедневном чтении в виде динамического содержимого с именем "result". В этом примере вы можете отправить эти сведения себе по электронной почте. Создайте новый шаг, использующий соединитель **Outlook** с действием **отправки сообщения электронной почты (версии 2)** (или любым другим клиентом электронной почты, который вы предпочитаете). Для выполнения действия используйте следующие значения.
    1. **To**: Ваш адрес электронной почты
    1. **Тема**: изменение ежедневного чтения
    1. **Текст**: результат "Отличие от вчера" *(динамическое содержимое из Excel)*

    :::image type="content" source="../../images/day-to-day-changes-flow-4.png" alt-text="Завершенный соединитель Outlook в Power Automate.":::
1. Сохраните поток и попробуйте его. Нажмите **кнопку "** Тест" на странице редактора потоков. Не забудьте разрешить доступ при появлении запроса.
