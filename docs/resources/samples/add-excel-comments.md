---
title: Добавление комментариев в Excel
description: Узнайте, как использовать Office скрипты для добавления комментариев в таблицу.
ms.date: 06/29/2021
localization_priority: Normal
ms.openlocfilehash: 77e308d020281c71751e2652f8dbaec00c263e44
ms.sourcegitcommit: 211c157ca746e266eeb079f5fa1925a1e35ab702
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/07/2021
ms.locfileid: "53313913"
---
# <a name="add-comments-in-excel"></a>Добавление комментариев в Excel

В этом примере показано, как добавлять комментарии в ячейку, [включая](https://support.microsoft.com/office/90701709-5dc1-41c7-aa48-b01d4a46e8c7) @mentioning коллегу.

## <a name="example-scenario"></a>Пример сценария

* Руководство группы поддерживает расписание смены. Руководство группы назначает ID сотрудника для записи смены.
* Руководство группы хочет уведомить сотрудника. Добавляя комментарий, @mentions сотрудник, сотрудник получает электронное сообщение из таблицы.
* Впоследствии сотрудник может просматривать книгу и отвечать на комментарий в удобное для него время.

## <a name="solution"></a>Решение

1. Сценарий извлекает сведения о сотрудниках из таблицы сотрудников.
1. Затем скрипт добавляет комментарий (включая соответствующую электронную почту сотрудника) в соответствующую ячейку в записи смены.
1. Существующие комментарии в ячейке удаляются перед добавлением нового комментария.

## <a name="sample-excel-file"></a>Пример Excel файла

Скачайте <a href="excel-comments.xlsx">excel-comments.xlsx</a> для готовой к использованию книги. Добавьте следующий скрипт, чтобы попробовать пример самостоятельно!

## <a name="sample-code-add-comments"></a>Пример кода: добавление комментариев

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Get the list of employees.
  const employees = workbook.getWorksheet('Employees').getUsedRange().getTexts();
  console.log(employees); 
  
  // Get the schedule information from the schedule table.
  const scheduleSheet = workbook.getWorksheet('Schedule');
  const table = scheduleSheet.getTables()[0];
  const range = table.getRangeBetweenHeaderAndTotal();
  const scheduleData = range.getTexts();

  // Look through the schedule for a matching employee.
  for (let i = 0; i < scheduleData.length; i++) {
    let employeeId = scheduleData[i][3];

    // Compare the employee ID in the schedule against the employee information table.
    let employeeInfo = employees.find(employeeRow => employeeRow[0] === employeeId);
    if (employeeInfo) {
      console.log("Found a match " + employeeInfo);
      let adminNotes = scheduleData[i][4];

      // Look for and delete old comments, so we avoid conflicts.
      let comment = workbook.getCommentByCell(range.getCell(i, 5));
      if (comment) {
        comment.delete();
      }

      // Add a comment using the admin notes as the text.
      workbook.addComment(range.getCell(i,5), {
        mentions: [{
          email: employeeInfo[1],
          id: 0, // This ID maps this mention to the `id=0` text in the comment.
          name: employeeInfo[2]
        }],
        richContent: `<at id=\"0\">${employeeInfo[2]}</at> ${adminNotes}`
      }, ExcelScript.ContentType.mention);        
      
    } else {
      console.log("No match for: " + employeeId);
    }
  }
}
```

## <a name="training-video-add-comments"></a>Обучающее видео: добавление комментариев

[Смотреть Sudhi Ramamurthy ходить через этот пример на YouTube](https://youtu.be/CpR78nkaOFw).
