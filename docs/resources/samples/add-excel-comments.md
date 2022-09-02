---
title: Добавление примечаний в Excel
description: Узнайте, как использовать сценарии Office для добавления комментариев на лист.
ms.date: 06/29/2021
ms.localizationpriority: medium
ms.openlocfilehash: 90f072805e6798a4f9d6e74889ccca15610c87bd
ms.sourcegitcommit: a6504f8b0d6b717457c6e0b5306c35ad3900914e
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/02/2022
ms.locfileid: "67572495"
---
# <a name="add-comments-in-excel"></a>Добавление примечаний в Excel

В этом примере показано, как добавить комментарии в ячейку, [@mentioning](https://support.microsoft.com/office/90701709-5dc1-41c7-aa48-b01d4a46e8c7) коллегу.

## <a name="example-scenario"></a>Пример сценария

* Руководителя группы поддерживает расписание смены. Ведущий сотрудник назначает идентификатор сотрудника записи смены.
* Руководство команды хочет уведомить сотрудника. Добавляя комментарий, @mentions сотруднику, сотруднику отправляется сообщение электронной почты с листа.
* Впоследствии сотрудник может просматривать книгу и отвечать на комментарий при необходимости.

## <a name="solution"></a>Решение

1. Сценарий извлекает сведения о сотрудниках из листа сотрудника.
1. Затем сценарий добавляет комментарий (включая соответствующее сообщение электронной почты сотрудника) в соответствующую ячейку в записи смены.
1. Существующие примечания в ячейке удаляются перед добавлением нового комментария.

## <a name="sample-excel-file"></a>Пример файла Excel

[ Скачайтеexcel-comments.xlsx](excel-comments.xlsx) для готовой к использованию книги. Добавьте следующий скрипт, чтобы попробовать пример самостоятельно!

## <a name="sample-code-add-comments"></a>Пример кода: добавление примечаний

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

## <a name="training-video-add-comments"></a>Обучающее видео: добавление примечаний

[Просмотрите этот пример на YouTube](https://youtu.be/CpR78nkaOFw), чтобы просмотреть этот пример.
