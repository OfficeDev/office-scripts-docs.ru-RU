---
title: Добавление комментариев в Excel
description: Узнайте, как использовать Office скрипты для добавления комментариев в таблицу.
ms.date: 05/03/2021
localization_priority: Normal
ms.openlocfilehash: e5e5d17c076eceaf06fddeea1a67d31ee3581f31
ms.sourcegitcommit: 763d341857bcb209b2f2c278a82fdb63d0e18f0a
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 05/08/2021
ms.locfileid: "52285936"
---
# <a name="add-comments-in-excel"></a><span data-ttu-id="52420-103">Добавление комментариев в Excel</span><span class="sxs-lookup"><span data-stu-id="52420-103">Add comments in Excel</span></span>

<span data-ttu-id="52420-104">В этом примере показано, как добавлять комментарии в ячейку, [включая](https://support.microsoft.com/office/90701709-5dc1-41c7-aa48-b01d4a46e8c7) @mentioning коллегу.</span><span class="sxs-lookup"><span data-stu-id="52420-104">This sample shows how to add comments to a cell including [@mentioning](https://support.microsoft.com/office/90701709-5dc1-41c7-aa48-b01d4a46e8c7) a colleague.</span></span>

## <a name="example-scenario"></a><span data-ttu-id="52420-105">Пример сценария</span><span class="sxs-lookup"><span data-stu-id="52420-105">Example scenario</span></span>

* <span data-ttu-id="52420-106">Руководство группы поддерживает расписание смены.</span><span class="sxs-lookup"><span data-stu-id="52420-106">The team lead maintains the shift schedule.</span></span> <span data-ttu-id="52420-107">Руководство группы назначает ID сотрудника для записи смены.</span><span class="sxs-lookup"><span data-stu-id="52420-107">The team lead assigns an employee ID to the shift record.</span></span>
* <span data-ttu-id="52420-108">Руководство группы хочет уведомить сотрудника.</span><span class="sxs-lookup"><span data-stu-id="52420-108">The team lead wishes to notify the employee.</span></span> <span data-ttu-id="52420-109">Добавляя комментарий, @mentions сотрудник, сотрудник получает электронное сообщение из таблицы.</span><span class="sxs-lookup"><span data-stu-id="52420-109">By adding a comment that @mentions the employee, the employee is emailed with a custom message from the worksheet.</span></span>
* <span data-ttu-id="52420-110">Впоследствии сотрудник может просматривать книгу и отвечать на комментарий в удобное для него время.</span><span class="sxs-lookup"><span data-stu-id="52420-110">Subsequently, the employee can view the workbook and respond to the comment at their convenience.</span></span>

## <a name="solution"></a><span data-ttu-id="52420-111">Решение</span><span class="sxs-lookup"><span data-stu-id="52420-111">Solution</span></span>

1. <span data-ttu-id="52420-112">Сценарий извлекает сведения о сотрудниках из таблицы сотрудников.</span><span class="sxs-lookup"><span data-stu-id="52420-112">The script extracts employee information from the employee worksheet.</span></span>
1. <span data-ttu-id="52420-113">Затем скрипт добавляет комментарий (включая соответствующую электронную почту сотрудника) в соответствующую ячейку в записи смены.</span><span class="sxs-lookup"><span data-stu-id="52420-113">The script then adds a comment (including the relevant employee email) to the appropriate cell in the shift record.</span></span>
1. <span data-ttu-id="52420-114">Существующие комментарии в ячейке удаляются перед добавлением нового комментария.</span><span class="sxs-lookup"><span data-stu-id="52420-114">Existing comments in the cell are removed before adding the new comment.</span></span>

## <a name="sample-code-add-comments"></a><span data-ttu-id="52420-115">Пример кода: добавление комментариев</span><span class="sxs-lookup"><span data-stu-id="52420-115">Sample code: Add comments</span></span>

<span data-ttu-id="52420-116">Скачайте файл <a href="excel-comments.xlsx">excel-comments.xlsx, </a> используемый в этом примере, и попробуйте его самостоятельно!</span><span class="sxs-lookup"><span data-stu-id="52420-116">Download the file <a href="excel-comments.xlsx">excel-comments.xlsx</a> used in this sample and try it out yourself!</span></span>

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

## <a name="training-video-add-comments"></a><span data-ttu-id="52420-117">Обучающее видео: добавление комментариев</span><span class="sxs-lookup"><span data-stu-id="52420-117">Training video: Add comments</span></span>

<span data-ttu-id="52420-118">[Смотреть Sudhi Ramamurthy ходить через этот пример на YouTube](https://youtu.be/CpR78nkaOFw).</span><span class="sxs-lookup"><span data-stu-id="52420-118">[Watch Sudhi Ramamurthy walk through this sample on YouTube](https://youtu.be/CpR78nkaOFw).</span></span>
