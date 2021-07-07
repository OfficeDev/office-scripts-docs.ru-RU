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
# <a name="add-comments-in-excel"></a><span data-ttu-id="51159-103">Добавление комментариев в Excel</span><span class="sxs-lookup"><span data-stu-id="51159-103">Add comments in Excel</span></span>

<span data-ttu-id="51159-104">В этом примере показано, как добавлять комментарии в ячейку, [включая](https://support.microsoft.com/office/90701709-5dc1-41c7-aa48-b01d4a46e8c7) @mentioning коллегу.</span><span class="sxs-lookup"><span data-stu-id="51159-104">This sample shows how to add comments to a cell including [@mentioning](https://support.microsoft.com/office/90701709-5dc1-41c7-aa48-b01d4a46e8c7) a colleague.</span></span>

## <a name="example-scenario"></a><span data-ttu-id="51159-105">Пример сценария</span><span class="sxs-lookup"><span data-stu-id="51159-105">Example scenario</span></span>

* <span data-ttu-id="51159-106">Руководство группы поддерживает расписание смены.</span><span class="sxs-lookup"><span data-stu-id="51159-106">The team lead maintains the shift schedule.</span></span> <span data-ttu-id="51159-107">Руководство группы назначает ID сотрудника для записи смены.</span><span class="sxs-lookup"><span data-stu-id="51159-107">The team lead assigns an employee ID to the shift record.</span></span>
* <span data-ttu-id="51159-108">Руководство группы хочет уведомить сотрудника.</span><span class="sxs-lookup"><span data-stu-id="51159-108">The team lead wishes to notify the employee.</span></span> <span data-ttu-id="51159-109">Добавляя комментарий, @mentions сотрудник, сотрудник получает электронное сообщение из таблицы.</span><span class="sxs-lookup"><span data-stu-id="51159-109">By adding a comment that @mentions the employee, the employee is emailed with a custom message from the worksheet.</span></span>
* <span data-ttu-id="51159-110">Впоследствии сотрудник может просматривать книгу и отвечать на комментарий в удобное для него время.</span><span class="sxs-lookup"><span data-stu-id="51159-110">Subsequently, the employee can view the workbook and respond to the comment at their convenience.</span></span>

## <a name="solution"></a><span data-ttu-id="51159-111">Решение</span><span class="sxs-lookup"><span data-stu-id="51159-111">Solution</span></span>

1. <span data-ttu-id="51159-112">Сценарий извлекает сведения о сотрудниках из таблицы сотрудников.</span><span class="sxs-lookup"><span data-stu-id="51159-112">The script extracts employee information from the employee worksheet.</span></span>
1. <span data-ttu-id="51159-113">Затем скрипт добавляет комментарий (включая соответствующую электронную почту сотрудника) в соответствующую ячейку в записи смены.</span><span class="sxs-lookup"><span data-stu-id="51159-113">The script then adds a comment (including the relevant employee email) to the appropriate cell in the shift record.</span></span>
1. <span data-ttu-id="51159-114">Существующие комментарии в ячейке удаляются перед добавлением нового комментария.</span><span class="sxs-lookup"><span data-stu-id="51159-114">Existing comments in the cell are removed before adding the new comment.</span></span>

## <a name="sample-excel-file"></a><span data-ttu-id="51159-115">Пример Excel файла</span><span class="sxs-lookup"><span data-stu-id="51159-115">Sample Excel file</span></span>

<span data-ttu-id="51159-116">Скачайте <a href="excel-comments.xlsx">excel-comments.xlsx</a> для готовой к использованию книги.</span><span class="sxs-lookup"><span data-stu-id="51159-116">Download <a href="excel-comments.xlsx">excel-comments.xlsx</a> for a ready-to-use workbook.</span></span> <span data-ttu-id="51159-117">Добавьте следующий скрипт, чтобы попробовать пример самостоятельно!</span><span class="sxs-lookup"><span data-stu-id="51159-117">Add the following script to try the sample yourself!</span></span>

## <a name="sample-code-add-comments"></a><span data-ttu-id="51159-118">Пример кода: добавление комментариев</span><span class="sxs-lookup"><span data-stu-id="51159-118">Sample code: Add comments</span></span>

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

## <a name="training-video-add-comments"></a><span data-ttu-id="51159-119">Обучающее видео: добавление комментариев</span><span class="sxs-lookup"><span data-stu-id="51159-119">Training video: Add comments</span></span>

<span data-ttu-id="51159-120">[Смотреть Sudhi Ramamurthy ходить через этот пример на YouTube](https://youtu.be/CpR78nkaOFw).</span><span class="sxs-lookup"><span data-stu-id="51159-120">[Watch Sudhi Ramamurthy walk through this sample on YouTube](https://youtu.be/CpR78nkaOFw).</span></span>
