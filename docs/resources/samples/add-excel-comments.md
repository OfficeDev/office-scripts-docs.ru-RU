---
title: Добавление комментариев в Excel
description: Узнайте, как использовать скрипты Office для добавления комментариев в таблицу.
ms.date: 03/29/2021
localization_priority: Normal
ms.openlocfilehash: aaaf26df6973bd081290b0fbb67edecad8627e53
ms.sourcegitcommit: 5d24e77df70aa2c1c982275d53213c2a9323ff86
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 04/02/2021
ms.locfileid: "51571530"
---
# <a name="add-comments-in-excel"></a><span data-ttu-id="1a09e-103">Добавление комментариев в Excel</span><span class="sxs-lookup"><span data-stu-id="1a09e-103">Add comments in Excel</span></span>

<span data-ttu-id="1a09e-104">В этом примере показано, как добавлять комментарии в ячейку, [включая](https://support.microsoft.com/office/90701709-5dc1-41c7-aa48-b01d4a46e8c7) @mentioning коллегу.</span><span class="sxs-lookup"><span data-stu-id="1a09e-104">This sample shows how to add comments to a cell including [@mentioning](https://support.microsoft.com/office/90701709-5dc1-41c7-aa48-b01d4a46e8c7) a colleague.</span></span>

## <a name="example-scenario"></a><span data-ttu-id="1a09e-105">Пример сценария</span><span class="sxs-lookup"><span data-stu-id="1a09e-105">Example scenario</span></span>

* <span data-ttu-id="1a09e-106">Руководство группы поддерживает расписание смены.</span><span class="sxs-lookup"><span data-stu-id="1a09e-106">The team lead maintains the shift schedule.</span></span> <span data-ttu-id="1a09e-107">Руководство группы назначает ID сотрудника для записи смены.</span><span class="sxs-lookup"><span data-stu-id="1a09e-107">The team lead assigns an employee ID to the shift record.</span></span>
* <span data-ttu-id="1a09e-108">Руководство группы хочет уведомить сотрудника.</span><span class="sxs-lookup"><span data-stu-id="1a09e-108">The team lead wishes to notify the employee.</span></span> <span data-ttu-id="1a09e-109">Добавляя комментарий, @mentions сотрудник, сотрудник получает электронное сообщение из таблицы.</span><span class="sxs-lookup"><span data-stu-id="1a09e-109">By adding a comment that @mentions the employee, the employee is emailed with a custom message from the worksheet.</span></span>
* <span data-ttu-id="1a09e-110">Впоследствии сотрудник может просматривать книгу и отвечать на комментарий в удобное для него время.</span><span class="sxs-lookup"><span data-stu-id="1a09e-110">Subsequently, the employee can view the workbook and respond to the comment at their convenience.</span></span>

## <a name="solution"></a><span data-ttu-id="1a09e-111">Решение</span><span class="sxs-lookup"><span data-stu-id="1a09e-111">Solution</span></span>

1. <span data-ttu-id="1a09e-112">Сценарий извлекает сведения о сотрудниках из таблицы сотрудников.</span><span class="sxs-lookup"><span data-stu-id="1a09e-112">The script extracts employee information from the employee worksheet.</span></span>
1. <span data-ttu-id="1a09e-113">Затем скрипт добавляет комментарий (включая соответствующую электронную почту сотрудника) в соответствующую ячейку в записи смены.</span><span class="sxs-lookup"><span data-stu-id="1a09e-113">The script then adds a comment (including the relevant employee email) to the appropriate cell in the shift record.</span></span>
1. <span data-ttu-id="1a09e-114">Существующие комментарии в ячейке удаляются перед добавлением нового комментария.</span><span class="sxs-lookup"><span data-stu-id="1a09e-114">Existing comments in the cell are removed before adding the new comment.</span></span>

## <a name="sample-code-add-comments"></a><span data-ttu-id="1a09e-115">Пример кода: добавление комментариев</span><span class="sxs-lookup"><span data-stu-id="1a09e-115">Sample code: Add comments</span></span>

<span data-ttu-id="1a09e-116">Скачайте файл <a href="excel-comments.xlsx">excel-comments.xlsx, </a> используемый в этом примере, и попробуйте его самостоятельно!</span><span class="sxs-lookup"><span data-stu-id="1a09e-116">Download the file <a href="excel-comments.xlsx">excel-comments.xlsx</a> used in this sample and try it out yourself!</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
    const employees = workbook.getWorksheet('Employees').getUsedRange().getTexts();
    console.log(employees); 

    const scheduleSheet = workbook.getWorksheet('Schedule');
    const table = scheduleSheet.getTables()[0];
    const range = table.getRangeBetweenHeaderAndTotal();
    const scheduleData = range.getTexts();

    for (let i=0; i < scheduleData.length; i++) {
      let eId = scheduleData[i][3];

      let employeeInfo = employees.find(e => e[0] === eId);
      if (employeeInfo) {
        console.log("Found a match " + employeeInfo);
        let adminNotes = scheduleData[i][4];
        try { 
          let comment = workbook.getCommentByCell(range.getCell(i, 5));
          comment.delete();
        } catch {
            console.log("Ignore if there is no existing comment in the cell");
        }
        workbook.addComment(range.getCell(i,5), {
          mentions: [{
            email: employeeInfo[1],
            id: 0,
            name: employeeInfo[2]
          }],
          richContent: `<at id=\"0\">${employeeInfo[2]}</at> ${adminNotes}`
        }, ExcelScript.ContentType.mention);        
        
      } else {
        console.log("No match for: " + eId);
      }
    }
    return;
}
```

## <a name="training-video-add-comments"></a><span data-ttu-id="1a09e-117">Обучающее видео: добавление комментариев</span><span class="sxs-lookup"><span data-stu-id="1a09e-117">Training video: Add comments</span></span>

<span data-ttu-id="1a09e-118">[![Просмотр пошагового видео о добавлении комментариев в файл Excel](../../images/comments-vid.jpg)](https://youtu.be/CpR78nkaOFw "Пошаговая видеозапись добавления комментариев в файл Excel")</span><span class="sxs-lookup"><span data-stu-id="1a09e-118">[![Watch step-by-step video on how to add comments in an Excel file](../../images/comments-vid.jpg)](https://youtu.be/CpR78nkaOFw "Step-by-step video on how to add comments in an Excel file")</span></span>
