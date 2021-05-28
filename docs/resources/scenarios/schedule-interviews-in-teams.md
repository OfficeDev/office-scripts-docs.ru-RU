---
title: Расписание интервью в Teams
description: Узнайте, как использовать Office скрипты для отправки собрания Teams из Excel данных.
ms.date: 05/25/2021
localization_priority: Normal
ms.openlocfilehash: f93d9ceca6603ddb9e7123a393787fcf54597cca
ms.sourcegitcommit: 339ecbb9914d54f919e3475018888fb5d00abe89
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 05/28/2021
ms.locfileid: "52697786"
---
# <a name="office-scripts-sample-scenario-schedule-interviews-in-teams"></a><span data-ttu-id="90113-103">Office Пример сценария: Расписание интервью в Teams</span><span class="sxs-lookup"><span data-stu-id="90113-103">Office Scripts sample scenario: Schedule interviews in Teams</span></span>

<span data-ttu-id="90113-104">В этом сценарии вы будете вербовщиком кадров, запланируя встречи с кандидатами в Teams.</span><span class="sxs-lookup"><span data-stu-id="90113-104">In this scenario, you're an HR recruiter scheduling interview meetings with candidates in Teams.</span></span> <span data-ttu-id="90113-105">Вы управляете расписанием собеседований кандидатов в Excel файле.</span><span class="sxs-lookup"><span data-stu-id="90113-105">You manage the interview schedule of candidates in an Excel file.</span></span> <span data-ttu-id="90113-106">Необходимо отправить приглашение на Teams как кандидату, так и интервьюеру.</span><span class="sxs-lookup"><span data-stu-id="90113-106">You'll need to send the Teams meeting invite to both the candidate and interviewers.</span></span> <span data-ttu-id="90113-107">Затем необходимо обновить файл Excel с подтверждением того, что Teams были отправлены собрания.</span><span class="sxs-lookup"><span data-stu-id="90113-107">You then need to update the Excel file with the confirmation that Teams meetings have been sent.</span></span>

<span data-ttu-id="90113-108">Решение состоит из трех этапов, объединенных в один Power Automate потока.</span><span class="sxs-lookup"><span data-stu-id="90113-108">The solution has three steps that are combined in a single Power Automate flow.</span></span>

1. <span data-ttu-id="90113-109">Скрипт извлекает данные из таблицы и возвращает массив объектов в качестве данных JSON.</span><span class="sxs-lookup"><span data-stu-id="90113-109">A script extracts data from a table and returns an array of objects as JSON data.</span></span>
1. <span data-ttu-id="90113-110">Затем данные отправляются в Teams **создать Teams** собрания для отправки приглашений.</span><span class="sxs-lookup"><span data-stu-id="90113-110">The data is then sent to the Teams **Create a Teams meeting** action to send invites.</span></span>
1. <span data-ttu-id="90113-111">Эти же данные JSON отправляются в другой скрипт, чтобы обновить состояние приглашения.</span><span class="sxs-lookup"><span data-stu-id="90113-111">The same JSON data is sent to another script to update the status of the invitation.</span></span>

## <a name="scripting-skills-covered"></a><span data-ttu-id="90113-112">Навыки скриптов, охватываемых</span><span class="sxs-lookup"><span data-stu-id="90113-112">Scripting skills covered</span></span>

* <span data-ttu-id="90113-113">Power Automate потоков</span><span class="sxs-lookup"><span data-stu-id="90113-113">Power Automate flows</span></span>
* <span data-ttu-id="90113-114">Teams интеграции</span><span class="sxs-lookup"><span data-stu-id="90113-114">Teams integration</span></span>
* <span data-ttu-id="90113-115">Размыв таблиц</span><span class="sxs-lookup"><span data-stu-id="90113-115">Table parsing</span></span>

## <a name="sample-excel-file"></a><span data-ttu-id="90113-116">Пример Excel файла</span><span class="sxs-lookup"><span data-stu-id="90113-116">Sample Excel file</span></span>

<span data-ttu-id="90113-117">Скачайте файл <a href="hr-schedule.xlsx">hr-schedule.xlsx, </a> используемый в этом решении, и попробуйте его самостоятельно!</span><span class="sxs-lookup"><span data-stu-id="90113-117">Download the file <a href="hr-schedule.xlsx">hr-schedule.xlsx</a> used in this solution and try it out yourself!</span></span> <span data-ttu-id="90113-118">Обязательно измените хотя бы один из адресов электронной почты, чтобы получить приглашение.</span><span class="sxs-lookup"><span data-stu-id="90113-118">Be sure to change at least one of the email addresses so that you receive an invite.</span></span>

## <a name="sample-code-extract-table-data-to-schedule-invites"></a><span data-ttu-id="90113-119">Пример кода. Извлечение данных таблицы для расписания приглашений</span><span class="sxs-lookup"><span data-stu-id="90113-119">Sample code: Extract table data to schedule invites</span></span>

<span data-ttu-id="90113-120">Назови этот **сценарий Schedule Interviews** для потока.</span><span class="sxs-lookup"><span data-stu-id="90113-120">Name this script **Schedule Interviews** for the flow.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook): InterviewInvite[] {
  const MEETING_DURATION = workbook.getWorksheet("Constants").getRange("B1").getValue() as number;
  const MESSAGE_TEMPLATE = workbook.getWorksheet("Constants").getRange("B2").getValue() as string;

  // Get the interview candidate information.
  const sheet = workbook.getWorksheet("Interviews");
  const table = sheet.getTables()[0];
  const dataRows = table.getRangeBetweenHeaderAndTotal().getValues();

  // Convert the table rows into InterviewInvite objects for the flow.
  let invites: InterviewInvite[] = [];
  dataRows.forEach((row) => {
    const inviteSent = row[1] as boolean;
    if (!inviteSent) {
      const startTime = new Date(Math.round(((row[6] as number) - 25569) * 86400 * 1000));
      const finishTime = new Date(startTime.getTime() + MEETING_DURATION * 60 * 1000);
      const candidateName = row[2] as string;
      const interviewerName = row[4] as string;

      invites.push({
        ID: row[0] as string,
        Candidate: candidateName,
        CandidateEmail: row[3] as string,
        Interviewer: row[4] as string,
        InterviewerEmail: row[5] as string,
        StartTime: startTime.toISOString(),
        FinishTime: finishTime.toISOString(),
        Message: generateInviteMessage(MESSAGE_TEMPLATE, candidateName, interviewerName)
      });
    }    
  });

  console.log(JSON.stringify(invites));
  return invites;
}

function generateInviteMessage(
  messageTemplate: string,
   candidate: string,
   interviewer: string) : string {
  return messageTemplate.replace("_Candidate_", candidate).replace("_Interviewer_", interviewer);
}

// The interview invite information.
interface InterviewInvite {
  ID: string
  Candidate: string
  CandidateEmail: string
  Interviewer: string
  InterviewerEmail: string
  StartTime: string
  FinishTime: string
  Message: string
}
```

## <a name="sample-code-mark-rows-as-invited"></a><span data-ttu-id="90113-121">Пример кода: пометить строки как приглашенные</span><span class="sxs-lookup"><span data-stu-id="90113-121">Sample code: Mark rows as invited</span></span>

<span data-ttu-id="90113-122">Назови этот **скрипт Запись отправленных приглашений** для потока.</span><span class="sxs-lookup"><span data-stu-id="90113-122">Name this script **Record Sent Invites** for the flow.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook, invites: InterviewInvite[]) {
  const table = workbook.getWorksheet("Interviews").getTables()[0];

  // Get the ID and Invite Sent columns from the table.
  const idColumn = table.getColumnByName("ID");
  const idRange = idColumn.getRangeBetweenHeaderAndTotal().getValues();
  const inviteSentColumn = table.getColumnByName("Invite Sent?");

  const dataRowCount = idRange.length;

  // Find matching IDs to mark the correct row.
  for (let row = 0; row < dataRowCount; row++){
    let inviteSent = invites.find((invite) => {
      return invite.ID == idRange[row][0] as string;
    });

    if (inviteSent) {
      inviteSentColumn.getRangeBetweenHeaderAndTotal().getCell(row, 0).setValue(true);
      console.log(`Invite for ${inviteSent.Candidate} has been sent.`);
    }
  } 
}

// The interview invite information.
interface InterviewInvite {
  ID: string
  Candidate: string
  CandidateEmail: string
  Interviewer: string
  InterviewerEmail: string
  StartTime: string
  FinishTime: string
  Message: string
}
```

## <a name="sample-flow-run-the-interview-scheduling-scripts-and-send-the-teams-meetings"></a><span data-ttu-id="90113-123">Пример потока: запустите сценарии планирования интервью и отправьте Teams собрания</span><span class="sxs-lookup"><span data-stu-id="90113-123">Sample flow: Run the interview scheduling scripts and send the Teams meetings</span></span>

1. <span data-ttu-id="90113-124">Создайте новый **поток мгновенных облаков.**</span><span class="sxs-lookup"><span data-stu-id="90113-124">Create a new **Instant cloud flow**.</span></span>
1. <span data-ttu-id="90113-125">Выберите **вручную вызвать поток и** нажмите **кнопку Создать**.</span><span class="sxs-lookup"><span data-stu-id="90113-125">Select **Manually trigger a flow** and press **Create**.</span></span>
1. <span data-ttu-id="90113-126">Добавьте новый **шаг,** использующий **соединителю Excel Online (Бизнес)** и действие **скрипта Run.**</span><span class="sxs-lookup"><span data-stu-id="90113-126">Add a **New step** that uses the **Excel Online (Business)** connector and the **Run script** action.</span></span> <span data-ttu-id="90113-127">Заполнять соединитектор следующими значениями.</span><span class="sxs-lookup"><span data-stu-id="90113-127">Complete the connector with the following values.</span></span>
    1. <span data-ttu-id="90113-128">**Расположение**: OneDrive для бизнеса</span><span class="sxs-lookup"><span data-stu-id="90113-128">**Location**: OneDrive for Business</span></span>
    1. <span data-ttu-id="90113-129">**Библиотека документов**: OneDrive</span><span class="sxs-lookup"><span data-stu-id="90113-129">**Document Library**: OneDrive</span></span>
    1. <span data-ttu-id="90113-130">**Файл**: hr-interviews.xlsx *(выбранный через браузер файлов)*</span><span class="sxs-lookup"><span data-stu-id="90113-130">**File**: hr-interviews.xlsx *(Chosen through the file browser)*</span></span>
    1. **Сценарий.** Расписание интервью Снимок экрана завершенного :::image type="content" source="../../images/schedule-interviews-1.png" alt-text="соединиттеля Excel Online (Бизнес)"::: для получения данных интервью из книги в Power Automate
1. <span data-ttu-id="90113-132">Добавьте новый **шаг,** использующий **действие Create a Teams собрания.**</span><span class="sxs-lookup"><span data-stu-id="90113-132">Add a **New step** that uses the **Create a Teams meeting** action.</span></span> <span data-ttu-id="90113-133">При выборе динамического контента Excel соединители, для каждого блока создается применение к каждому блоку. </span><span class="sxs-lookup"><span data-stu-id="90113-133">As you select dynamic content from the Excel connector, an **Apply to each** block will be generated for your flow.</span></span> <span data-ttu-id="90113-134">Заполнять соединитектор следующими значениями.</span><span class="sxs-lookup"><span data-stu-id="90113-134">Complete the connector with the following values.</span></span>
    1. <span data-ttu-id="90113-135">**Calendar id**: Calendar</span><span class="sxs-lookup"><span data-stu-id="90113-135">**Calendar id**: Calendar</span></span>
    1. <span data-ttu-id="90113-136">**Тема:** Интервью Contoso</span><span class="sxs-lookup"><span data-stu-id="90113-136">**Subject**: Contoso Interview</span></span>
    1. <span data-ttu-id="90113-137">**Сообщение.** **Сообщение** (Excel значение)</span><span class="sxs-lookup"><span data-stu-id="90113-137">**Message**: **Message** (the Excel value)</span></span>
    1. <span data-ttu-id="90113-138">**Часовой пояс**: Тихоокеанское стандартное время</span><span class="sxs-lookup"><span data-stu-id="90113-138">**Time zone**: Pacific Standard Time</span></span>
    1. <span data-ttu-id="90113-139">**Время начала:** **StartTime** (Excel значение)</span><span class="sxs-lookup"><span data-stu-id="90113-139">**Start time**: **StartTime** (the Excel value)</span></span>
    1. <span data-ttu-id="90113-140">**End time:** **FinishTime** (Excel значение)</span><span class="sxs-lookup"><span data-stu-id="90113-140">**End time**: **FinishTime** (the Excel value)</span></span>
    1. **Необходимые участники:** **CandidateEmail;** **InterviewerEmail** (Excel значений) Снимок экрана завершенного соединиттеля Teams для расписания :::image type="content" source="../../images/schedule-interviews-2.png" alt-text="собраний в Power Automate":::
1. <span data-ttu-id="90113-142">В том же **применить к каждому блоку** добавить **еще один соединителю Excel Online (Бизнес)** с действием **сценария Run.**</span><span class="sxs-lookup"><span data-stu-id="90113-142">In the same **Apply to each** block, add another **Excel Online (Business)** connector with the **Run script** action.</span></span> <span data-ttu-id="90113-143">Используйте следующие значения.</span><span class="sxs-lookup"><span data-stu-id="90113-143">Use the following values.</span></span>
    1. <span data-ttu-id="90113-144">**Расположение**: OneDrive для бизнеса</span><span class="sxs-lookup"><span data-stu-id="90113-144">**Location**: OneDrive for Business</span></span>
    1. <span data-ttu-id="90113-145">**Библиотека документов**: OneDrive</span><span class="sxs-lookup"><span data-stu-id="90113-145">**Document Library**: OneDrive</span></span>
    1. <span data-ttu-id="90113-146">**Файл**: hr-interviews.xlsx *(выбранный через браузер файлов)*</span><span class="sxs-lookup"><span data-stu-id="90113-146">**File**: hr-interviews.xlsx *(Chosen through the file browser)*</span></span>
    1. <span data-ttu-id="90113-147">**Сценарий:** Запись отправленных приглашений</span><span class="sxs-lookup"><span data-stu-id="90113-147">**Script**: Record Sent Invites</span></span>
    1. **приглашает:** результат **(Excel** значение) Снимок экрана завершенного соединиттеля Excel :::image type="content" source="../../images/schedule-interviews-3.png" alt-text="Online (Бизнес)"::: для записи того, что приглашения были отправлены в Power Automate
1. <span data-ttu-id="90113-149">Сохраните поток и попробуйте его.</span><span class="sxs-lookup"><span data-stu-id="90113-149">Save the flow and try it out.</span></span>

## <a name="training-video-send-a-teams-meeting-from-excel-data"></a><span data-ttu-id="90113-150">Обучающее видео: отправка Teams собрания из Excel данных</span><span class="sxs-lookup"><span data-stu-id="90113-150">Training video: Send a Teams meeting from Excel data</span></span>

<span data-ttu-id="90113-151">[Смотреть Sudhi Ramamurthy ходить через версию этого примера на YouTube](https://youtu.be/HyBdx52NOE8).</span><span class="sxs-lookup"><span data-stu-id="90113-151">[Watch Sudhi Ramamurthy walk through a version of this sample on YouTube](https://youtu.be/HyBdx52NOE8).</span></span> <span data-ttu-id="90113-152">В его версии используется более надежный скрипт, который обрабатывает изменение столбцов и устаревшее время собраний.</span><span class="sxs-lookup"><span data-stu-id="90113-152">His version uses a more robust script that handles changing columns and obsolete meeting times.</span></span>
