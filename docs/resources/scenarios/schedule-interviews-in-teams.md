---
title: Планирование собеседований в Teams
description: Узнайте, как использовать Office скрипты для отправки собрания Teams из Excel данных.
ms.date: 06/29/2021
localization_priority: Normal
ms.openlocfilehash: cb24da12637add805d86da4d07ce878509c6a5f6
ms.sourcegitcommit: 211c157ca746e266eeb079f5fa1925a1e35ab702
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/07/2021
ms.locfileid: "53313731"
---
# <a name="office-scripts-sample-scenario-schedule-interviews-in-teams"></a><span data-ttu-id="5fadc-103">Office Пример сценария: Расписание интервью в Teams</span><span class="sxs-lookup"><span data-stu-id="5fadc-103">Office Scripts sample scenario: Schedule interviews in Teams</span></span>

<span data-ttu-id="5fadc-104">В этом сценарии вы будете вербовщиком кадров, запланируя встречи с кандидатами в Teams.</span><span class="sxs-lookup"><span data-stu-id="5fadc-104">In this scenario, you're an HR recruiter scheduling interview meetings with candidates in Teams.</span></span> <span data-ttu-id="5fadc-105">Вы управляете расписанием собеседований кандидатов в Excel файле.</span><span class="sxs-lookup"><span data-stu-id="5fadc-105">You manage the interview schedule of candidates in an Excel file.</span></span> <span data-ttu-id="5fadc-106">Необходимо отправить приглашение на Teams как кандидату, так и интервьюеру.</span><span class="sxs-lookup"><span data-stu-id="5fadc-106">You'll need to send the Teams meeting invite to both the candidate and interviewers.</span></span> <span data-ttu-id="5fadc-107">Затем необходимо обновить файл Excel с подтверждением того, что Teams были отправлены собрания.</span><span class="sxs-lookup"><span data-stu-id="5fadc-107">You then need to update the Excel file with the confirmation that Teams meetings have been sent.</span></span>

<span data-ttu-id="5fadc-108">Решение состоит из трех этапов, объединенных в один Power Automate потока.</span><span class="sxs-lookup"><span data-stu-id="5fadc-108">The solution has three steps that are combined in a single Power Automate flow.</span></span>

1. <span data-ttu-id="5fadc-109">Скрипт извлекает данные из таблицы и возвращает массив объектов в качестве данных JSON.</span><span class="sxs-lookup"><span data-stu-id="5fadc-109">A script extracts data from a table and returns an array of objects as JSON data.</span></span>
1. <span data-ttu-id="5fadc-110">Затем данные отправляются в Teams **создать Teams** собрания для отправки приглашений.</span><span class="sxs-lookup"><span data-stu-id="5fadc-110">The data is then sent to the Teams **Create a Teams meeting** action to send invites.</span></span>
1. <span data-ttu-id="5fadc-111">Эти же данные JSON отправляются в другой скрипт, чтобы обновить состояние приглашения.</span><span class="sxs-lookup"><span data-stu-id="5fadc-111">The same JSON data is sent to another script to update the status of the invitation.</span></span>

## <a name="scripting-skills-covered"></a><span data-ttu-id="5fadc-112">Навыки скриптов, охватываемых</span><span class="sxs-lookup"><span data-stu-id="5fadc-112">Scripting skills covered</span></span>

* <span data-ttu-id="5fadc-113">Power Automate потоков</span><span class="sxs-lookup"><span data-stu-id="5fadc-113">Power Automate flows</span></span>
* <span data-ttu-id="5fadc-114">Teams интеграции</span><span class="sxs-lookup"><span data-stu-id="5fadc-114">Teams integration</span></span>
* <span data-ttu-id="5fadc-115">Размыв таблиц</span><span class="sxs-lookup"><span data-stu-id="5fadc-115">Table parsing</span></span>

## <a name="sample-excel-file"></a><span data-ttu-id="5fadc-116">Пример Excel файла</span><span class="sxs-lookup"><span data-stu-id="5fadc-116">Sample Excel file</span></span>

<span data-ttu-id="5fadc-117">Скачайте файл <a href="hr-schedule.xlsx">hr-schedule.xlsx, </a> используемый в этом решении, и попробуйте его самостоятельно!</span><span class="sxs-lookup"><span data-stu-id="5fadc-117">Download the file <a href="hr-schedule.xlsx">hr-schedule.xlsx</a> used in this solution and try it out yourself!</span></span> <span data-ttu-id="5fadc-118">Обязательно измените хотя бы один из адресов электронной почты, чтобы получить приглашение.</span><span class="sxs-lookup"><span data-stu-id="5fadc-118">Be sure to change at least one of the email addresses so that you receive an invite.</span></span>

## <a name="sample-code-extract-table-data-to-schedule-invites"></a><span data-ttu-id="5fadc-119">Пример кода. Извлечение данных таблицы для расписания приглашений</span><span class="sxs-lookup"><span data-stu-id="5fadc-119">Sample code: Extract table data to schedule invites</span></span>

<span data-ttu-id="5fadc-120">Добавьте этот скрипт в свою коллекцию скриптов.</span><span class="sxs-lookup"><span data-stu-id="5fadc-120">Add this script to your script collection.</span></span> <span data-ttu-id="5fadc-121">Назови **его Schedule Interviews** для потока.</span><span class="sxs-lookup"><span data-stu-id="5fadc-121">Name it **Schedule Interviews** for the flow.</span></span>

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

## <a name="sample-code-mark-rows-as-invited"></a><span data-ttu-id="5fadc-122">Пример кода: пометить строки как приглашенные</span><span class="sxs-lookup"><span data-stu-id="5fadc-122">Sample code: Mark rows as invited</span></span>

<span data-ttu-id="5fadc-123">Добавьте этот скрипт в свою коллекцию скриптов.</span><span class="sxs-lookup"><span data-stu-id="5fadc-123">Add this script to your script collection.</span></span> <span data-ttu-id="5fadc-124">Назови **его Запись отправленных приглашений** для потока.</span><span class="sxs-lookup"><span data-stu-id="5fadc-124">Name it **Record Sent Invites** for the flow.</span></span>

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

## <a name="sample-flow-run-the-interview-scheduling-scripts-and-send-the-teams-meetings"></a><span data-ttu-id="5fadc-125">Пример потока: запустите сценарии планирования интервью и отправьте Teams собрания</span><span class="sxs-lookup"><span data-stu-id="5fadc-125">Sample flow: Run the interview scheduling scripts and send the Teams meetings</span></span>

1. <span data-ttu-id="5fadc-126">Создайте новый **поток мгновенных облаков.**</span><span class="sxs-lookup"><span data-stu-id="5fadc-126">Create a new **Instant cloud flow**.</span></span>
1. <span data-ttu-id="5fadc-127">Выберите **вручную вызвать поток и** выберите **Создать**.</span><span class="sxs-lookup"><span data-stu-id="5fadc-127">Choose **Manually trigger a flow** and select **Create**.</span></span>
1. <span data-ttu-id="5fadc-128">Добавьте новый **шаг,** использующий **соединителю Excel Online (Бизнес)** и действие **скрипта Run.**</span><span class="sxs-lookup"><span data-stu-id="5fadc-128">Add a **New step** that uses the **Excel Online (Business)** connector and the **Run script** action.</span></span> <span data-ttu-id="5fadc-129">Заполнять соединитектор следующими значениями.</span><span class="sxs-lookup"><span data-stu-id="5fadc-129">Complete the connector with the following values.</span></span>
    1. <span data-ttu-id="5fadc-130">**Расположение**: OneDrive для бизнеса</span><span class="sxs-lookup"><span data-stu-id="5fadc-130">**Location**: OneDrive for Business</span></span>
    1. <span data-ttu-id="5fadc-131">**Библиотека документов**: OneDrive</span><span class="sxs-lookup"><span data-stu-id="5fadc-131">**Document Library**: OneDrive</span></span>
    1. <span data-ttu-id="5fadc-132">**Файл**: hr-interviews.xlsx *(выбранный через браузер файлов)*</span><span class="sxs-lookup"><span data-stu-id="5fadc-132">**File**: hr-interviews.xlsx *(Chosen through the file browser)*</span></span>
    1. **Сценарий.** Запланировать скриншот интервью с завершенным :::image type="content" source="../../images/schedule-interviews-1.png" alt-text="соединитетелем Excel Online (Бизнес),"::: чтобы получить данные интервью из книги в Power Automate.
1. <span data-ttu-id="5fadc-134">Добавьте новый **шаг,** использующий **действие Create a Teams собрания.**</span><span class="sxs-lookup"><span data-stu-id="5fadc-134">Add a **New step** that uses the **Create a Teams meeting** action.</span></span> <span data-ttu-id="5fadc-135">При выборе динамического контента Excel соединители, для каждого блока создается применение к каждому блоку. </span><span class="sxs-lookup"><span data-stu-id="5fadc-135">As you select dynamic content from the Excel connector, an **Apply to each** block will be generated for your flow.</span></span> <span data-ttu-id="5fadc-136">Заполнять соединитектор следующими значениями.</span><span class="sxs-lookup"><span data-stu-id="5fadc-136">Complete the connector with the following values.</span></span>
    1. <span data-ttu-id="5fadc-137">**Calendar id**: Calendar</span><span class="sxs-lookup"><span data-stu-id="5fadc-137">**Calendar id**: Calendar</span></span>
    1. <span data-ttu-id="5fadc-138">**Тема:** Интервью Contoso</span><span class="sxs-lookup"><span data-stu-id="5fadc-138">**Subject**: Contoso Interview</span></span>
    1. <span data-ttu-id="5fadc-139">**Сообщение.** **Сообщение** (Excel значение)</span><span class="sxs-lookup"><span data-stu-id="5fadc-139">**Message**: **Message** (the Excel value)</span></span>
    1. <span data-ttu-id="5fadc-140">**Часовой пояс**: Тихоокеанское стандартное время</span><span class="sxs-lookup"><span data-stu-id="5fadc-140">**Time zone**: Pacific Standard Time</span></span>
    1. <span data-ttu-id="5fadc-141">**Время начала:** **StartTime** (Excel значение)</span><span class="sxs-lookup"><span data-stu-id="5fadc-141">**Start time**: **StartTime** (the Excel value)</span></span>
    1. <span data-ttu-id="5fadc-142">**End time:** **FinishTime** (Excel значение)</span><span class="sxs-lookup"><span data-stu-id="5fadc-142">**End time**: **FinishTime** (the Excel value)</span></span>
    1. **Необходимые участники:** **CandidateEmail;** **InterviewerEmail** (Excel значений) Снимок экрана завершенного соединиттеля Teams для расписания :::image type="content" source="../../images/schedule-interviews-2.png" alt-text="собраний в Power Automate.":::
1. <span data-ttu-id="5fadc-144">В том же **применить к каждому блоку** добавить **еще один соединителю Excel Online (Бизнес)** с действием **сценария Run.**</span><span class="sxs-lookup"><span data-stu-id="5fadc-144">In the same **Apply to each** block, add another **Excel Online (Business)** connector with the **Run script** action.</span></span> <span data-ttu-id="5fadc-145">Используйте следующие значения.</span><span class="sxs-lookup"><span data-stu-id="5fadc-145">Use the following values.</span></span>
    1. <span data-ttu-id="5fadc-146">**Расположение**: OneDrive для бизнеса</span><span class="sxs-lookup"><span data-stu-id="5fadc-146">**Location**: OneDrive for Business</span></span>
    1. <span data-ttu-id="5fadc-147">**Библиотека документов**: OneDrive</span><span class="sxs-lookup"><span data-stu-id="5fadc-147">**Document Library**: OneDrive</span></span>
    1. <span data-ttu-id="5fadc-148">**Файл**: hr-interviews.xlsx *(выбранный через браузер файлов)*</span><span class="sxs-lookup"><span data-stu-id="5fadc-148">**File**: hr-interviews.xlsx *(Chosen through the file browser)*</span></span>
    1. <span data-ttu-id="5fadc-149">**Сценарий:** Запись отправленных приглашений</span><span class="sxs-lookup"><span data-stu-id="5fadc-149">**Script**: Record Sent Invites</span></span>
    1. **приглашает:**  результат (Excel) Снимок экрана завершенного соединиттеля :::image type="content" source="../../images/schedule-interviews-3.png" alt-text="Excel Online (Бизнес)"::: для записи, что приглашения были отправлены в Power Automate.
1. <span data-ttu-id="5fadc-151">Сохраните поток и попробуйте его. Используйте **кнопку Test** на странице редактора потока или запустите поток через вкладку **Мои потоки.** Не забудьте разрешить доступ при запросе.</span><span class="sxs-lookup"><span data-stu-id="5fadc-151">Save the flow and try it out. Use the **Test** button on the flow editor page or run the flow through your **My flows** tab. Be sure to allow access when prompted.</span></span>

## <a name="training-video-send-a-teams-meeting-from-excel-data"></a><span data-ttu-id="5fadc-152">Обучающее видео: отправка Teams собрания из Excel данных</span><span class="sxs-lookup"><span data-stu-id="5fadc-152">Training video: Send a Teams meeting from Excel data</span></span>

<span data-ttu-id="5fadc-153">[Смотреть Sudhi Ramamurthy ходить через версию этого примера на YouTube](https://youtu.be/HyBdx52NOE8).</span><span class="sxs-lookup"><span data-stu-id="5fadc-153">[Watch Sudhi Ramamurthy walk through a version of this sample on YouTube](https://youtu.be/HyBdx52NOE8).</span></span> <span data-ttu-id="5fadc-154">В его версии используется более надежный скрипт, который обрабатывает изменение столбцов и устаревшее время собраний.</span><span class="sxs-lookup"><span data-stu-id="5fadc-154">His version uses a more robust script that handles changing columns and obsolete meeting times.</span></span>
