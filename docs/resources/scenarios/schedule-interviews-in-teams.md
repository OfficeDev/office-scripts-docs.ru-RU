---
title: Планирование собеседований в Teams
description: Узнайте, как использовать Office скрипты для отправки собрания Teams из Excel данных.
ms.date: 06/29/2021
ms.localizationpriority: medium
ms.openlocfilehash: 25b70f2ee3f71c101d4ee20068c020edb5e0ac77
ms.sourcegitcommit: 7023b9e23499806901a5ecf8ebc460b76887cca6
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/31/2022
ms.locfileid: "64585431"
---
# <a name="office-scripts-sample-scenario-schedule-interviews-in-teams"></a>Office сценария сценариев: расписание интервью в Teams

В этом сценарии вы будете сотрудником по подбору кадров, запланируя встречи с кандидатами в Teams. Вы управляете расписанием собеседований кандидатов в Excel файле. Необходимо отправить приглашение на Teams как кандидату, так и интервьюеры. Затем необходимо обновить файл Excel с подтверждением того, что Teams были отправлены собрания.

Решение состоит из трех этапов, объединенных в один Power Automate потока.

1. Скрипт извлекает данные из таблицы и возвращает массив объектов в качестве данных JSON.
1. Затем данные отправляются в Teams **создать Teams** собрания для отправки приглашений.
1. Эти же данные JSON отправляются в другой скрипт, чтобы обновить состояние приглашения.

## <a name="scripting-skills-covered"></a>Навыки скриптов, охватываемых

* Power Automate потоки
* Teams интеграции
* Размыв таблиц

## <a name="sample-excel-file"></a>Пример Excel файла

Скачайте файл <a href="hr-schedule.xlsx">hr-schedule.xlsx</a> , используемый в этом решении, и попробуйте его самостоятельно! Обязательно измените хотя бы один из адресов электронной почты, чтобы получить приглашение.

## <a name="sample-code-extract-table-data-to-schedule-invites"></a>Пример кода. Извлечение данных таблицы для расписания приглашений

Добавьте этот скрипт в свою коллекцию скриптов. Назови **его Schedule Interviews** для потока.

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

## <a name="sample-code-mark-rows-as-invited"></a>Пример кода: пометить строки как приглашенные

Добавьте этот скрипт в свою коллекцию скриптов. Назови **его Запись отправленных приглашений** для потока.

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

## <a name="sample-flow-run-the-interview-scheduling-scripts-and-send-the-teams-meetings"></a>Пример потока: запустите сценарии планирования интервью и отправьте Teams собрания

1. Создайте новый **поток мгновенных облаков**.
1. Выберите **вручную вызвать поток и** выберите **Создать**.
1. Добавьте новый **шаг,** использующий **соединителю Excel Online (Бизнес)** и действие **скрипта Run**. Заполнять соединитектор следующими значениями.
    1. **Расположение**: OneDrive для бизнеса
    1. **Библиотека документов**: OneDrive
    1. **Файл**: hr-interviews.xlsx *(выбранный через браузер файлов)*
    1. **Сценарий**. Запланировать скриншот интервью с завершенным :::image type="content" source="../../images/schedule-interviews-1.png" alt-text="соединитетелем Excel Online (Бизнес)"::: для получения данных интервью из книги в Power Automate.
1. Добавьте новый **шаг,** использующий действие **Create a Teams собрания**. При выборе динамического контента Excel соединители, для каждого блока создается применение к каждому блоку. Заполнять соединитектор следующими значениями.
    1. **ID календаря**: Календарь
    1. **Тема**: Интервью Contoso
    1. **Сообщение**. **Сообщение** (Excel)
    1. **Часовой пояс**: тихоокеанское стандартное время
    1. **Время начала**: **StartTime** (Excel значение)
    1. **End time**: **FinishTime** (Excel значение)
    1. **Необходимые участники**: **CandidateEmail**; **InterviewerEmail** (Excel значений) Снимок экрана завершенного соединиттеля Teams для расписания собраний :::image type="content" source="../../images/schedule-interviews-2.png" alt-text="в Power Automate.":::
1. В том же **режиме Применяйте к каждому** блоку **другой соединителю Excel Online (Business)** с действием **сценария Run**. Используйте следующие значения.
    1. **Расположение**: OneDrive для бизнеса
    1. **Библиотека документов**: OneDrive
    1. **Файл**: hr-interviews.xlsx *(выбранный через браузер файлов)*
    1. **Сценарий**: Запись отправленных приглашений
    1. **приглашает****: результат** (Excel) Снимок экрана завершенного соединиттеля :::image type="content" source="../../images/schedule-interviews-3.png" alt-text="Excel Online (Бизнес)"::: для записи, что приглашения были отправлены в Power Automate.
1. Сохраните поток и попробуйте его. Используйте **кнопку Test** на странице редактора потока или запустите поток через вкладку **Мои потоки** . Не забудьте разрешить доступ при запросе.

## <a name="training-video-send-a-teams-meeting-from-excel-data"></a>Обучающее видео: отправка Teams собрания из Excel данных

[Просмотрите версию](https://youtu.be/HyBdx52NOE8) этого примера на YouTube. В его версии используется более надежный скрипт, который обрабатывает изменение столбцов и устаревшее время собраний.
