---
title: Планирование собеседований в Teams
description: Узнайте, как использовать сценарии Office для отправки собрания Teams из данных Excel.
ms.date: 06/29/2021
ms.localizationpriority: medium
ms.openlocfilehash: 8e8c4af40398842e219dc3e2a80c6d2ee72d6b83
ms.sourcegitcommit: a6504f8b0d6b717457c6e0b5306c35ad3900914e
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/02/2022
ms.locfileid: "67572579"
---
# <a name="office-scripts-sample-scenario-schedule-interviews-in-teams"></a>Пример сценария сценариев Office: планирование собеседований в Teams

В этом сценарии вы являетесь сотрудником отдела кадров, планируете собрания по собеседованиям с кандидатами в Teams. Вы управляете расписанием собеседования кандидатов в файле Excel. Вам потребуется отправить приглашение на собрание Teams как кандидату, так и участникам собеседования. Затем необходимо обновить файл Excel с подтверждением отправки собраний Teams.

Решение включает три шага, объединенных в один поток Power Automate.

1. Скрипт извлекает данные из таблицы и возвращает массив объектов в виде [данных JSON](https://www.w3schools.com/whatis/whatis_json.asp) .
1. Затем данные отправляются в действие "Создание собрания **Teams** " для отправки приглашений.
1. Те же данные JSON отправляются другому скрипту для обновления состояния приглашения.

Дополнительные сведения о работе с JSON см. в статье "Использование JSON для передачи данных в скрипты [Office и из них"](../../develop/use-json.md).

## <a name="scripting-skills-covered"></a>Рассматриваются навыки навыков на написание скриптов

* Потоки Power Automate
* Интеграция Teams
* Синтаксический анализ таблицы

## <a name="sample-excel-file"></a>Пример файла Excel

Скачайте файл [hr-schedule.xlsx](hr-schedule.xlsx) который используется в этом решении, и попробуйте сами! Обязательно измените хотя бы один из адресов электронной почты, чтобы получить приглашение.

## <a name="sample-code-extract-table-data-to-schedule-invites"></a>Пример кода. Извлечение данных таблицы для планирования приглашений

Добавьте этот скрипт в коллекцию скриптов. Приведите **для потока имя Schedule Interviews** (Расписание собеседований).

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

## <a name="sample-code-mark-rows-as-invited"></a>Пример кода: пометка строк как приглашенных

Добавьте этот скрипт в коллекцию скриптов. Приведите для потока имя **"** Запись отправленных приглашений".

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

## <a name="sample-flow-run-the-interview-scheduling-scripts-and-send-the-teams-meetings"></a>Пример потока: запуск сценариев планирования собеседования и отправка собраний Teams

1. Создайте новый **мгновенный облачный поток**.
1. Выберите **"Вручную активировать поток" и** нажмите кнопку **"Создать"**.
1. Добавьте новый **шаг, использующий** **соединитель Excel Online (бизнес)** и действие **запуска скрипта** . Заполните соединитель следующими значениями.
    1. **Расположение**: OneDrive для бизнеса
    1. **Библиотека документов**: OneDrive
    1. **Файл**: hr-interviews.xlsx *(выбирается через браузер файлов)*
    1. **Сценарий**. Снимок экрана: снимок экрана с завершенным соединителем Excel Online (бизнес) для получения данных собеседования из книги :::image type="content" source="../../images/schedule-interviews-1.png" alt-text="в Power Automate.":::
1. Добавьте новый **шаг,** использующий **действие "Создание собрания Teams** ". При выборе динамического содержимого из соединителя Excel для потока будет создаваться приложение **Apply к** каждому блоку. Заполните соединитель следующими значениями.
    1. **Идентификатор календаря**: Календарь
    1. **Тема**: собеседование Contoso
    1. **Сообщение**: **сообщение** (значение Excel)
    1. **Часовой пояс**: тихоокеанское стандартное время
    1. **Время начала**: **startTime** (значение Excel)
    1. **Время окончания**: **FinishTime** (значение Excel)
    1. **Обязательные участники**: **CandidateEmail** ; **InterviewerEmail** (значения Excel) Снимок экрана: завершенный соединитель Teams для планирования собраний :::image type="content" source="../../images/schedule-interviews-2.png" alt-text="в Power Automate.":::
1. В том же **разделе "Применить к каждому блоку** " добавьте еще один соединитель **Excel Online (бизнес)** с действием **запуска скрипта** . Используйте следующие значения.
    1. **Расположение**: OneDrive для бизнеса
    1. **Библиотека документов**: OneDrive
    1. **Файл**: hr-interviews.xlsx *(выбирается через браузер файлов)*
    1. **Скрипт**: запись отправленных приглашений
    1. **invites**: **результат** (значение Excel) Снимок экрана готового соединителя Excel Online (business) для записи, что приглашения были отправлены :::image type="content" source="../../images/schedule-interviews-3.png" alt-text="в Power Automate.":::
1. Сохраните поток и попробуйте его. Нажмите **кнопку "** Тест" на странице редактора потоков или запустите поток на **вкладке "Мои потоки** ". Не забудьте разрешить доступ при появлении запроса.

## <a name="training-video-send-a-teams-meeting-from-excel-data"></a>Обучающее видео: отправка собрания Teams из данных Excel

[Просмотрите версию этого примера на YouTube](https://youtu.be/HyBdx52NOE8). В его версии используется более надежный сценарий, который обрабатывает изменение столбцов и устаревшее время собрания.
