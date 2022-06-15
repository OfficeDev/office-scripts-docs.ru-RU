---
title: Использование внешних вызовов Fetch в сценариях Office
description: Узнайте, как выполнять внешние вызовы API в Office скриптах.
ms.date: 06/10/2022
ms.localizationpriority: medium
ms.openlocfilehash: 569d74f1ca8996cd8fe8a4ba3163445d57676d27
ms.sourcegitcommit: dd01979d34b3499360d2f79a56f8a8f24f480eed
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 06/15/2022
ms.locfileid: "66088094"
---
# <a name="use-external-fetch-calls-in-office-scripts"></a>Использование внешних вызовов Fetch в сценариях Office

Этот сценарий получает основные сведения о пользовательском GitHub репозиториях. В нем показано, как использовать `fetch` в простом сценарии. Дополнительные сведения об использовании или других внешних `fetch` вызовах см. в разделе "Поддержка вызовов [внешнего API" Office scripts](../../develop/external-calls.md). Сведения о работе с объектами [JSON]] (https://www.w3schools.com/whatis/whatis_json.asp)например, с объектами, возвращаемые интерфейсами API GitHub, см. в статье "Использование JSON для передачи данных в Office [скрипты и из них](../../develop/use-json.md)".

Дополнительные сведения об API GItHub, используемых в справочнике GitHub [API](https://docs.github.com/rest/reference/repos#list-repositories-for-a-user). Вы также можете просмотреть необработанные выходные данные вызова API`https://api.github.com/users/{USERNAME}/repos`, посетив веб-браузер (обязательно замените заполнитель {USERNAME} идентификатором GitHub).

![Пример получения сведений о репозиториях](../../images/git.png)

## <a name="sample-code-get-basic-information-about-users-github-repositories"></a>Пример кода: получение основных сведений о репозиториях GitHub пользователя

```TypeScript
async function main(workbook: ExcelScript.Workbook) {
  // Call the GitHub REST API.
  // Replace the {USERNAME} placeholder with your GitHub username.
  const response = await fetch('https://api.github.com/users/{USERNAME}/repos');
  const repos: Repository[] = await response.json();

  // Create an array to hold the returned values.
  const rows: (string | boolean | number)[][] = [];

  // Convert each repository block into a row.
  for (let repo of repos) {
    rows.push([repo.id, repo.name, repo.license?.name, repo.license?.url]);
  }
  // Create a header row.
  const sheet = workbook.getActiveWorksheet();
  sheet.getRange('A1:D1').setValues([["ID", "Name", "License Name", "License URL"]]);

  // Add the data to the current worksheet, starting at "A2".
  const range = sheet.getRange('A2').getResizedRange(rows.length - 1, rows[0].length - 1);
  range.setValues(rows);
}

// An interface matching the returned JSON for a GitHub repository.
interface Repository {
  name: string,
  id: string,
  license?: License
}

// An interface matching the returned JSON for a GitHub repo license.
interface License {
  name: string,
  url: string
}
```

## <a name="training-video-how-to-make-external-api-calls"></a>Обучающее видео: как выполнять внешние вызовы API

[Просмотрите этот пример на YouTube](https://youtu.be/fulP29J418E), чтобы просмотреть этот пример.
