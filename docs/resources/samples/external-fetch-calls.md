---
title: Использование внешних вызовов Fetch в сценариях Office
description: Узнайте, как делать внешние вызовы API в Office Скрипты.
ms.date: 05/14/2021
localization_priority: Normal
ms.openlocfilehash: df8814cbab16969a1140aecfe526fd68e609d43c
ms.sourcegitcommit: 4687693f02fc90a57ba30c461f35046e02e6f5fb
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 05/19/2021
ms.locfileid: "52545754"
---
# <a name="use-external-fetch-calls-in-office-scripts"></a><span data-ttu-id="b9b98-103">Использование внешних вызовов Fetch в сценариях Office</span><span class="sxs-lookup"><span data-stu-id="b9b98-103">Use external fetch calls in Office Scripts</span></span>

<span data-ttu-id="b9b98-104">Этот скрипт получает основные сведения о репозиториях GitHub пользователя.</span><span class="sxs-lookup"><span data-stu-id="b9b98-104">This script gets basic information about a user's GitHub repositories.</span></span> <span data-ttu-id="b9b98-105">В нем показано, как `fetch` использовать в простом сценарии.</span><span class="sxs-lookup"><span data-stu-id="b9b98-105">It shows how to use `fetch` in a simple scenario.</span></span> <span data-ttu-id="b9b98-106">Дополнительные сведения об использовании или других внешних вызовах читайте в публикации Поддержка вызовов внешнего `fetch` [API в Office Scripts](../../develop/external-calls.md)</span><span class="sxs-lookup"><span data-stu-id="b9b98-106">For more information about using `fetch` or other external calls, read [External API call support in Office Scripts](../../develop/external-calls.md)</span></span>

<span data-ttu-id="b9b98-107">Дополнительные данные о API GItHub, используемых в ссылке GitHub [API.](https://docs.github.com/rest/reference/repos#list-repositories-for-a-user)</span><span class="sxs-lookup"><span data-stu-id="b9b98-107">You can learn more about the GItHub APIs being used in the [GitHub API reference](https://docs.github.com/rest/reference/repos#list-repositories-for-a-user).</span></span> <span data-ttu-id="b9b98-108">Вы также можете увидеть необработанный вывод вызова API, посетив веб-браузер (не забудьте заменить местообладатель `https://api.github.com/users/{USERNAME}/repos` {USERNAME} своим GitHub ID).</span><span class="sxs-lookup"><span data-stu-id="b9b98-108">You can also see the raw API call output by visiting `https://api.github.com/users/{USERNAME}/repos` in a web browser (be sure to replace the {USERNAME} placeholder with your GitHub ID).</span></span>

![Пример получения данных репозиториев](../../images/git.png)

## <a name="sample-code-get-basic-information-about-users-github-repositories"></a><span data-ttu-id="b9b98-110">Пример кода. Получите базовую информацию о GitHub хранилищах пользователя</span><span class="sxs-lookup"><span data-stu-id="b9b98-110">Sample code: Get basic information about user's GitHub repositories</span></span>

```TypeScript
async function main(workbook: ExcelScript.Workbook) {
  // Call the GitHub REST API.
  // Replace the {USERNAME} placeholder with your GitHub username.
  const response = await fetch('https://api.github.com/users/{USERNAME}/repos');
  const repos: Repository[] = await response.json();
  
  // Create an array to hold the returned values.
  const rows: (string | boolean | number)[][] = [];

  // Convert each repository block into a row.
  for (let repo of repos){ 
    rows.push([repo.id, repo.name, repo.license?.name, repo.license?.url])
  }

  // Add the data to the current worksheet, starting at "A2".
  const sheet = workbook.getActiveWorksheet();
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

## <a name="training-video-how-to-make-external-api-calls"></a><span data-ttu-id="b9b98-111">Обучающее видео: как сделать внешние вызовы API</span><span class="sxs-lookup"><span data-stu-id="b9b98-111">Training video: How to make external API calls</span></span>

<span data-ttu-id="b9b98-112">[Смотреть Sudhi Ramamurthy ходить через этот пример на YouTube](https://youtu.be/fulP29J418E).</span><span class="sxs-lookup"><span data-stu-id="b9b98-112">[Watch Sudhi Ramamurthy walk through this sample on YouTube](https://youtu.be/fulP29J418E).</span></span>
