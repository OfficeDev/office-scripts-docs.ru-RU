---
title: Использование внешних вызовов для Office скриптов
description: Узнайте, как делать внешние вызовы API в Office Скрипты.
ms.date: 04/05/2021
localization_priority: Normal
ms.openlocfilehash: a77ceb61c2ff46a7b6226b798462b7be2c8e1c54
ms.sourcegitcommit: 1f003c9924e651600c913d84094506125f1055ab
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 04/26/2021
ms.locfileid: "52026996"
---
# <a name="use-external-fetch-calls-in-office-scripts"></a><span data-ttu-id="67c21-103">Использование внешних вызовов для Office скриптов</span><span class="sxs-lookup"><span data-stu-id="67c21-103">Use external fetch calls in Office Scripts</span></span>

<span data-ttu-id="67c21-104">Этот скрипт получает основные сведения о репозиториях GitHub пользователя.</span><span class="sxs-lookup"><span data-stu-id="67c21-104">This script gets basic information about a user's GitHub repositories.</span></span> <span data-ttu-id="67c21-105">В нем показано, как `fetch` использовать в простом сценарии.</span><span class="sxs-lookup"><span data-stu-id="67c21-105">It shows how to use `fetch` in a simple scenario.</span></span>

<span data-ttu-id="67c21-106">Дополнительные данные о API GItHub, используемых в ссылке GitHub [API.](https://docs.github.com/rest/reference/repos#list-repositories-for-a-user)</span><span class="sxs-lookup"><span data-stu-id="67c21-106">You can learn more about the GItHub APIs being used in the [GitHub API reference](https://docs.github.com/rest/reference/repos#list-repositories-for-a-user).</span></span> <span data-ttu-id="67c21-107">Вы также можете увидеть необработанный результат вызова API, посетив веб-браузер (не забудьте заменить местообладатель `https://api.github.com/users/{USERNAME}/repos` {USERNAME} на код Github).</span><span class="sxs-lookup"><span data-stu-id="67c21-107">You can also see the raw API call output by visiting `https://api.github.com/users/{USERNAME}/repos` in a web browser (be sure to replace the {USERNAME} placeholder with your Github ID).</span></span>

![Пример получения данных репозиториев](../../images/git.png)

## <a name="sample-code-get-basic-information-about-users-github-repositories"></a><span data-ttu-id="67c21-109">Пример кода. Получите базовую информацию о GitHub хранилищах пользователя</span><span class="sxs-lookup"><span data-stu-id="67c21-109">Sample code: Get basic information about user's GitHub repositories</span></span>

```TypeScript
async function main(workbook: ExcelScript.Workbook) {

  // Replace the {USERNAME} placeholder with your GitHub username.
  const response = await fetch('https://api.github.com/users/{USERNAME}/repos');
  const repos: Repository[] = await response.json();
  
  const rows: (string | boolean | number)[][] = [];
  for (let repo of repos){ 
    rows.push([repo.id, repo.name, repo.license?.name, repo.license?.url])
  }
  const sheet = workbook.getActiveWorksheet();
  const range = sheet.getRange('A2').getResizedRange(rows.length - 1, rows[0].length - 1);
  range.setValues(rows);
  return;
}

interface Repository {
  name: string,
  id: string,
  license?: License 
}

interface License {
  name: string,
  url: string
}
```

## <a name="training-video-how-to-make-external-api-calls"></a><span data-ttu-id="67c21-110">Обучающее видео: как сделать внешние вызовы API</span><span class="sxs-lookup"><span data-stu-id="67c21-110">Training video: How to make external API calls</span></span>

<span data-ttu-id="67c21-111">[![Просмотр видео о том, как делать внешние вызовы API](../../images/api-vid.png)](https://youtu.be/fulP29J418E "Видео о том, как делать внешние вызовы API")</span><span class="sxs-lookup"><span data-stu-id="67c21-111">[![Watch video on how to make external API calls](../../images/api-vid.png)](https://youtu.be/fulP29J418E "Video on how to make external API calls")</span></span>
