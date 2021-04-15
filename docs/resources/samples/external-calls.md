---
title: Вызовы внешнего API в сценариях Office
description: Узнайте, как делать внешние вызовы API в скриптах Office.
ms.date: 03/30/2021
localization_priority: Normal
ms.openlocfilehash: 0ed57ed3b97309dbb7ea196695dcc347e133b3cf
ms.sourcegitcommit: 45ffe3dbd2c834b78592ad35928cf8096f5e80bc
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 04/14/2021
ms.locfileid: "51754805"
---
# <a name="external-api-calls-from-office-scripts"></a><span data-ttu-id="f14e1-103">Внешние вызовы API из office Scripts</span><span class="sxs-lookup"><span data-stu-id="f14e1-103">External API calls from Office Scripts</span></span>

<span data-ttu-id="f14e1-104">Скрипты Office позволяют поддерживать [ограниченные внешние вызовы API.](../../develop/external-calls.md)</span><span class="sxs-lookup"><span data-stu-id="f14e1-104">Office Scripts allows [limited external API call support](../../develop/external-calls.md).</span></span>

> [!IMPORTANT]
>
> * <span data-ttu-id="f14e1-105">Нет способа войти или использовать потоки проверки подлинности OAuth2.</span><span class="sxs-lookup"><span data-stu-id="f14e1-105">There is no way to sign in or use OAuth2 type of authentication flows.</span></span> <span data-ttu-id="f14e1-106">Все ключи и учетные данные должны быть жестко закодированы (или считываться из другого источника).</span><span class="sxs-lookup"><span data-stu-id="f14e1-106">All keys and credentials have to be hardcoded (or read from another source).</span></span>
> * <span data-ttu-id="f14e1-107">Нет инфраструктуры для хранения учетных данных и ключей API.</span><span class="sxs-lookup"><span data-stu-id="f14e1-107">There is no infrastructure to store API credentials and keys.</span></span> <span data-ttu-id="f14e1-108">Этим должен управлять пользователь.</span><span class="sxs-lookup"><span data-stu-id="f14e1-108">This will have to be managed by the user.</span></span>
> * <span data-ttu-id="f14e1-109">Внешние вызовы могут привести к воздействию конфиденциальных данных на нежелательные конечные точки или внешним данным, которые будут занесены во внутренние книги.</span><span class="sxs-lookup"><span data-stu-id="f14e1-109">External calls may result in sensitive data being exposed to undesirable endpoints, or external data to be brought into internal workbooks.</span></span> <span data-ttu-id="f14e1-110">Администратор может установить защиту брандмауэра от таких вызовов.</span><span class="sxs-lookup"><span data-stu-id="f14e1-110">Your admin can establish firewall protection against such calls.</span></span> <span data-ttu-id="f14e1-111">Не забудьте проверить местные политики, прежде чем полагаться на внешние вызовы.</span><span class="sxs-lookup"><span data-stu-id="f14e1-111">Be sure to check with local policies prior to relying on external calls.</span></span>
> * <span data-ttu-id="f14e1-112">Если сценарий использует вызов API, он не будет работать в сценарии Power Automate.</span><span class="sxs-lookup"><span data-stu-id="f14e1-112">If a script uses an API call, it will not function in a Power Automate scenario.</span></span> <span data-ttu-id="f14e1-113">Для получения данных из внешней службы необходимо использовать действие http или эквивалентные действия Power Automate.</span><span class="sxs-lookup"><span data-stu-id="f14e1-113">You'll have to use Power Automate's HTTP action or equivalent actions to pull data from or push it to an external service.</span></span>
> * <span data-ttu-id="f14e1-114">Внешний вызов API включает асинхронный синтаксис API и требует немного расширенных знаний о том, как работает коммуникация async.</span><span class="sxs-lookup"><span data-stu-id="f14e1-114">An external API call involves asynchronous API syntax and requires slightly advanced knowledge of the way async communication works.</span></span>
> * <span data-ttu-id="f14e1-115">Убедитесь, что перед принятием зависимости необходимо проверить объем пропускной способности данных.</span><span class="sxs-lookup"><span data-stu-id="f14e1-115">Be sure to check the amount of data throughput prior to taking a dependency.</span></span> <span data-ttu-id="f14e1-116">Например, стягивание всего внешнего наборов данных может оказаться не самым лучшим вариантом, а вместо этого следует использовать pagination для получения данных в куски.</span><span class="sxs-lookup"><span data-stu-id="f14e1-116">For instance, pulling down the entire external dataset may not be the best option and instead pagination should be used to get data in chunks.</span></span>

## <a name="useful-knowledge-and-resources"></a><span data-ttu-id="f14e1-117">Полезные знания и ресурсы</span><span class="sxs-lookup"><span data-stu-id="f14e1-117">Useful knowledge and resources</span></span>

* <span data-ttu-id="f14e1-118">[REST API.](https://en.wikipedia.org/wiki/Representational_state_transfer)Скорее всего, вы будете использовать вызов API.</span><span class="sxs-lookup"><span data-stu-id="f14e1-118">[REST API](https://en.wikipedia.org/wiki/Representational_state_transfer): Most likely way you'll use the API call.</span></span>
* <span data-ttu-id="f14e1-119">: Понять, как это работает. [ `async` `await` ](https://developer.mozilla.org/docs/Learn/JavaScript/Asynchronous/Async_await)</span><span class="sxs-lookup"><span data-stu-id="f14e1-119">[`async` `await`](https://developer.mozilla.org/docs/Learn/JavaScript/Asynchronous/Async_await): Understand how this works.</span></span>
* <span data-ttu-id="f14e1-120">[`fetch`](https://developer.mozilla.org/docs/Web/API/Fetch_API/Using_Fetch): Понять, как это работает.</span><span class="sxs-lookup"><span data-stu-id="f14e1-120">[`fetch`](https://developer.mozilla.org/docs/Web/API/Fetch_API/Using_Fetch): Understand how this works.</span></span>

## <a name="steps"></a><span data-ttu-id="f14e1-121">Шаги</span><span class="sxs-lookup"><span data-stu-id="f14e1-121">Steps</span></span>

1. <span data-ttu-id="f14e1-122">`main`Пометите функцию как асинхронную функцию, добавив `async` префикс.</span><span class="sxs-lookup"><span data-stu-id="f14e1-122">Mark your `main` function as an asynchronous function by adding `async` prefix.</span></span> <span data-ttu-id="f14e1-123">Например, `async function main(workbook: ExcelScript.Workbook)`.</span><span class="sxs-lookup"><span data-stu-id="f14e1-123">For example, `async function main(workbook: ExcelScript.Workbook)`.</span></span>
1. <span data-ttu-id="f14e1-124">Какой тип вызова API вы делаете?</span><span class="sxs-lookup"><span data-stu-id="f14e1-124">Which type of API call are you making?</span></span> <span data-ttu-id="f14e1-125">`GET`, `POST`, `PUT`, `DELETE`, `PATCH`?</span><span class="sxs-lookup"><span data-stu-id="f14e1-125">`GET`, `POST`, `PUT`, `DELETE`, `PATCH`?</span></span> <span data-ttu-id="f14e1-126">Подробные сведения можно найти в материале REST API.</span><span class="sxs-lookup"><span data-stu-id="f14e1-126">Refer to REST API material for details.</span></span>
1. <span data-ttu-id="f14e1-127">Получение конечной точки API службы, требований к проверке подлинности, заголовок и т. д.</span><span class="sxs-lookup"><span data-stu-id="f14e1-127">Obtain the service API endpoint, authentication requirements, headers, etc.</span></span>
1. <span data-ttu-id="f14e1-128">Определите входные данные или `interface` выходные данные, которые помогут с завершением кода и проверкой времени разработки.</span><span class="sxs-lookup"><span data-stu-id="f14e1-128">Define the input or output `interface` to help with code completion and development time verification.</span></span> <span data-ttu-id="f14e1-129">Подробные [сведения](#training-video-how-to-make-external-api-calls) см. в видео.</span><span class="sxs-lookup"><span data-stu-id="f14e1-129">See [video](#training-video-how-to-make-external-api-calls) for details.</span></span>
1. <span data-ttu-id="f14e1-130">Код, тест, оптимизация.</span><span class="sxs-lookup"><span data-stu-id="f14e1-130">Code, test, optimize.</span></span> <span data-ttu-id="f14e1-131">Вы можете создать функцию для обычного вызова API, чтобы сделать ее многоразовой из других частей скрипта или повторного использования в другом скрипте (таким образом скопировать-вклеить становится намного проще).</span><span class="sxs-lookup"><span data-stu-id="f14e1-131">You can create a function for your API call routine to make it reusable from other parts of your script or for reuse in a different script (copy-paste becomes much easier this way).</span></span>

## <a name="scenario"></a><span data-ttu-id="f14e1-132">Сценарий</span><span class="sxs-lookup"><span data-stu-id="f14e1-132">Scenario</span></span>

<span data-ttu-id="f14e1-133">Этот скрипт получает основные сведения о репозиториях GitHub пользователя.</span><span class="sxs-lookup"><span data-stu-id="f14e1-133">This script gets basic information about the user's GitHub repositories.</span></span>

## <a name="resources-used-in-the-sample"></a><span data-ttu-id="f14e1-134">Ресурсы, используемые в примере</span><span class="sxs-lookup"><span data-stu-id="f14e1-134">Resources used in the sample</span></span>

1. [<span data-ttu-id="f14e1-135">Получите ссылку на API Github для репозиториев.</span><span class="sxs-lookup"><span data-stu-id="f14e1-135">Get repositories Github API reference.</span></span>](https://docs.github.com/rest/reference/repos#list-repositories-for-a-user)
1. <span data-ttu-id="f14e1-136">Выход вызова API: перейдите в веб-браузер или любой интерфейс HTTP и введите, заменив местообладатель `https://api.github.com/users/{USERNAME}/repos` {USERNAME} вашим кодом Github.</span><span class="sxs-lookup"><span data-stu-id="f14e1-136">API call output: Go to a web browser or any HTTP interface and type in `https://api.github.com/users/{USERNAME}/repos`, replacing the {USERNAME} placeholder with your Github ID.</span></span>
1. <span data-ttu-id="f14e1-137">Извлеченные сведения: repo.name, repo.size, repo.owner.id, repo.license?. имя</span><span class="sxs-lookup"><span data-stu-id="f14e1-137">Information fetched: repo.name, repo.size, repo.owner.id, repo.license?.name</span></span>

## <a name="sample-code-get-basic-information-about-users-github-repositories"></a><span data-ttu-id="f14e1-138">Пример кода. Получите основные сведения о репозиториях GitHub пользователя</span><span class="sxs-lookup"><span data-stu-id="f14e1-138">Sample code: Get basic information about user's GitHub repositories</span></span>

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

## <a name="training-video-how-to-make-external-api-calls"></a><span data-ttu-id="f14e1-139">Обучающее видео: как сделать внешние вызовы API</span><span class="sxs-lookup"><span data-stu-id="f14e1-139">Training video: How to make external API calls</span></span>

<span data-ttu-id="f14e1-140">[![Просмотр видео о том, как делать внешние вызовы API](../../images/api-vid.png)](https://youtu.be/fulP29J418E "Видео о том, как делать внешние вызовы API")</span><span class="sxs-lookup"><span data-stu-id="f14e1-140">[![Watch video on how to make external API calls](../../images/api-vid.png)](https://youtu.be/fulP29J418E "Video on how to make external API calls")</span></span>
