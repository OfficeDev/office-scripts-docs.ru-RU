---
title: Вызовы внешнего API в сценариях Office
description: Узнайте, как делать внешние вызовы API в скриптах Office.
ms.date: 03/30/2021
localization_priority: Normal
ms.openlocfilehash: d0abfa0bb1adedc7535059ed359b8053d9f1c84d
ms.sourcegitcommit: 5d24e77df70aa2c1c982275d53213c2a9323ff86
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 04/02/2021
ms.locfileid: "51571470"
---
# <a name="external-api-calls-from-office-scripts"></a>Внешние вызовы API из office Scripts

Скрипты Office позволяют поддерживать [ограниченные внешние вызовы API.](../../develop/external-calls.md)

> [!IMPORTANT]
>
> * Нет способа войти или использовать потоки проверки подлинности OAuth2. Все ключи и учетные данные должны быть жестко закодированы (или считываться из другого источника).
> * Нет инфраструктуры для хранения учетных данных и ключей API. Этим должен управлять пользователь.
> * Внешние вызовы могут привести к воздействию конфиденциальных данных на нежелательные конечные точки или внешним данным, которые будут занесены во внутренние книги. Администратор может установить защиту брандмауэра от таких вызовов. Не забудьте проверить местные политики, прежде чем полагаться на внешние вызовы.
> * Если сценарий использует вызов API, он не будет работать в сценарии Power Automate. Для получения данных из внешней службы необходимо использовать действие http или эквивалентные действия Power Automate.
> * Внешний вызов API включает асинхронный синтаксис API и требует немного расширенных знаний о том, как работает коммуникация async.
> * Убедитесь, что перед принятием зависимости необходимо проверить объем пропускной способности данных. Например, стягивание всего внешнего наборов данных может оказаться не самым лучшим вариантом, а вместо этого следует использовать pagination для получения данных в куски.

## <a name="useful-knowledge-and-resources"></a>Полезные знания и ресурсы

* [REST API.](https://en.wikipedia.org/wiki/Representational_state_transfer)Скорее всего, вы будете использовать вызов API.
* : Понять, как это работает. [ `async` `await` ](https://developer.mozilla.org/docs/Learn/JavaScript/Asynchronous/Async_await)
* [`fetch`](https://developer.mozilla.org/docs/Web/API/Fetch_API/Using_Fetch): Понять, как это работает.

## <a name="steps"></a>Шаги

1. `main`Пометите функцию как асинхронную функцию, добавив `async` префикс. Например, `async function main(workbook: ExcelScript.Workbook)`.
1. Какой тип вызова API вы делаете? `GET`, `POST`, `PUT`, `DELETE`, `PATCH`? Подробные сведения можно найти в материале REST API.
1. Получение конечной точки API службы, требований к проверке подлинности, заголовок и т. д.
1. Определите входные данные или `interface` выходные данные, которые помогут с завершением кода и проверкой времени разработки. Подробные [сведения](#training-video-how-to-make-external-api-calls) см. в видео.
1. Код, тест, оптимизация. Вы можете создать функцию для обычного вызова API, чтобы сделать ее многоразовой из других частей скрипта или повторного использования в другом скрипте (таким образом скопировать-вклеить становится намного проще).

## <a name="scenario"></a>Сценарий

Этот скрипт получает основные сведения о репозиториях GitHub пользователя.

![Пример получения данных репозиториев](../../images/git.png)

## <a name="resources-used-in-the-sample"></a>Ресурсы, используемые в примере

1. [Получите ссылку на API Github для репозиториев.](https://docs.github.com/rest/reference/repos#list-repositories-for-a-user)
1. Выход вызова API: перейдите в веб-браузер или любой интерфейс HTTP и введите, заменив местообладатель `https://api.github.com/users/{USERNAME}/repos` {USERNAME} вашим кодом Github.
1. Извлеченные сведения: repo.name, repo.size, repo.owner.id, repo.license?. имя

## <a name="sample-code-get-basic-information-about-users-github-repositories"></a>Пример кода. Получите основные сведения о репозиториях GitHub пользователя

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

## <a name="training-video-how-to-make-external-api-calls"></a>Обучающее видео: как сделать внешние вызовы API

[![Просмотр видео о том, как делать внешние вызовы API](../../images/api-vid.png)](https://youtu.be/fulP29J418E "Видео о том, как делать внешние вызовы API")
