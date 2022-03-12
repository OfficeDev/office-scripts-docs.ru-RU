---
title: Поддержка внешнего вызова API в сценариях Office
description: Поддержка и руководство по ведению внешних вызовов API в Office Скрипт.
ms.date: 05/21/2021
ms.localizationpriority: medium
ms.openlocfilehash: e7be505f13529e1d3bcff22ce9fa18cc36148f7b
ms.sourcegitcommit: 79ce4fad6d284b1aa71f5ad6d2938d9ad6a09fee
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/12/2022
ms.locfileid: "63459608"
---
# <a name="external-api-call-support-in-office-scripts"></a>Поддержка внешнего вызова API в сценариях Office

Скрипты поддерживают вызовы внешних служб. Используйте эти службы для поставок данных и других сведений в вашу книгу.

> [!CAUTION]
> Внешние вызовы могут привести к воздействию конфиденциальных данных на нежелательные конечные точки. Администратор может установить защиту брандмауэра от таких вызовов.

> [!IMPORTANT]
> Вызовы на внешние API можно делать только через приложение Excel, а не Power Automate [при обычных обстоятельствах](#external-calls-from-power-automate).

## <a name="configure-your-script-for-external-calls"></a>Настройка сценария для внешних вызовов

Внешние вызовы [асинхронны](https://developer.mozilla.org/docs/Learn/JavaScript/Asynchronous/Async_await) и требуют, чтобы сценарий был помечен как `async`. Добавьте префикс `async` в функцию `main` и `Promise`вернетесь, как показано здесь:

```typescript
async function main(workbook: ExcelScript.Workbook) : Promise <void>
```

> [!NOTE]
> Скрипты, возвращая другие сведения, могут возвращать `Promise` один из этих типов. Например, если сценарию необходимо вернуть `Employee` объект, возвращаемая подпись будет `: Promise <Employee>`

Чтобы звонить в эту службу, необходимо изучить интерфейсы внешней службы. Если вы используете API `fetch` [или REST](https://wikipedia.org/wiki/Representational_state_transfer), необходимо определить структуру JSON возвращаемого данных. Для ввода и вывода из скрипта `interface` рассмотрите возможность создания, чтобы соответствовать необходимым структурам JSON. Это обеспечивает скрипту больше безопасности типа. Пример этого см. в примере [Using fetch from Office Scripts](../resources/samples/external-fetch-calls.md).

### <a name="limitations-with-external-calls-from-office-scripts"></a>Ограничения внешних вызовов из Office скриптов

* Нет способа войти или использовать потоки проверки подлинности OAuth2. Все ключи и учетные данные должны быть жестко закодированы (или считываться из другого источника).
* Нет инфраструктуры для хранения учетных данных и ключей API. Этим должен управлять пользователь.
* Файлы cookie документов и `localStorage`объекты `sessionStorage` не поддерживаются.
* Внешние вызовы могут привести к воздействию конфиденциальных данных на нежелательные конечные точки или внешним данным, которые будут занесены во внутренние книги. Администратор может установить защиту брандмауэра от таких вызовов. Не забудьте проверить местные политики, прежде чем полагаться на внешние вызовы.
* Убедитесь, что перед принятием зависимости необходимо проверить объем пропускной способности данных. Например, стягивание всего внешнего наборов данных может оказаться не самым лучшим вариантом, а вместо этого следует использовать pagination для получения данных в куски.

## <a name="retrieve-information-with-fetch"></a>Извлечение сведений с помощью `fetch`

[API извлекает](https://developer.mozilla.org/docs/Web/API/Fetch_API) сведения из внешних служб. Это API `async` , поэтому необходимо настроить подпись `main` скрипта. Сделайте функцию `main` `async`. Вы также должны быть уверены в вызове `await` `fetch` и `json` ирисовке. Это обеспечивает завершение этих операций до завершения сценария.

Все полученные JSON данные должны `fetch` соответствовать интерфейсу, определенному в сценарии. Возвращенное значение должно быть назначено определенному типу, так как Office [скрипты не поддерживают `any` тип](typescript-restrictions.md#no-any-type-in-office-scripts). Чтобы узнать имена и типы возвращаемого свойства, необходимо обратиться к документации для службы. Затем добавьте в скрипт совпадающий интерфейс или интерфейс.

Следующий сценарий использует для `fetch` получения данных JSON с тестового сервера в заданном URL-адресе. Обратите внимание на `JSONData` интерфейс для хранения данных в качестве типа совпадения.

```TypeScript
async function main(workbook: ExcelScript.Workbook){
  // Retrieve sample JSON data from a test server.
  let fetchResult = await fetch('https://jsonplaceholder.typicode.com/todos/1');

  // Convert the returned data to the expected JSON structure.
  let json : JSONData = await fetchResult.json();

  // Display the content in a readable format.
  console.log(JSON.stringify(json));
}

/**
 * An interface that matches the returned JSON structure.
 * The property names match exactly.
 */
interface JSONData {
  userId: number;
  id: number;
  title: string;
  completed: boolean;
}
```

### <a name="other-fetch-samples"></a>Другие `fetch` примеры

* В [примере Use external fetch calls in Office Scripts](../resources/samples/external-fetch-calls.md) показано, как получить базовую информацию о GitHub хранилищах пользователя.
* Пример [сценария Office сценариев:](../resources/scenarios/noaa-data-fetch.md) Graph данных уровня воды из NOAA демонстрирует команду получения, используемую для получения записей из базы данных "Приливы и течения" Национального управления океанических и атмосферных исследований.

## <a name="external-calls-from-power-automate"></a>Внешние вызовы из Power Automate

Любой внешний вызов API не удается при запуске сценария с Power Automate. Это поведенческая разница между запуском скрипта через Excel и Power Automate. Убедитесь в том, что перед созданием их в поток необходимо проверить сценарии для таких ссылок.

Для получения данных из внешней службы необходимо использовать [HTTP с помощью Azure AD](/connectors/webcontents/) или других эквивалентных действий.

> [!WARNING]
> Внешние вызовы, сделанные через соединители Power Automate [Excel Online](/connectors/excelonlinebusiness), не удается поддерживать существующие политики предотвращения потери данных. Однако сценарии, которые Power Automate, делаются так за пределами организации и вне брандмауэров организации. Для дополнительной защиты от вредоносных пользователей в этой внешней среде администратор может управлять использованием Office скриптов. Администратор может отключить соединители Excel Online в Power Automate или отключить Office скрипты для Excel в Интернете с помощью элементов управления Office [скриптов](/microsoft-365/admin/manage/manage-office-scripts-settings).

## <a name="see-also"></a>См. также

* [Использование встроенных объектов JavaScript в сценариях Office](javascript-objects.md)
* [Использование внешних вызовов Fetch в сценариях Office](../resources/samples/external-fetch-calls.md)
* [Office сценарий сценариев: Graph данных уровня воды из NOAA](../resources/scenarios/noaa-data-fetch.md)
