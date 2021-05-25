---
title: Поддержка внешнего вызова API в сценариях Office
description: Поддержка и руководство по принятию внешних вызовов API в Office скрипта.
ms.date: 05/21/2021
localization_priority: Normal
ms.openlocfilehash: 5d768b53112473c1774f8fe8257b197ffead4a63
ms.sourcegitcommit: 09d8859d5269ada8f1d0e141f6b5a4f96d95a739
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 05/24/2021
ms.locfileid: "52631645"
---
# <a name="external-api-call-support-in-office-scripts"></a>Поддержка внешнего вызова API в сценариях Office

Скрипты поддерживают вызовы внешних служб. Используйте эти службы для поставок данных и других сведений в вашу книгу.

> [!CAUTION]
> Внешние вызовы могут привести к воздействию конфиденциальных данных на нежелательные конечные точки. Администратор может установить защиту брандмауэра от таких вызовов.

> [!IMPORTANT]
> Вызовы к внешним API можно сделать только через Excel, а не Power Automate при [обычных обстоятельствах.](#external-calls-from-power-automate)

## <a name="configure-your-script-for-external-calls"></a>Настройка сценария для внешних вызовов

Внешние вызовы [асинхронны](https://developer.mozilla.org/docs/Learn/JavaScript/Asynchronous/Async_await) и требуют, чтобы сценарий был помечен как `async` . Добавьте `async` префикс в функцию и вернетесь, `main` `Promise` как показано здесь:

```typescript
async function main(workbook: ExcelScript.Workbook) : Promise <void>
```

> [!NOTE]
> Скрипты, возвращая другие сведения, могут возвращать `Promise` один из этих типов. Например, если сценарию необходимо вернуть `Employee` объект, возвращаемая подпись будет `: Promise <Employee>`

Чтобы звонить в эту службу, необходимо изучить интерфейсы внешней службы. Если вы используете API rest или REST, вам необходимо определить структуру `fetch` JSON возвращаемого данных. [](https://wikipedia.org/wiki/Representational_state_transfer) Для ввода и вывода из скрипта рассмотрите возможность создания, чтобы соответствовать необходимым `interface` структурам JSON. Это обеспечивает скрипту больше безопасности типа. Пример этого см. в примере [Using fetch from Office Scripts.](../resources/samples/external-fetch-calls.md)

### <a name="limitations-with-external-calls-from-office-scripts"></a>Ограничения внешних вызовов из Office скриптов

* Нет способа войти или использовать потоки проверки подлинности OAuth2. Все ключи и учетные данные должны быть жестко закодированы (или считываться из другого источника).
* Нет инфраструктуры для хранения учетных данных и ключей API. Этим должен управлять пользователь.
* Файлы cookie документов `localStorage` и `sessionStorage` объекты не поддерживаются.
* Внешние вызовы могут привести к воздействию конфиденциальных данных на нежелательные конечные точки или внешним данным, которые будут занесены во внутренние книги. Администратор может установить защиту брандмауэра от таких вызовов. Не забудьте проверить местные политики, прежде чем полагаться на внешние вызовы.
* Убедитесь, что перед принятием зависимости необходимо проверить объем пропускной способности данных. Например, стягивание всего внешнего наборов данных может оказаться не самым лучшим вариантом, а вместо этого следует использовать pagination для получения данных в куски.

## <a name="retrieve-information-with-fetch"></a>Извлечение сведений с помощью `fetch`

API [извлекает](https://developer.mozilla.org/docs/Web/API/Fetch_API) сведения из внешних служб. Это `async` API, поэтому необходимо настроить подпись `main` скрипта. Сделайте `main` `async` функцию и делайте так, чтобы она возвращала `Promise<void>` . Вы также должны быть уверены `await` в `fetch` вызове и `json` ирисовке. Это обеспечивает завершение этих операций до завершения сценария.

Все полученные JSON данные `fetch` должны соответствовать интерфейсу, определенному в сценарии. Возвращенное значение должно быть назначено определенному типу, так как Office скрипты [не поддерживают `any` тип](typescript-restrictions.md#no-any-type-in-office-scripts). Чтобы узнать имена и типы возвращаемого свойства, необходимо обратиться к документации для службы. Затем добавьте в скрипт совпадающий интерфейс или интерфейс.

Следующий сценарий использует для получения данных JSON с `fetch` тестового сервера в заданном URL-адресе. Обратите внимание `JSONData` на интерфейс для хранения данных в качестве типа совпадения.

```TypeScript
async function main(workbook: ExcelScript.Workbook): Promise<void> {
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
* Пример [сценария Office сценариев:](../resources/scenarios/noaa-data-fetch.md) Graph данных уровня воды из NOAA демонстрирует команду извлечения, используемую для получения записей из базы данных "Приливы и течения" Национального управления океанических и атмосферных исследований.

## <a name="external-calls-from-power-automate"></a>Внешние вызовы из Power Automate

Любой внешний вызов API не удается при запуске сценария с Power Automate. Это поведенческая разница между запуском скрипта через приложение Excel и Power Automate. Убедитесь в том, что перед созданием их в поток необходимо проверить сценарии для таких ссылок.

Для получения данных из внешней службы необходимо использовать HTTP с [помощью Azure AD](/connectors/webcontents/) или других эквивалентных действий.

> [!WARNING]
> Внешние вызовы, сделанные через соединители Power Automate [Excel Online,](/connectors/excelonlinebusiness) сбой, чтобы помочь поддерживать существующие политики предотвращения потери данных. Однако сценарии, которые Power Automate, делаются так за пределами организации и вне брандмауэров организации. Для дополнительной защиты от вредоносных пользователей в этой внешней среде администратор может управлять использованием Office скриптов. Администратор может отключить соединители Excel Online в Power Automate или отключить Office скрипты для Excel в Интернете с помощью элементов управления администратором [Office скриптов.](/microsoft-365/admin/manage/manage-office-scripts-settings)

## <a name="see-also"></a>См. также

* [Использование встроенных объектов JavaScript в сценариях Office](javascript-objects.md)
* [Использование внешних вызовов Fetch в сценариях Office](../resources/samples/external-fetch-calls.md)
* [Office Пример сценария: Graph данных уровня воды из NOAA](../resources/scenarios/noaa-data-fetch.md)
