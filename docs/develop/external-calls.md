---
title: Поддержка внешнего вызова API в сценариях Office
description: Поддержка и руководство для внешних вызовов API в Office скрипте.
ms.date: 05/17/2021
localization_priority: Normal
ms.openlocfilehash: fd6ba0c57bf4cabb2d07421355cacff373f6706c
ms.sourcegitcommit: 4687693f02fc90a57ba30c461f35046e02e6f5fb
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 05/19/2021
ms.locfileid: "52545084"
---
# <a name="external-api-call-support-in-office-scripts"></a>Поддержка внешнего вызова API в сценариях Office

Авторы скриптов не должны ожидать последовательного поведения при использовании [внешних API](https://developer.mozilla.org/docs/Web/API) на этапе предварительного просмотра платформы. Таким образом, не полагайтесь на внешние API для критических сценариев сценариев.

Вызовы на внешние API могут быть сделаны только через Excel, а не через Power Automate при [нормальных обстоятельствах.](#external-calls-from-power-automate)

> [!CAUTION]
> Внешние вызовы могут привести к тому, что конфиденциальные данные будут подвержены нежелательным конечных точкам. Администратор может установить защиту брандмауэра от таких вызовов.

## <a name="configure-your-script-for-external-calls"></a>Настройка скрипта для внешних вызовов

Внешние [вызовы являются асинхронными](https://developer.mozilla.org/docs/Learn/JavaScript/Asynchronous/Async_await) и требуют, чтобы ваш скрипт помечен как `async` . Добавьте `async` приставку к вашей `main` функции и верните `Promise` ее, как показано здесь:

```typescript
async function main(workbook: ExcelScript.Workbook) : Promise <void>
```

> [!NOTE]
> Скрипты, возвращают другую информацию, `Promise` могут вернуть этот тип. Например, если скрипту необходимо вернуть `Employee` объект, ответная подпись будет `: Promise <Employee>`

Вам нужно будет изучить интерфейсы внешних служб, чтобы звонить в эту службу. Если вы `fetch` используете [API или REST](https://wikipedia.org/wiki/Representational_state_transfer)API, вам необходимо определить структуру JSON возвращенных данных. Для ввода и вывода из скрипта, рассмотрите возможность создания в `interface` качестве меры, необходимой для создания необходимых структур JSON. Это дает скрипту больше безопасности типа. Вы можете увидеть пример этого в [использовании извлечения из Office скриптов](../resources/samples/external-fetch-calls.md).

### <a name="limitations-with-external-calls-from-office-scripts"></a>Ограничения с внешними вызовами из Office скриптов

* Нет никакого способа войти или использовать потоки OAuth2 типа аутентификации. Все ключи и учетные данные должны быть жестко закодированы (или читать из другого источника).
* Нет инфраструктуры для хранения учетных данных и ключей API. Этим должен управлять пользователь.
* Файлы `localStorage` cookie-файлов и объекты не `sessionStorage` поддерживаются. 
* Внешние вызовы могут привести к тому, что конфиденциальные данные будут подвержены нежелательным конечных точкам или внешние данные будут завесят во внутренние рабочие книжки. Администратор может установить защиту брандмауэра от таких вызовов. Не забудьте проверить с местными политиками, прежде чем полагаться на внешние вызовы.
* Не забудьте проверить объем пропускной способности данных до принятия зависимости. Например, стягивание всего внешнего набора данных может быть не лучшим вариантом, а вместо этого pagination следует использовать для получения данных в кусках.

## <a name="retrieve-information-with-fetch"></a>Получение информации с помощью `fetch`

API [получает информацию](https://developer.mozilla.org/docs/Web/API/Fetch_API) из внешних служб. Это `async` API, поэтому вам нужно настроить `main` подпись вашего скрипта. Сделать `main` `async` функцию и заставить его вернуть `Promise<void>` . Вы также должны быть уверены `await` в `fetch` вызове `json` и поиске. Это гарантирует, что эти операции будут завершены до окончания сценария.

Любые данные JSON, полученные с `fetch` помощью, должны соответствовать интерфейсу, определенному в скрипте. Возвращенное значение должно быть назначено определенному [типу, Office скрипты не поддерживают `any` тип.](typescript-restrictions.md#no-any-type-in-office-scripts) Вы должны обратиться к документации для вашего сервиса, чтобы увидеть, какие имена и типы возвращенных свойств. Затем добавьте соответствующий интерфейс или интерфейсы в свой скрипт.

Следующий скрипт используется `fetch` для извлечения данных JSON с тестового сервера в данном URL. Обратите внимание `JSONData` на интерфейс для хранения данных в качестве подходящего типа.

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

### <a name="other-fetch-samples"></a>Другие `fetch` образцы

* Внешние [вызовы извлечения в Office скриптов](../resources/samples/external-fetch-calls.md) показывает, как получить основную информацию о GitHub пользователя.
* Сценарий [Office Scripts: Graph данные об уровне](../resources/scenarios/noaa-data-fetch.md) воды от NOAA демонстрируют команду извлечения, используемую для извлечения записей из базы данных Приливов и течений Национального управления океанических и атмосферных исследований.

## <a name="external-calls-from-power-automate"></a>Внешние звонки из Power Automate

Любой внешний вызов API выходит из строя при запуске скрипта с Power Automate. Это поведенческая разница между запуском скрипта через приложение Excel и Power Automate. Не забудьте проверить ваши сценарии для таких ссылок, прежде чем строить их в поток.

Вам придется использовать HTTP с [Azure AD или другими эквивалентными](/connectors/webcontents/) действиями, чтобы вытащить данные из или перейти на внешнюю службу.

> [!WARNING]
> Внешние вызовы, сделанные через Power Automate [Excel Online, не удается,](/connectors/excelonlinebusiness) чтобы помочь поддерживать существующие политики предотвращения потери данных. Однако сценарии, которые проходят через Power Automate, сделаны это за пределами вашей организации и за пределами брандмауэров организации. Для дополнительной защиты от вредоносных пользователей в этой внешней среде администратор может контролировать использование Office скриптов. Администратор может либо отключить разъем Excel Online в Power Automate или выключить Office скрипты для Excel в Интернете с [помощью Office управления скриптами.](/microsoft-365/admin/manage/manage-office-scripts-settings)

## <a name="see-also"></a>См. также

* [Использование встроенных объектов JavaScript в сценариях Office](javascript-objects.md)
* [Использование внешних вызовов Fetch в сценариях Office](../resources/samples/external-fetch-calls.md)
* [Office Сценарий выборки сценария: Graph данные уровня воды от NOAA](../resources/scenarios/noaa-data-fetch.md)
