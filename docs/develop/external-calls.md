---
title: Поддержка внешнего вызова API в сценариях Office
description: Поддержка и руководство по выполнению внешних вызовов API в Office скрипта.
ms.date: 06/10/2022
ms.localizationpriority: medium
ms.openlocfilehash: 472b2e1b4aa38366b68b573fa959deee616b9dbe
ms.sourcegitcommit: aecbd5baf1e2122d836c3eef3b15649e132bc68e
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 06/16/2022
ms.locfileid: "66128225"
---
# <a name="external-api-call-support-in-office-scripts"></a>Поддержка внешнего вызова API в сценариях Office

Скрипты поддерживают вызовы внешних служб. Используйте эти службы для предоставления данных и других сведений в книгу.

> [!CAUTION]
> Внешние вызовы могут привести к раскрытию конфиденциальных данных нежелательным конечным точкам. Администратор может установить защиту брандмауэра от таких вызовов.

> [!IMPORTANT]
> Вызовы внешних API могут выполняться только через Excel, а не через Power Automate в [обычных условиях](#external-calls-from-power-automate). Внешние вызовы также не поддерживаются для скриптов, хранящихся на SharePoint сайте.

## <a name="configure-your-script-for-external-calls"></a>Настройка скрипта для внешних вызовов

Внешние вызовы [являются асинхронными](https://developer.mozilla.org/docs/Learn/JavaScript/Asynchronous/Async_await) и требуют, чтобы сценарий был помечен как `async`. Добавьте префикс `async` в `main` функцию и ведите его возврат `Promise`, как показано ниже:

```typescript
async function main(workbook: ExcelScript.Workbook) : Promise <void>
```

> [!NOTE]
> Скрипты, возвращаемые другими сведениями, могут возвращать объект `Promise` этого типа. Например, если скрипту необходимо вернуть объект `Employee` , то возвращаемая сигнатура будет выглядеть следующим образом: `: Promise <Employee>`

Для выполнения вызовов к этой службе необходимо изучить интерфейсы внешней службы. При использовании или `fetch` [REST API](https://wikipedia.org/wiki/Representational_state_transfer) необходимо определить структуру JSON возвращаемых данных. Для входных и выходных данных скрипта рассмотрите `interface` возможность создания в соответствии с необходимыми структурами JSON. Это обеспечивает более безопасную тип скрипта. Пример этого можно увидеть в разделе "[Использование выборки из Office скриптов"](../resources/samples/external-fetch-calls.md).

### <a name="limitations-with-external-calls-from-office-scripts"></a>Ограничения внешних вызовов из Office скриптов

* Вход или использование потоков проверки подлинности OAuth2 не поддерживается. Все ключи и учетные данные должны быть жестко заданы (или считываться из другого источника).
* Нет инфраструктуры для хранения учетных данных и ключей API. Этим должен управлять пользователь.
* Файлы cookie документов `localStorage`и объекты `sessionStorage` не поддерживаются.
* Внешние вызовы могут привести к раскрытию конфиденциальных данных нежелательным конечным точкам или к переносу внешних данных во внутренние книги. Администратор может установить защиту брандмауэра от таких вызовов. Не забудьте проверить локальные политики, прежде чем полагаться на внешние вызовы.
* Перед получением зависимости обязательно проверьте объем пропускной способности данных. Например, извлечение всего внешнего набора данных может оказаться не лучшим вариантом. Вместо этого для получения данных фрагментами следует использовать разбиение на страницы.

## <a name="retrieve-information-with-fetch"></a>Получение сведений с помощью `fetch`

[API fetch извлекает](https://developer.mozilla.org/docs/Web/API/Fetch_API) сведения из внешних служб. Это API `async` , поэтому необходимо `main` настроить подпись скрипта. Сделайте функцию `main` `async`. Вы также должны быть уверены в вызове `await` `fetch` и `json` извлечении. Это гарантирует, что эти операции будут завершены до завершения скрипта.

Все данные JSON, полученные путем `fetch` , должны соответствовать интерфейсу, определенному в скрипте. Возвращаемое значение должно быть назначено определенному типу, Office скрипты [не поддерживают `any` этот тип](typescript-restrictions.md#no-any-type-in-office-scripts). Сведения об именах и типах возвращаемых свойств см. в документации по службе. Затем добавьте соответствующий интерфейс или интерфейсы в скрипт.

Следующий сценарий использует для `fetch` получения данных JSON с тестового сервера по указанному URL-адресу. Обратите внимание `JSONData` на интерфейс для хранения данных в виде соответствующего типа.

```TypeScript
async function main(workbook: ExcelScript.Workbook) {
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

* В примере использования внешних выборок [в Office scripts](../resources/samples/external-fetch-calls.md) показано, как получить основные сведения о репозиториях GitHub пользователя.
* Пример [сценария Office](../resources/scenarios/noaa-data-fetch.md) сценариев: Graph данных на уровне воды из NOAA демонстрирует команду fetch, используемую для извлечения записей из базы данных Tides and Currents национального управления по океании и океании.

## <a name="external-calls-from-power-automate"></a>Внешние вызовы из Power Automate

Любой внешний вызов API завершается сбоем при выполнении скрипта с Power Automate. Это различие в поведении между выполнением скрипта через приложение Excel и Power Automate. Не забудьте проверить скрипты на наличие таких ссылок, прежде чем создавать их в потоке.

Для извлечения или отправки данных во внешнюю службу [необходимо использовать протокол HTTP Azure AD](/connectors/webcontents/) или другие эквивалентные действия.

> [!WARNING]
> Внешние вызовы, выполненные через соединитель Power Automate [Excel Online](/connectors/excelonlinebusiness), завершались сбоем, чтобы помочь сохранить существующие политики защиты от потери данных. Однако скрипты, выполняемые через Power Automate, выполняются за пределами вашей организации и за пределами брандмауэров вашей организации. Для дополнительной защиты от вредоносных пользователей во внешней среде администратор может управлять использованием Office сценариев. Администратор может отключить соединитель Excel Online в Power Automate или отключить Office скрипты для Excel в Интернете с помощью элементов управления Office [scripts](/microsoft-365/admin/manage/manage-office-scripts-settings).

## <a name="see-also"></a>См. также

* [Использование JSON для передачи данных в скрипты Office и из них](use-json.md)
* [Использование встроенных объектов JavaScript в сценариях Office](javascript-objects.md)
* [Использование внешних вызовов Fetch в сценариях Office](../resources/samples/external-fetch-calls.md)
* [Office сценария сценариев: Graph данных на уровне воды из NOAA](../resources/scenarios/noaa-data-fetch.md)
