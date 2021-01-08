---
title: Поддержка внешнего вызова API в сценариях Office
description: Поддержка и руководство по внешним вызовам API в сценарии Office.
ms.date: 01/05/2021
localization_priority: Normal
ms.openlocfilehash: 1091031bc2e12f3e1e79b177c69874ee4ce61dd8
ms.sourcegitcommit: 30c4b731dc8d18fca5aa74ce59e18a4a63eb4ffc
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 01/08/2021
ms.locfileid: "49784146"
---
# <a name="external-api-call-support-in-office-scripts"></a>Поддержка внешнего вызова API в сценариях Office

Авторы сценариев не должны ожидать согласованного поведения при использовании внешних [API](https://developer.mozilla.org/docs/Web/API) на этапе предварительного просмотра платформы. Таким образом, не полагайтесь на внешние API для критически важных сценариев.

Вызовы внешних API можно делать только через приложение Excel, а не с помощью Power Automate [в обычных условиях.](#external-calls-from-power-automate)

> [!CAUTION]
> Внешние вызовы могут привести к передаче конфиденциальных данных нежелательным конечным точкам. Администратор может установить защиту брандмауэра от таких вызовов.

## <a name="working-with-fetch"></a>Работа с `fetch`

API [получения извлекает](https://developer.mozilla.org/docs/Web/API/Fetch_API) сведения из внешних служб. Это API, поэтому вам потребуется настроить подпись `async` `main` скрипта. Сделайте `main` `async` функцию и делайте так, чтобы она возвращала `Promise<void>` . Кроме того, следует убедиться в `await` `fetch` вызове и `json` иных вызовах. Это гарантирует, что эти операции будут завершены до завершения скрипта.

Следующий сценарий использует `fetch` для получения данных JSON с тестового сервера по заданном URL-адресу.

```typescript
async function main(workbook: ExcelScript.Workbook): Promise <void> {
  /* 
   * Retrieve JSON data from a test server.
   */
  let fetchResult = await fetch('https://jsonplaceholder.typicode.com/todos/1');
  let json = await fetchResult.json();

  // Displays the content from https://jsonplaceholder.typicode.com/todos/1
  console.log(JSON.stringify(json));
}
```

Пример [сценария сценариев Office:](../resources/scenarios/noaa-data-fetch.md) данные на уровне ватерли Graph из NOAA демонстрируют команду получения, используемую для извлечения записей из базы данных "Посцены и текущие данные" национального правительства.

## <a name="external-calls-from-power-automate"></a>Внешние вызовы из Power Automate

При запуске сценария с помощью Power Automate все внешние вызовы API не будут работать. Это различие в поведении между запуском сценария через клиент Excel и с помощью Power Automate. Обязательно проверяйте такие ссылки в скриптах перед их созданием в потоке.

> [!WARNING]
> Внешние вызовы, сделанные через соединитель Power Automate [Excel Online,](/connectors/excelonlinebusiness) не поддерживают существующие политики защиты от потери данных. Однако сценарии, которые запускаются с помощью Power Automate, делают это за пределами организации и за пределами брандмауэров организации. Для дополнительной защиты от злоумышленников во внешней среде администратор может управлять использованием сценариев Office. Администратор может отключить соединитель Excel Online в Power Automate или отключить скрипты Office для Excel в Интернете с помощью элементов управления администратора [сценариев Office.](/microsoft-365/admin/manage/manage-office-scripts-settings)

## <a name="see-also"></a>См. также

- [Использование встроенных объектов JavaScript в сценариях Office](javascript-objects.md)
- [Пример сценария сценариев Office: данные на уровне ватерли Graph из NOAA](../resources/scenarios/noaa-data-fetch.md)
