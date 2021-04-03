---
title: Поддержка внешнего вызова API в сценариях Office
description: Поддержка и руководство по ведению внешних вызовов API в скрипте Office.
ms.date: 01/05/2021
localization_priority: Normal
ms.openlocfilehash: 74b8750f609370370759ca4a4a1daa998363ac2e
ms.sourcegitcommit: 5d24e77df70aa2c1c982275d53213c2a9323ff86
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 04/02/2021
ms.locfileid: "51570313"
---
# <a name="external-api-call-support-in-office-scripts"></a>Поддержка внешнего вызова API в сценариях Office

Авторы сценариев не должны ожидать последовательного поведения при использовании [внешних API](https://developer.mozilla.org/docs/Web/API) на этапе предварительного просмотра платформы. Таким образом, не полагаться на внешние API для сценариев критических сценариев.

Вызовы к внешним API можно делать только через приложение Excel, а не через Power Automate [при нормальных обстоятельствах.](#external-calls-from-power-automate)

> [!CAUTION]
> Внешние вызовы могут привести к воздействию конфиденциальных данных на нежелательные конечные точки. Администратор может установить защиту брандмауэра от таких вызовов.

## <a name="working-with-fetch"></a>Работа с `fetch`

API [извлекает](https://developer.mozilla.org/docs/Web/API/Fetch_API) сведения из внешних служб. Это API, поэтому необходимо настроить подпись `async` `main` скрипта. Сделайте `main` `async` функцию и делайте так, чтобы она возвращала `Promise<void>` . Вы также должны быть уверены `await` в `fetch` вызове и `json` ирисовке. Это обеспечивает завершение этих операций до завершения сценария.

Следующий сценарий использует для получения данных JSON с `fetch` тестового сервера в заданном URL-адресе.

```TypeScript
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

Пример [сценария Office Scripts.](../resources/scenarios/noaa-data-fetch.md) На диаграмме данных уровня воды из NOAA демонстрируется команда извлекаемой информации, используемая для получения записей из базы данных "Приливы и течения" Национального управления океанических и атмосферных исследований.

## <a name="external-calls-from-power-automate"></a>Внешние вызовы из Power Automate

Любые внешние вызовы API сбой при запуске скрипта с power Automate. Это поведенческая разница между запуском скрипта через клиента Excel и power Automate. Убедитесь в том, что перед созданием их в поток необходимо проверить сценарии для таких ссылок.

> [!WARNING]
> Внешние вызовы, сделанные через соединитель [Excel Online](/connectors/excelonlinebusiness) Power Automate, не удается поддерживать существующие политики предотвращения потери данных. Однако сценарии, которые запускаются с помощью Power Automate, делаются так за пределами организации и вне брандмауэров организации. Для дополнительной защиты от вредоносных пользователей в этой внешней среде администратор может управлять использованием скриптов Office. Администратор может отключить соединитель Excel Online в Power Automate или отключить скрипты Office для Excel в Интернете с помощью элементов управления администратором [office Scripts.](/microsoft-365/admin/manage/manage-office-scripts-settings)

## <a name="see-also"></a>См. также

- [Использование встроенных объектов JavaScript в сценариях Office](javascript-objects.md)
- [Пример сценария office Scripts: Граф данных уровня воды из NOAA](../resources/scenarios/noaa-data-fetch.md)
