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
# <a name="external-api-call-support-in-office-scripts"></a><span data-ttu-id="0f182-103">Поддержка внешнего вызова API в сценариях Office</span><span class="sxs-lookup"><span data-stu-id="0f182-103">External API call support in Office Scripts</span></span>

<span data-ttu-id="0f182-104">Авторы сценариев не должны ожидать согласованного поведения при использовании внешних [API](https://developer.mozilla.org/docs/Web/API) на этапе предварительного просмотра платформы.</span><span class="sxs-lookup"><span data-stu-id="0f182-104">Script authors shouldn't expect consistent behavior when using [external APIs](https://developer.mozilla.org/docs/Web/API) during the platform's preview phase.</span></span> <span data-ttu-id="0f182-105">Таким образом, не полагайтесь на внешние API для критически важных сценариев.</span><span class="sxs-lookup"><span data-stu-id="0f182-105">As such, do not rely on external APIs for critical script scenarios.</span></span>

<span data-ttu-id="0f182-106">Вызовы внешних API можно делать только через приложение Excel, а не с помощью Power Automate [в обычных условиях.](#external-calls-from-power-automate)</span><span class="sxs-lookup"><span data-stu-id="0f182-106">Calls to external APIs can be only be made through the Excel application, not through Power Automate [under normal circumstances](#external-calls-from-power-automate).</span></span>

> [!CAUTION]
> <span data-ttu-id="0f182-107">Внешние вызовы могут привести к передаче конфиденциальных данных нежелательным конечным точкам.</span><span class="sxs-lookup"><span data-stu-id="0f182-107">External calls may result in sensitive data being exposed to undesirable endpoints.</span></span> <span data-ttu-id="0f182-108">Администратор может установить защиту брандмауэра от таких вызовов.</span><span class="sxs-lookup"><span data-stu-id="0f182-108">Your admin can establish firewall protection against such calls.</span></span>

## <a name="working-with-fetch"></a><span data-ttu-id="0f182-109">Работа с `fetch`</span><span class="sxs-lookup"><span data-stu-id="0f182-109">Working with `fetch`</span></span>

<span data-ttu-id="0f182-110">API [получения извлекает](https://developer.mozilla.org/docs/Web/API/Fetch_API) сведения из внешних служб.</span><span class="sxs-lookup"><span data-stu-id="0f182-110">The [fetch API](https://developer.mozilla.org/docs/Web/API/Fetch_API) retrieves information from external services.</span></span> <span data-ttu-id="0f182-111">Это API, поэтому вам потребуется настроить подпись `async` `main` скрипта.</span><span class="sxs-lookup"><span data-stu-id="0f182-111">It is an `async` API, so you will need to adjust the `main` signature of your script.</span></span> <span data-ttu-id="0f182-112">Сделайте `main` `async` функцию и делайте так, чтобы она возвращала `Promise<void>` .</span><span class="sxs-lookup"><span data-stu-id="0f182-112">Make the `main` function `async` and have it return a `Promise<void>`.</span></span> <span data-ttu-id="0f182-113">Кроме того, следует убедиться в `await` `fetch` вызове и `json` иных вызовах.</span><span class="sxs-lookup"><span data-stu-id="0f182-113">You should also be sure to `await` the `fetch` call and `json` retrieval.</span></span> <span data-ttu-id="0f182-114">Это гарантирует, что эти операции будут завершены до завершения скрипта.</span><span class="sxs-lookup"><span data-stu-id="0f182-114">This ensures those operations complete before the script ends.</span></span>

<span data-ttu-id="0f182-115">Следующий сценарий использует `fetch` для получения данных JSON с тестового сервера по заданном URL-адресу.</span><span class="sxs-lookup"><span data-stu-id="0f182-115">The following script uses `fetch` to retrieve JSON data from the test server in the given URL.</span></span>

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

<span data-ttu-id="0f182-116">Пример [сценария сценариев Office:](../resources/scenarios/noaa-data-fetch.md) данные на уровне ватерли Graph из NOAA демонстрируют команду получения, используемую для извлечения записей из базы данных "Посцены и текущие данные" национального правительства.</span><span class="sxs-lookup"><span data-stu-id="0f182-116">The [Office Scripts sample scenario: Graph water-level data from NOAA](../resources/scenarios/noaa-data-fetch.md) demonstrates the fetch command being used to retrieve records from the National Oceanic and Atmospheric Administration's Tides and Currents database.</span></span>

## <a name="external-calls-from-power-automate"></a><span data-ttu-id="0f182-117">Внешние вызовы из Power Automate</span><span class="sxs-lookup"><span data-stu-id="0f182-117">External calls from Power Automate</span></span>

<span data-ttu-id="0f182-118">При запуске сценария с помощью Power Automate все внешние вызовы API не будут работать.</span><span class="sxs-lookup"><span data-stu-id="0f182-118">Any external API calls fail when a script is run with Power Automate.</span></span> <span data-ttu-id="0f182-119">Это различие в поведении между запуском сценария через клиент Excel и с помощью Power Automate.</span><span class="sxs-lookup"><span data-stu-id="0f182-119">This is a behavioral difference between running a script through the Excel client and through Power Automate.</span></span> <span data-ttu-id="0f182-120">Обязательно проверяйте такие ссылки в скриптах перед их созданием в потоке.</span><span class="sxs-lookup"><span data-stu-id="0f182-120">Be sure to check your scripts for such references before building them into a flow.</span></span>

> [!WARNING]
> <span data-ttu-id="0f182-121">Внешние вызовы, сделанные через соединитель Power Automate [Excel Online,](/connectors/excelonlinebusiness) не поддерживают существующие политики защиты от потери данных.</span><span class="sxs-lookup"><span data-stu-id="0f182-121">External calls made through the Power Automate [Excel Online connector](/connectors/excelonlinebusiness) fail in order to help uphold existing data loss prevention policies.</span></span> <span data-ttu-id="0f182-122">Однако сценарии, которые запускаются с помощью Power Automate, делают это за пределами организации и за пределами брандмауэров организации.</span><span class="sxs-lookup"><span data-stu-id="0f182-122">However, scripts that are run through Power Automate are done so outside of your organization, and outside of your organization's firewalls.</span></span> <span data-ttu-id="0f182-123">Для дополнительной защиты от злоумышленников во внешней среде администратор может управлять использованием сценариев Office.</span><span class="sxs-lookup"><span data-stu-id="0f182-123">For additional protection from malicious users in this external environment, your admin can control the use of Office Scripts.</span></span> <span data-ttu-id="0f182-124">Администратор может отключить соединитель Excel Online в Power Automate или отключить скрипты Office для Excel в Интернете с помощью элементов управления администратора [сценариев Office.](/microsoft-365/admin/manage/manage-office-scripts-settings)</span><span class="sxs-lookup"><span data-stu-id="0f182-124">Your admin can either disable the Excel Online connector in Power Automate or turn off Office Scripts for Excel on the web through the [Office Scripts administrator controls](/microsoft-365/admin/manage/manage-office-scripts-settings).</span></span>

## <a name="see-also"></a><span data-ttu-id="0f182-125">См. также</span><span class="sxs-lookup"><span data-stu-id="0f182-125">See also</span></span>

- [<span data-ttu-id="0f182-126">Использование встроенных объектов JavaScript в сценариях Office</span><span class="sxs-lookup"><span data-stu-id="0f182-126">Using built-in JavaScript objects in Office Scripts</span></span>](javascript-objects.md)
- [<span data-ttu-id="0f182-127">Пример сценария сценариев Office: данные на уровне ватерли Graph из NOAA</span><span class="sxs-lookup"><span data-stu-id="0f182-127">Office Scripts sample scenario: Graph water-level data from NOAA</span></span>](../resources/scenarios/noaa-data-fetch.md)
