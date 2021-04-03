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
# <a name="external-api-call-support-in-office-scripts"></a><span data-ttu-id="ea05d-103">Поддержка внешнего вызова API в сценариях Office</span><span class="sxs-lookup"><span data-stu-id="ea05d-103">External API call support in Office Scripts</span></span>

<span data-ttu-id="ea05d-104">Авторы сценариев не должны ожидать последовательного поведения при использовании [внешних API](https://developer.mozilla.org/docs/Web/API) на этапе предварительного просмотра платформы.</span><span class="sxs-lookup"><span data-stu-id="ea05d-104">Script authors shouldn't expect consistent behavior when using [external APIs](https://developer.mozilla.org/docs/Web/API) during the platform's preview phase.</span></span> <span data-ttu-id="ea05d-105">Таким образом, не полагаться на внешние API для сценариев критических сценариев.</span><span class="sxs-lookup"><span data-stu-id="ea05d-105">As such, do not rely on external APIs for critical script scenarios.</span></span>

<span data-ttu-id="ea05d-106">Вызовы к внешним API можно делать только через приложение Excel, а не через Power Automate [при нормальных обстоятельствах.](#external-calls-from-power-automate)</span><span class="sxs-lookup"><span data-stu-id="ea05d-106">Calls to external APIs can be only be made through the Excel application, not through Power Automate [under normal circumstances](#external-calls-from-power-automate).</span></span>

> [!CAUTION]
> <span data-ttu-id="ea05d-107">Внешние вызовы могут привести к воздействию конфиденциальных данных на нежелательные конечные точки.</span><span class="sxs-lookup"><span data-stu-id="ea05d-107">External calls may result in sensitive data being exposed to undesirable endpoints.</span></span> <span data-ttu-id="ea05d-108">Администратор может установить защиту брандмауэра от таких вызовов.</span><span class="sxs-lookup"><span data-stu-id="ea05d-108">Your admin can establish firewall protection against such calls.</span></span>

## <a name="working-with-fetch"></a><span data-ttu-id="ea05d-109">Работа с `fetch`</span><span class="sxs-lookup"><span data-stu-id="ea05d-109">Working with `fetch`</span></span>

<span data-ttu-id="ea05d-110">API [извлекает](https://developer.mozilla.org/docs/Web/API/Fetch_API) сведения из внешних служб.</span><span class="sxs-lookup"><span data-stu-id="ea05d-110">The [fetch API](https://developer.mozilla.org/docs/Web/API/Fetch_API) retrieves information from external services.</span></span> <span data-ttu-id="ea05d-111">Это API, поэтому необходимо настроить подпись `async` `main` скрипта.</span><span class="sxs-lookup"><span data-stu-id="ea05d-111">It is an `async` API, so you will need to adjust the `main` signature of your script.</span></span> <span data-ttu-id="ea05d-112">Сделайте `main` `async` функцию и делайте так, чтобы она возвращала `Promise<void>` .</span><span class="sxs-lookup"><span data-stu-id="ea05d-112">Make the `main` function `async` and have it return a `Promise<void>`.</span></span> <span data-ttu-id="ea05d-113">Вы также должны быть уверены `await` в `fetch` вызове и `json` ирисовке.</span><span class="sxs-lookup"><span data-stu-id="ea05d-113">You should also be sure to `await` the `fetch` call and `json` retrieval.</span></span> <span data-ttu-id="ea05d-114">Это обеспечивает завершение этих операций до завершения сценария.</span><span class="sxs-lookup"><span data-stu-id="ea05d-114">This ensures those operations complete before the script ends.</span></span>

<span data-ttu-id="ea05d-115">Следующий сценарий использует для получения данных JSON с `fetch` тестового сервера в заданном URL-адресе.</span><span class="sxs-lookup"><span data-stu-id="ea05d-115">The following script uses `fetch` to retrieve JSON data from the test server in the given URL.</span></span>

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

<span data-ttu-id="ea05d-116">Пример [сценария Office Scripts.](../resources/scenarios/noaa-data-fetch.md) На диаграмме данных уровня воды из NOAA демонстрируется команда извлекаемой информации, используемая для получения записей из базы данных "Приливы и течения" Национального управления океанических и атмосферных исследований.</span><span class="sxs-lookup"><span data-stu-id="ea05d-116">The [Office Scripts sample scenario: Graph water-level data from NOAA](../resources/scenarios/noaa-data-fetch.md) demonstrates the fetch command being used to retrieve records from the National Oceanic and Atmospheric Administration's Tides and Currents database.</span></span>

## <a name="external-calls-from-power-automate"></a><span data-ttu-id="ea05d-117">Внешние вызовы из Power Automate</span><span class="sxs-lookup"><span data-stu-id="ea05d-117">External calls from Power Automate</span></span>

<span data-ttu-id="ea05d-118">Любые внешние вызовы API сбой при запуске скрипта с power Automate.</span><span class="sxs-lookup"><span data-stu-id="ea05d-118">Any external API calls fail when a script is run with Power Automate.</span></span> <span data-ttu-id="ea05d-119">Это поведенческая разница между запуском скрипта через клиента Excel и power Automate.</span><span class="sxs-lookup"><span data-stu-id="ea05d-119">This is a behavioral difference between running a script through the Excel client and through Power Automate.</span></span> <span data-ttu-id="ea05d-120">Убедитесь в том, что перед созданием их в поток необходимо проверить сценарии для таких ссылок.</span><span class="sxs-lookup"><span data-stu-id="ea05d-120">Be sure to check your scripts for such references before building them into a flow.</span></span>

> [!WARNING]
> <span data-ttu-id="ea05d-121">Внешние вызовы, сделанные через соединитель [Excel Online](/connectors/excelonlinebusiness) Power Automate, не удается поддерживать существующие политики предотвращения потери данных.</span><span class="sxs-lookup"><span data-stu-id="ea05d-121">External calls made through the Power Automate [Excel Online connector](/connectors/excelonlinebusiness) fail in order to help uphold existing data loss prevention policies.</span></span> <span data-ttu-id="ea05d-122">Однако сценарии, которые запускаются с помощью Power Automate, делаются так за пределами организации и вне брандмауэров организации.</span><span class="sxs-lookup"><span data-stu-id="ea05d-122">However, scripts that are run through Power Automate are done so outside of your organization, and outside of your organization's firewalls.</span></span> <span data-ttu-id="ea05d-123">Для дополнительной защиты от вредоносных пользователей в этой внешней среде администратор может управлять использованием скриптов Office.</span><span class="sxs-lookup"><span data-stu-id="ea05d-123">For additional protection from malicious users in this external environment, your admin can control the use of Office Scripts.</span></span> <span data-ttu-id="ea05d-124">Администратор может отключить соединитель Excel Online в Power Automate или отключить скрипты Office для Excel в Интернете с помощью элементов управления администратором [office Scripts.](/microsoft-365/admin/manage/manage-office-scripts-settings)</span><span class="sxs-lookup"><span data-stu-id="ea05d-124">Your admin can either disable the Excel Online connector in Power Automate or turn off Office Scripts for Excel on the web through the [Office Scripts administrator controls](/microsoft-365/admin/manage/manage-office-scripts-settings).</span></span>

## <a name="see-also"></a><span data-ttu-id="ea05d-125">См. также</span><span class="sxs-lookup"><span data-stu-id="ea05d-125">See also</span></span>

- [<span data-ttu-id="ea05d-126">Использование встроенных объектов JavaScript в сценариях Office</span><span class="sxs-lookup"><span data-stu-id="ea05d-126">Using built-in JavaScript objects in Office Scripts</span></span>](javascript-objects.md)
- [<span data-ttu-id="ea05d-127">Пример сценария office Scripts: Граф данных уровня воды из NOAA</span><span class="sxs-lookup"><span data-stu-id="ea05d-127">Office Scripts sample scenario: Graph water-level data from NOAA</span></span>](../resources/scenarios/noaa-data-fetch.md)
