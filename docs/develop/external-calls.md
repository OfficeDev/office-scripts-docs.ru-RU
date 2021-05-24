---
title: Поддержка внешнего вызова API в сценариях Office
description: Поддержка и руководство по принятию внешних вызовов API в Office скрипта.
ms.date: 05/17/2021
localization_priority: Normal
ms.openlocfilehash: fd6ba0c57bf4cabb2d07421355cacff373f6706c
ms.sourcegitcommit: 4687693f02fc90a57ba30c461f35046e02e6f5fb
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 05/19/2021
ms.locfileid: "52545084"
---
# <a name="external-api-call-support-in-office-scripts"></a><span data-ttu-id="3cc4b-103">Поддержка внешнего вызова API в сценариях Office</span><span class="sxs-lookup"><span data-stu-id="3cc4b-103">External API call support in Office Scripts</span></span>

<span data-ttu-id="3cc4b-104">Авторы сценариев не должны ожидать последовательного поведения при использовании [внешних API](https://developer.mozilla.org/docs/Web/API) на этапе предварительного просмотра платформы.</span><span class="sxs-lookup"><span data-stu-id="3cc4b-104">Script authors shouldn't expect consistent behavior when using [external APIs](https://developer.mozilla.org/docs/Web/API) during the platform's preview phase.</span></span> <span data-ttu-id="3cc4b-105">Таким образом, не полагаться на внешние API для сценариев критических сценариев.</span><span class="sxs-lookup"><span data-stu-id="3cc4b-105">As such, do not rely on external APIs for critical script scenarios.</span></span>

<span data-ttu-id="3cc4b-106">Вызовы к внешним API можно сделать только через Excel, а не Power Automate при [обычных обстоятельствах.](#external-calls-from-power-automate)</span><span class="sxs-lookup"><span data-stu-id="3cc4b-106">Calls to external APIs can be only be made through the Excel application, not through Power Automate [under normal circumstances](#external-calls-from-power-automate).</span></span>

> [!CAUTION]
> <span data-ttu-id="3cc4b-107">Внешние вызовы могут привести к воздействию конфиденциальных данных на нежелательные конечные точки.</span><span class="sxs-lookup"><span data-stu-id="3cc4b-107">External calls may result in sensitive data being exposed to undesirable endpoints.</span></span> <span data-ttu-id="3cc4b-108">Администратор может установить защиту брандмауэра от таких вызовов.</span><span class="sxs-lookup"><span data-stu-id="3cc4b-108">Your admin can establish firewall protection against such calls.</span></span>

## <a name="configure-your-script-for-external-calls"></a><span data-ttu-id="3cc4b-109">Настройка сценария для внешних вызовов</span><span class="sxs-lookup"><span data-stu-id="3cc4b-109">Configure your script for external calls</span></span>

<span data-ttu-id="3cc4b-110">Внешние вызовы [асинхронны](https://developer.mozilla.org/docs/Learn/JavaScript/Asynchronous/Async_await) и требуют, чтобы сценарий был помечен как `async` .</span><span class="sxs-lookup"><span data-stu-id="3cc4b-110">External calls are [asynchronous](https://developer.mozilla.org/docs/Learn/JavaScript/Asynchronous/Async_await) and require that your script is marked as `async`.</span></span> <span data-ttu-id="3cc4b-111">Добавьте `async` префикс в функцию и вернетесь, `main` `Promise` как показано здесь:</span><span class="sxs-lookup"><span data-stu-id="3cc4b-111">Add the `async` prefix to your `main` function and have it return a `Promise`, as shown here:</span></span>

```typescript
async function main(workbook: ExcelScript.Workbook) : Promise <void>
```

> [!NOTE]
> <span data-ttu-id="3cc4b-112">Скрипты, возвращая другие сведения, могут возвращать `Promise` один из этих типов.</span><span class="sxs-lookup"><span data-stu-id="3cc4b-112">Scripts that return other information can return a `Promise` of that type.</span></span> <span data-ttu-id="3cc4b-113">Например, если сценарию необходимо вернуть `Employee` объект, возвращаемая подпись будет `: Promise <Employee>`</span><span class="sxs-lookup"><span data-stu-id="3cc4b-113">For example, if your script needs to return an `Employee` object, the return signature would be `: Promise <Employee>`</span></span>

<span data-ttu-id="3cc4b-114">Чтобы звонить в эту службу, необходимо изучить интерфейсы внешней службы.</span><span class="sxs-lookup"><span data-stu-id="3cc4b-114">You'll need to learn the external service's interfaces to make calls to that service.</span></span> <span data-ttu-id="3cc4b-115">Если вы используете API rest или REST, вам необходимо определить структуру `fetch` JSON возвращаемого данных. [](https://wikipedia.org/wiki/Representational_state_transfer)</span><span class="sxs-lookup"><span data-stu-id="3cc4b-115">If you are using `fetch` or [REST APIs](https://wikipedia.org/wiki/Representational_state_transfer), you need to determine the JSON structure of the returned data.</span></span> <span data-ttu-id="3cc4b-116">Для ввода и вывода из скрипта рассмотрите возможность создания, чтобы соответствовать необходимым `interface` структурам JSON.</span><span class="sxs-lookup"><span data-stu-id="3cc4b-116">For both input to and output from your script, consider making an `interface` to match the needed JSON structures.</span></span> <span data-ttu-id="3cc4b-117">Это обеспечивает скрипту больше безопасности типа.</span><span class="sxs-lookup"><span data-stu-id="3cc4b-117">This gives the script more type safety.</span></span> <span data-ttu-id="3cc4b-118">Пример этого см. в примере [Using fetch from Office Scripts.](../resources/samples/external-fetch-calls.md)</span><span class="sxs-lookup"><span data-stu-id="3cc4b-118">You can see an example of this in [Using fetch from Office Scripts](../resources/samples/external-fetch-calls.md).</span></span>

### <a name="limitations-with-external-calls-from-office-scripts"></a><span data-ttu-id="3cc4b-119">Ограничения внешних вызовов из Office скриптов</span><span class="sxs-lookup"><span data-stu-id="3cc4b-119">Limitations with external calls from Office Scripts</span></span>

* <span data-ttu-id="3cc4b-120">Нет способа войти или использовать потоки проверки подлинности OAuth2.</span><span class="sxs-lookup"><span data-stu-id="3cc4b-120">There is no way to sign in or use OAuth2 type of authentication flows.</span></span> <span data-ttu-id="3cc4b-121">Все ключи и учетные данные должны быть жестко закодированы (или считываться из другого источника).</span><span class="sxs-lookup"><span data-stu-id="3cc4b-121">All keys and credentials have to be hardcoded (or read from another source).</span></span>
* <span data-ttu-id="3cc4b-122">Нет инфраструктуры для хранения учетных данных и ключей API.</span><span class="sxs-lookup"><span data-stu-id="3cc4b-122">There is no infrastructure to store API credentials and keys.</span></span> <span data-ttu-id="3cc4b-123">Этим должен управлять пользователь.</span><span class="sxs-lookup"><span data-stu-id="3cc4b-123">This will have to be managed by the user.</span></span>
* <span data-ttu-id="3cc4b-124">Файлы cookie документов `localStorage` и `sessionStorage` объекты не поддерживаются.</span><span class="sxs-lookup"><span data-stu-id="3cc4b-124">Document cookies, `localStorage`, and `sessionStorage` objects are not supported.</span></span> 
* <span data-ttu-id="3cc4b-125">Внешние вызовы могут привести к воздействию конфиденциальных данных на нежелательные конечные точки или внешним данным, которые будут занесены во внутренние книги.</span><span class="sxs-lookup"><span data-stu-id="3cc4b-125">External calls may result in sensitive data being exposed to undesirable endpoints, or external data to be brought into internal workbooks.</span></span> <span data-ttu-id="3cc4b-126">Администратор может установить защиту брандмауэра от таких вызовов.</span><span class="sxs-lookup"><span data-stu-id="3cc4b-126">Your admin can establish firewall protection against such calls.</span></span> <span data-ttu-id="3cc4b-127">Не забудьте проверить местные политики, прежде чем полагаться на внешние вызовы.</span><span class="sxs-lookup"><span data-stu-id="3cc4b-127">Be sure to check with local policies prior to relying on external calls.</span></span>
* <span data-ttu-id="3cc4b-128">Убедитесь, что перед принятием зависимости необходимо проверить объем пропускной способности данных.</span><span class="sxs-lookup"><span data-stu-id="3cc4b-128">Be sure to check the amount of data throughput prior to taking a dependency.</span></span> <span data-ttu-id="3cc4b-129">Например, стягивание всего внешнего наборов данных может оказаться не самым лучшим вариантом, а вместо этого следует использовать pagination для получения данных в куски.</span><span class="sxs-lookup"><span data-stu-id="3cc4b-129">For instance, pulling down the entire external dataset may not be the best option and instead pagination should be used to get data in chunks.</span></span>

## <a name="retrieve-information-with-fetch"></a><span data-ttu-id="3cc4b-130">Извлечение сведений с помощью `fetch`</span><span class="sxs-lookup"><span data-stu-id="3cc4b-130">Retrieve information with `fetch`</span></span>

<span data-ttu-id="3cc4b-131">API [извлекает](https://developer.mozilla.org/docs/Web/API/Fetch_API) сведения из внешних служб.</span><span class="sxs-lookup"><span data-stu-id="3cc4b-131">The [fetch API](https://developer.mozilla.org/docs/Web/API/Fetch_API) retrieves information from external services.</span></span> <span data-ttu-id="3cc4b-132">Это `async` API, поэтому необходимо настроить подпись `main` скрипта.</span><span class="sxs-lookup"><span data-stu-id="3cc4b-132">It is an `async` API, so you need to adjust the `main` signature of your script.</span></span> <span data-ttu-id="3cc4b-133">Сделайте `main` `async` функцию и делайте так, чтобы она возвращала `Promise<void>` .</span><span class="sxs-lookup"><span data-stu-id="3cc4b-133">Make the `main` function `async` and have it return a `Promise<void>`.</span></span> <span data-ttu-id="3cc4b-134">Вы также должны быть уверены `await` в `fetch` вызове и `json` ирисовке.</span><span class="sxs-lookup"><span data-stu-id="3cc4b-134">You should also be sure to `await` the `fetch` call and `json` retrieval.</span></span> <span data-ttu-id="3cc4b-135">Это обеспечивает завершение этих операций до завершения сценария.</span><span class="sxs-lookup"><span data-stu-id="3cc4b-135">This ensures those operations complete before the script ends.</span></span>

<span data-ttu-id="3cc4b-136">Все полученные JSON данные `fetch` должны соответствовать интерфейсу, определенному в сценарии.</span><span class="sxs-lookup"><span data-stu-id="3cc4b-136">Any JSON data retrieved by `fetch` must match an interface defined in the script.</span></span> <span data-ttu-id="3cc4b-137">Возвращенное значение должно быть назначено определенному типу, так как Office скрипты [не поддерживают `any` тип](typescript-restrictions.md#no-any-type-in-office-scripts).</span><span class="sxs-lookup"><span data-stu-id="3cc4b-137">The returned value must be assigned to a specific type because [Office Scripts do not support the `any` type](typescript-restrictions.md#no-any-type-in-office-scripts).</span></span> <span data-ttu-id="3cc4b-138">Чтобы узнать имена и типы возвращаемого свойства, необходимо обратиться к документации для службы.</span><span class="sxs-lookup"><span data-stu-id="3cc4b-138">You should refer to the documentation for your service to see what the names and types of the returned properties are.</span></span> <span data-ttu-id="3cc4b-139">Затем добавьте в скрипт совпадающий интерфейс или интерфейс.</span><span class="sxs-lookup"><span data-stu-id="3cc4b-139">Then, add the matching interface or interfaces to your script.</span></span>

<span data-ttu-id="3cc4b-140">Следующий сценарий использует для получения данных JSON с `fetch` тестового сервера в заданном URL-адресе.</span><span class="sxs-lookup"><span data-stu-id="3cc4b-140">The following script uses `fetch` to retrieve JSON data from the test server in the given URL.</span></span> <span data-ttu-id="3cc4b-141">Обратите внимание `JSONData` на интерфейс для хранения данных в качестве типа совпадения.</span><span class="sxs-lookup"><span data-stu-id="3cc4b-141">Note the `JSONData` interface to store the data as a matching type.</span></span>

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

### <a name="other-fetch-samples"></a><span data-ttu-id="3cc4b-142">Другие `fetch` примеры</span><span class="sxs-lookup"><span data-stu-id="3cc4b-142">Other `fetch` samples</span></span>

* <span data-ttu-id="3cc4b-143">В [примере Use external fetch calls in Office Scripts](../resources/samples/external-fetch-calls.md) показано, как получить базовую информацию о GitHub хранилищах пользователя.</span><span class="sxs-lookup"><span data-stu-id="3cc4b-143">The [Use external fetch calls in Office Scripts](../resources/samples/external-fetch-calls.md) sample shows how to get basic information about a user's GitHub repositories.</span></span>
* <span data-ttu-id="3cc4b-144">Пример [сценария Office сценариев:](../resources/scenarios/noaa-data-fetch.md) Graph данных уровня воды из NOAA демонстрирует команду извлечения, используемую для получения записей из базы данных "Приливы и течения" Национального управления океанических и атмосферных исследований.</span><span class="sxs-lookup"><span data-stu-id="3cc4b-144">The [Office Scripts sample scenario: Graph water-level data from NOAA](../resources/scenarios/noaa-data-fetch.md) demonstrates the fetch command being used to retrieve records from the National Oceanic and Atmospheric Administration's Tides and Currents database.</span></span>

## <a name="external-calls-from-power-automate"></a><span data-ttu-id="3cc4b-145">Внешние вызовы из Power Automate</span><span class="sxs-lookup"><span data-stu-id="3cc4b-145">External calls from Power Automate</span></span>

<span data-ttu-id="3cc4b-146">Любой внешний вызов API не удается при запуске сценария с Power Automate.</span><span class="sxs-lookup"><span data-stu-id="3cc4b-146">Any external API call fails when a script is run with Power Automate.</span></span> <span data-ttu-id="3cc4b-147">Это поведенческая разница между запуском скрипта через приложение Excel и Power Automate.</span><span class="sxs-lookup"><span data-stu-id="3cc4b-147">This is a behavioral difference between running a script through the Excel application and through Power Automate.</span></span> <span data-ttu-id="3cc4b-148">Убедитесь в том, что перед созданием их в поток необходимо проверить сценарии для таких ссылок.</span><span class="sxs-lookup"><span data-stu-id="3cc4b-148">Be sure to check your scripts for such references before building them into a flow.</span></span>

<span data-ttu-id="3cc4b-149">Для получения данных из внешней службы необходимо использовать HTTP с [помощью Azure AD](/connectors/webcontents/) или других эквивалентных действий.</span><span class="sxs-lookup"><span data-stu-id="3cc4b-149">You'll have to use [HTTP with Azure AD](/connectors/webcontents/) or other equivalent actions to pull data from or push it to an external service.</span></span>

> [!WARNING]
> <span data-ttu-id="3cc4b-150">Внешние вызовы, сделанные через соединители Power Automate [Excel Online,](/connectors/excelonlinebusiness) сбой, чтобы помочь поддерживать существующие политики предотвращения потери данных.</span><span class="sxs-lookup"><span data-stu-id="3cc4b-150">External calls made through the Power Automate [Excel Online connector](/connectors/excelonlinebusiness) fail in order to help uphold existing data loss prevention policies.</span></span> <span data-ttu-id="3cc4b-151">Однако сценарии, которые Power Automate, делаются так за пределами организации и вне брандмауэров организации.</span><span class="sxs-lookup"><span data-stu-id="3cc4b-151">However, scripts that are run through Power Automate are done so outside of your organization, and outside of your organization's firewalls.</span></span> <span data-ttu-id="3cc4b-152">Для дополнительной защиты от вредоносных пользователей в этой внешней среде администратор может управлять использованием Office скриптов.</span><span class="sxs-lookup"><span data-stu-id="3cc4b-152">For additional protection from malicious users in this external environment, your admin can control the use of Office Scripts.</span></span> <span data-ttu-id="3cc4b-153">Администратор может отключить соединители Excel Online в Power Automate или отключить Office скрипты для Excel в Интернете с помощью элементов управления администратором [Office скриптов.](/microsoft-365/admin/manage/manage-office-scripts-settings)</span><span class="sxs-lookup"><span data-stu-id="3cc4b-153">Your admin can either disable the Excel Online connector in Power Automate or turn off Office Scripts for Excel on the web through the [Office Scripts administrator controls](/microsoft-365/admin/manage/manage-office-scripts-settings).</span></span>

## <a name="see-also"></a><span data-ttu-id="3cc4b-154">См. также</span><span class="sxs-lookup"><span data-stu-id="3cc4b-154">See also</span></span>

* [<span data-ttu-id="3cc4b-155">Использование встроенных объектов JavaScript в сценариях Office</span><span class="sxs-lookup"><span data-stu-id="3cc4b-155">Using built-in JavaScript objects in Office Scripts</span></span>](javascript-objects.md)
* [<span data-ttu-id="3cc4b-156">Использование внешних вызовов Fetch в сценариях Office</span><span class="sxs-lookup"><span data-stu-id="3cc4b-156">Use external fetch calls in Office Scripts</span></span>](../resources/samples/external-fetch-calls.md)
* [<span data-ttu-id="3cc4b-157">Office Пример сценария: Graph данных уровня воды из NOAA</span><span class="sxs-lookup"><span data-stu-id="3cc4b-157">Office Scripts sample scenario: Graph water-level data from NOAA</span></span>](../resources/scenarios/noaa-data-fetch.md)
