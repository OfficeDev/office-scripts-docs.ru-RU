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
# <a name="external-api-call-support-in-office-scripts"></a><span data-ttu-id="1119c-103">Поддержка внешнего вызова API в сценариях Office</span><span class="sxs-lookup"><span data-stu-id="1119c-103">External API call support in Office Scripts</span></span>

<span data-ttu-id="1119c-104">Авторы скриптов не должны ожидать последовательного поведения при использовании [внешних API](https://developer.mozilla.org/docs/Web/API) на этапе предварительного просмотра платформы.</span><span class="sxs-lookup"><span data-stu-id="1119c-104">Script authors shouldn't expect consistent behavior when using [external APIs](https://developer.mozilla.org/docs/Web/API) during the platform's preview phase.</span></span> <span data-ttu-id="1119c-105">Таким образом, не полагайтесь на внешние API для критических сценариев сценариев.</span><span class="sxs-lookup"><span data-stu-id="1119c-105">As such, do not rely on external APIs for critical script scenarios.</span></span>

<span data-ttu-id="1119c-106">Вызовы на внешние API могут быть сделаны только через Excel, а не через Power Automate при [нормальных обстоятельствах.](#external-calls-from-power-automate)</span><span class="sxs-lookup"><span data-stu-id="1119c-106">Calls to external APIs can be only be made through the Excel application, not through Power Automate [under normal circumstances](#external-calls-from-power-automate).</span></span>

> [!CAUTION]
> <span data-ttu-id="1119c-107">Внешние вызовы могут привести к тому, что конфиденциальные данные будут подвержены нежелательным конечных точкам.</span><span class="sxs-lookup"><span data-stu-id="1119c-107">External calls may result in sensitive data being exposed to undesirable endpoints.</span></span> <span data-ttu-id="1119c-108">Администратор может установить защиту брандмауэра от таких вызовов.</span><span class="sxs-lookup"><span data-stu-id="1119c-108">Your admin can establish firewall protection against such calls.</span></span>

## <a name="configure-your-script-for-external-calls"></a><span data-ttu-id="1119c-109">Настройка скрипта для внешних вызовов</span><span class="sxs-lookup"><span data-stu-id="1119c-109">Configure your script for external calls</span></span>

<span data-ttu-id="1119c-110">Внешние [вызовы являются асинхронными](https://developer.mozilla.org/docs/Learn/JavaScript/Asynchronous/Async_await) и требуют, чтобы ваш скрипт помечен как `async` .</span><span class="sxs-lookup"><span data-stu-id="1119c-110">External calls are [asynchronous](https://developer.mozilla.org/docs/Learn/JavaScript/Asynchronous/Async_await) and require that your script is marked as `async`.</span></span> <span data-ttu-id="1119c-111">Добавьте `async` приставку к вашей `main` функции и верните `Promise` ее, как показано здесь:</span><span class="sxs-lookup"><span data-stu-id="1119c-111">Add the `async` prefix to your `main` function and have it return a `Promise`, as shown here:</span></span>

```typescript
async function main(workbook: ExcelScript.Workbook) : Promise <void>
```

> [!NOTE]
> <span data-ttu-id="1119c-112">Скрипты, возвращают другую информацию, `Promise` могут вернуть этот тип.</span><span class="sxs-lookup"><span data-stu-id="1119c-112">Scripts that return other information can return a `Promise` of that type.</span></span> <span data-ttu-id="1119c-113">Например, если скрипту необходимо вернуть `Employee` объект, ответная подпись будет `: Promise <Employee>`</span><span class="sxs-lookup"><span data-stu-id="1119c-113">For example, if your script needs to return an `Employee` object, the return signature would be `: Promise <Employee>`</span></span>

<span data-ttu-id="1119c-114">Вам нужно будет изучить интерфейсы внешних служб, чтобы звонить в эту службу.</span><span class="sxs-lookup"><span data-stu-id="1119c-114">You'll need to learn the external service's interfaces to make calls to that service.</span></span> <span data-ttu-id="1119c-115">Если вы `fetch` используете [API или REST](https://wikipedia.org/wiki/Representational_state_transfer)API, вам необходимо определить структуру JSON возвращенных данных.</span><span class="sxs-lookup"><span data-stu-id="1119c-115">If you are using `fetch` or [REST APIs](https://wikipedia.org/wiki/Representational_state_transfer), you need to determine the JSON structure of the returned data.</span></span> <span data-ttu-id="1119c-116">Для ввода и вывода из скрипта, рассмотрите возможность создания в `interface` качестве меры, необходимой для создания необходимых структур JSON.</span><span class="sxs-lookup"><span data-stu-id="1119c-116">For both input to and output from your script, consider making an `interface` to match the needed JSON structures.</span></span> <span data-ttu-id="1119c-117">Это дает скрипту больше безопасности типа.</span><span class="sxs-lookup"><span data-stu-id="1119c-117">This gives the script more type safety.</span></span> <span data-ttu-id="1119c-118">Вы можете увидеть пример этого в [использовании извлечения из Office скриптов](../resources/samples/external-fetch-calls.md).</span><span class="sxs-lookup"><span data-stu-id="1119c-118">You can see an example of this in [Using fetch from Office Scripts](../resources/samples/external-fetch-calls.md).</span></span>

### <a name="limitations-with-external-calls-from-office-scripts"></a><span data-ttu-id="1119c-119">Ограничения с внешними вызовами из Office скриптов</span><span class="sxs-lookup"><span data-stu-id="1119c-119">Limitations with external calls from Office Scripts</span></span>

* <span data-ttu-id="1119c-120">Нет никакого способа войти или использовать потоки OAuth2 типа аутентификации.</span><span class="sxs-lookup"><span data-stu-id="1119c-120">There is no way to sign in or use OAuth2 type of authentication flows.</span></span> <span data-ttu-id="1119c-121">Все ключи и учетные данные должны быть жестко закодированы (или читать из другого источника).</span><span class="sxs-lookup"><span data-stu-id="1119c-121">All keys and credentials have to be hardcoded (or read from another source).</span></span>
* <span data-ttu-id="1119c-122">Нет инфраструктуры для хранения учетных данных и ключей API.</span><span class="sxs-lookup"><span data-stu-id="1119c-122">There is no infrastructure to store API credentials and keys.</span></span> <span data-ttu-id="1119c-123">Этим должен управлять пользователь.</span><span class="sxs-lookup"><span data-stu-id="1119c-123">This will have to be managed by the user.</span></span>
* <span data-ttu-id="1119c-124">Файлы `localStorage` cookie-файлов и объекты не `sessionStorage` поддерживаются.</span><span class="sxs-lookup"><span data-stu-id="1119c-124">Document cookies, `localStorage`, and `sessionStorage` objects are not supported.</span></span> 
* <span data-ttu-id="1119c-125">Внешние вызовы могут привести к тому, что конфиденциальные данные будут подвержены нежелательным конечных точкам или внешние данные будут завесят во внутренние рабочие книжки.</span><span class="sxs-lookup"><span data-stu-id="1119c-125">External calls may result in sensitive data being exposed to undesirable endpoints, or external data to be brought into internal workbooks.</span></span> <span data-ttu-id="1119c-126">Администратор может установить защиту брандмауэра от таких вызовов.</span><span class="sxs-lookup"><span data-stu-id="1119c-126">Your admin can establish firewall protection against such calls.</span></span> <span data-ttu-id="1119c-127">Не забудьте проверить с местными политиками, прежде чем полагаться на внешние вызовы.</span><span class="sxs-lookup"><span data-stu-id="1119c-127">Be sure to check with local policies prior to relying on external calls.</span></span>
* <span data-ttu-id="1119c-128">Не забудьте проверить объем пропускной способности данных до принятия зависимости.</span><span class="sxs-lookup"><span data-stu-id="1119c-128">Be sure to check the amount of data throughput prior to taking a dependency.</span></span> <span data-ttu-id="1119c-129">Например, стягивание всего внешнего набора данных может быть не лучшим вариантом, а вместо этого pagination следует использовать для получения данных в кусках.</span><span class="sxs-lookup"><span data-stu-id="1119c-129">For instance, pulling down the entire external dataset may not be the best option and instead pagination should be used to get data in chunks.</span></span>

## <a name="retrieve-information-with-fetch"></a><span data-ttu-id="1119c-130">Получение информации с помощью `fetch`</span><span class="sxs-lookup"><span data-stu-id="1119c-130">Retrieve information with `fetch`</span></span>

<span data-ttu-id="1119c-131">API [получает информацию](https://developer.mozilla.org/docs/Web/API/Fetch_API) из внешних служб.</span><span class="sxs-lookup"><span data-stu-id="1119c-131">The [fetch API](https://developer.mozilla.org/docs/Web/API/Fetch_API) retrieves information from external services.</span></span> <span data-ttu-id="1119c-132">Это `async` API, поэтому вам нужно настроить `main` подпись вашего скрипта.</span><span class="sxs-lookup"><span data-stu-id="1119c-132">It is an `async` API, so you need to adjust the `main` signature of your script.</span></span> <span data-ttu-id="1119c-133">Сделать `main` `async` функцию и заставить его вернуть `Promise<void>` .</span><span class="sxs-lookup"><span data-stu-id="1119c-133">Make the `main` function `async` and have it return a `Promise<void>`.</span></span> <span data-ttu-id="1119c-134">Вы также должны быть уверены `await` в `fetch` вызове `json` и поиске.</span><span class="sxs-lookup"><span data-stu-id="1119c-134">You should also be sure to `await` the `fetch` call and `json` retrieval.</span></span> <span data-ttu-id="1119c-135">Это гарантирует, что эти операции будут завершены до окончания сценария.</span><span class="sxs-lookup"><span data-stu-id="1119c-135">This ensures those operations complete before the script ends.</span></span>

<span data-ttu-id="1119c-136">Любые данные JSON, полученные с `fetch` помощью, должны соответствовать интерфейсу, определенному в скрипте.</span><span class="sxs-lookup"><span data-stu-id="1119c-136">Any JSON data retrieved by `fetch` must match an interface defined in the script.</span></span> <span data-ttu-id="1119c-137">Возвращенное значение должно быть назначено определенному [типу, Office скрипты не поддерживают `any` тип.](typescript-restrictions.md#no-any-type-in-office-scripts)</span><span class="sxs-lookup"><span data-stu-id="1119c-137">The returned value must be assigned to a specific type because [Office Scripts do not support the `any` type](typescript-restrictions.md#no-any-type-in-office-scripts).</span></span> <span data-ttu-id="1119c-138">Вы должны обратиться к документации для вашего сервиса, чтобы увидеть, какие имена и типы возвращенных свойств.</span><span class="sxs-lookup"><span data-stu-id="1119c-138">You should refer to the documentation for your service to see what the names and types of the returned properties are.</span></span> <span data-ttu-id="1119c-139">Затем добавьте соответствующий интерфейс или интерфейсы в свой скрипт.</span><span class="sxs-lookup"><span data-stu-id="1119c-139">Then, add the matching interface or interfaces to your script.</span></span>

<span data-ttu-id="1119c-140">Следующий скрипт используется `fetch` для извлечения данных JSON с тестового сервера в данном URL.</span><span class="sxs-lookup"><span data-stu-id="1119c-140">The following script uses `fetch` to retrieve JSON data from the test server in the given URL.</span></span> <span data-ttu-id="1119c-141">Обратите внимание `JSONData` на интерфейс для хранения данных в качестве подходящего типа.</span><span class="sxs-lookup"><span data-stu-id="1119c-141">Note the `JSONData` interface to store the data as a matching type.</span></span>

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

### <a name="other-fetch-samples"></a><span data-ttu-id="1119c-142">Другие `fetch` образцы</span><span class="sxs-lookup"><span data-stu-id="1119c-142">Other `fetch` samples</span></span>

* <span data-ttu-id="1119c-143">Внешние [вызовы извлечения в Office скриптов](../resources/samples/external-fetch-calls.md) показывает, как получить основную информацию о GitHub пользователя.</span><span class="sxs-lookup"><span data-stu-id="1119c-143">The [Use external fetch calls in Office Scripts](../resources/samples/external-fetch-calls.md) sample shows how to get basic information about a user's GitHub repositories.</span></span>
* <span data-ttu-id="1119c-144">Сценарий [Office Scripts: Graph данные об уровне](../resources/scenarios/noaa-data-fetch.md) воды от NOAA демонстрируют команду извлечения, используемую для извлечения записей из базы данных Приливов и течений Национального управления океанических и атмосферных исследований.</span><span class="sxs-lookup"><span data-stu-id="1119c-144">The [Office Scripts sample scenario: Graph water-level data from NOAA](../resources/scenarios/noaa-data-fetch.md) demonstrates the fetch command being used to retrieve records from the National Oceanic and Atmospheric Administration's Tides and Currents database.</span></span>

## <a name="external-calls-from-power-automate"></a><span data-ttu-id="1119c-145">Внешние звонки из Power Automate</span><span class="sxs-lookup"><span data-stu-id="1119c-145">External calls from Power Automate</span></span>

<span data-ttu-id="1119c-146">Любой внешний вызов API выходит из строя при запуске скрипта с Power Automate.</span><span class="sxs-lookup"><span data-stu-id="1119c-146">Any external API call fails when a script is run with Power Automate.</span></span> <span data-ttu-id="1119c-147">Это поведенческая разница между запуском скрипта через приложение Excel и Power Automate.</span><span class="sxs-lookup"><span data-stu-id="1119c-147">This is a behavioral difference between running a script through the Excel application and through Power Automate.</span></span> <span data-ttu-id="1119c-148">Не забудьте проверить ваши сценарии для таких ссылок, прежде чем строить их в поток.</span><span class="sxs-lookup"><span data-stu-id="1119c-148">Be sure to check your scripts for such references before building them into a flow.</span></span>

<span data-ttu-id="1119c-149">Вам придется использовать HTTP с [Azure AD или другими эквивалентными](/connectors/webcontents/) действиями, чтобы вытащить данные из или перейти на внешнюю службу.</span><span class="sxs-lookup"><span data-stu-id="1119c-149">You'll have to use [HTTP with Azure AD](/connectors/webcontents/) or other equivalent actions to pull data from or push it to an external service.</span></span>

> [!WARNING]
> <span data-ttu-id="1119c-150">Внешние вызовы, сделанные через Power Automate [Excel Online, не удается,](/connectors/excelonlinebusiness) чтобы помочь поддерживать существующие политики предотвращения потери данных.</span><span class="sxs-lookup"><span data-stu-id="1119c-150">External calls made through the Power Automate [Excel Online connector](/connectors/excelonlinebusiness) fail in order to help uphold existing data loss prevention policies.</span></span> <span data-ttu-id="1119c-151">Однако сценарии, которые проходят через Power Automate, сделаны это за пределами вашей организации и за пределами брандмауэров организации.</span><span class="sxs-lookup"><span data-stu-id="1119c-151">However, scripts that are run through Power Automate are done so outside of your organization, and outside of your organization's firewalls.</span></span> <span data-ttu-id="1119c-152">Для дополнительной защиты от вредоносных пользователей в этой внешней среде администратор может контролировать использование Office скриптов.</span><span class="sxs-lookup"><span data-stu-id="1119c-152">For additional protection from malicious users in this external environment, your admin can control the use of Office Scripts.</span></span> <span data-ttu-id="1119c-153">Администратор может либо отключить разъем Excel Online в Power Automate или выключить Office скрипты для Excel в Интернете с [помощью Office управления скриптами.](/microsoft-365/admin/manage/manage-office-scripts-settings)</span><span class="sxs-lookup"><span data-stu-id="1119c-153">Your admin can either disable the Excel Online connector in Power Automate or turn off Office Scripts for Excel on the web through the [Office Scripts administrator controls](/microsoft-365/admin/manage/manage-office-scripts-settings).</span></span>

## <a name="see-also"></a><span data-ttu-id="1119c-154">См. также</span><span class="sxs-lookup"><span data-stu-id="1119c-154">See also</span></span>

* [<span data-ttu-id="1119c-155">Использование встроенных объектов JavaScript в сценариях Office</span><span class="sxs-lookup"><span data-stu-id="1119c-155">Using built-in JavaScript objects in Office Scripts</span></span>](javascript-objects.md)
* [<span data-ttu-id="1119c-156">Использование внешних вызовов Fetch в сценариях Office</span><span class="sxs-lookup"><span data-stu-id="1119c-156">Use external fetch calls in Office Scripts</span></span>](../resources/samples/external-fetch-calls.md)
* [<span data-ttu-id="1119c-157">Office Сценарий выборки сценария: Graph данные уровня воды от NOAA</span><span class="sxs-lookup"><span data-stu-id="1119c-157">Office Scripts sample scenario: Graph water-level data from NOAA</span></span>](../resources/scenarios/noaa-data-fetch.md)
