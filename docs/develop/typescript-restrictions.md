---
title: Ограничения TypeScript в сценариях Office
description: Особенности компиляторов TypeScript и linter, используемых редактором кода сценариев Office.
ms.date: 01/29/2021
localization_priority: Normal
ms.openlocfilehash: d67e208561ce6ddd706d4c80cf29d2f013a32032
ms.sourcegitcommit: 98c7bc26f51dc8427669c571135c503d73bcee4c
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 02/06/2021
ms.locfileid: "50125936"
---
# <a name="typescript-restrictions-in-office-scripts"></a><span data-ttu-id="27746-103">Ограничения TypeScript в сценариях Office</span><span class="sxs-lookup"><span data-stu-id="27746-103">TypeScript restrictions in Office Scripts</span></span>

<span data-ttu-id="27746-104">Сценарии Office используют язык TypeScript.</span><span class="sxs-lookup"><span data-stu-id="27746-104">Office Scripts use the TypeScript language.</span></span> <span data-ttu-id="27746-105">По большей части любой код TypeScript или JavaScript будет работать в сценарии Office.</span><span class="sxs-lookup"><span data-stu-id="27746-105">For the most part, any TypeScript or JavaScript code will work in an Office Script.</span></span> <span data-ttu-id="27746-106">Однако редактор кода на принудительное применение нескольких ограничений гарантирует, что ваш сценарий работает согласованно и в том же отношении с книгой Excel.</span><span class="sxs-lookup"><span data-stu-id="27746-106">However, there are a few restrictions enforced by the Code Editor to ensure your script works consistently and as intended with your Excel workbook.</span></span>

## <a name="no-any-type-in-office-scripts"></a><span data-ttu-id="27746-107">Нет типа "any" в сценариях Office</span><span class="sxs-lookup"><span data-stu-id="27746-107">No 'any' type in Office Scripts</span></span>

<span data-ttu-id="27746-108">Типы [записи](https://www.typescriptlang.org/docs/handbook/typescript-in-5-minutes.html) необязательны в TypeScript, так как эти типы могут быть высмеяны.</span><span class="sxs-lookup"><span data-stu-id="27746-108">Writing [types](https://www.typescriptlang.org/docs/handbook/typescript-in-5-minutes.html) is optional in TypeScript, because the types can be inferred.</span></span> <span data-ttu-id="27746-109">Однако сценарий Office требует, чтобы переменная не была [типом.](https://www.typescriptlang.org/docs/handbook/basic-types.html#any)</span><span class="sxs-lookup"><span data-stu-id="27746-109">However, Office Script requires that a variable can't be of [type any](https://www.typescriptlang.org/docs/handbook/basic-types.html#any).</span></span> <span data-ttu-id="27746-110">Явные и `any` неявные не допускаются в сценарии Office.</span><span class="sxs-lookup"><span data-stu-id="27746-110">Both explicit and implicit `any` are not allowed in an Office Script.</span></span> <span data-ttu-id="27746-111">Эти случаи сообщаются как ошибки.</span><span class="sxs-lookup"><span data-stu-id="27746-111">These cases are reported as errors.</span></span>

### <a name="explicit-any"></a><span data-ttu-id="27746-112">Explicit `any`</span><span class="sxs-lookup"><span data-stu-id="27746-112">Explicit `any`</span></span>

<span data-ttu-id="27746-113">Нельзя явно объявить переменную типа в `any` скриптах Office (то `let someVariable: any;` есть).</span><span class="sxs-lookup"><span data-stu-id="27746-113">You cannot explicitly declare a variable to be of type `any` in Office Scripts (that is, `let someVariable: any;`).</span></span> <span data-ttu-id="27746-114">Тип `any` вызывает проблемы при обработке Excel.</span><span class="sxs-lookup"><span data-stu-id="27746-114">The `any` type causes issues when processed by Excel.</span></span> <span data-ttu-id="27746-115">Например, необходимо `Range` знать, что значением является `string` , или `number` `boolean` .</span><span class="sxs-lookup"><span data-stu-id="27746-115">For example, a `Range` needs to know that a value is a `string`, `number`, or `boolean`.</span></span> <span data-ttu-id="27746-116">Вы получите ошибку времени компиляции (ошибку перед запуском сценария), если любая переменная явно определена как тип `any` в сценарии.</span><span class="sxs-lookup"><span data-stu-id="27746-116">You will receive a compile-time error (an error prior to running the script) if any variable is explicitly defined as the `any` type in the script.</span></span>

![Явное сообщение в тексте наведении курсоров редактора кода](../images/explicit-any-editor-message.png)

![Явная любая ошибка в окне консоли](../images/explicit-any-error-message.png)

<span data-ttu-id="27746-119">На снимке экрана выше `[5, 16] Explicit Any is not allowed` по указано, что #5 строка, #16 определяет `any` тип.</span><span class="sxs-lookup"><span data-stu-id="27746-119">In the above screenshot `[5, 16] Explicit Any is not allowed` indicates that line #5, column #16 defines `any` type.</span></span> <span data-ttu-id="27746-120">Это помогает найти ошибку.</span><span class="sxs-lookup"><span data-stu-id="27746-120">This helps you locate the error.</span></span>

<span data-ttu-id="27746-121">Чтобы обойти эту проблему, всегда определите тип переменной.</span><span class="sxs-lookup"><span data-stu-id="27746-121">To get around this issue, always define the type of the variable.</span></span> <span data-ttu-id="27746-122">Если вы не уверены в типе переменной, можно использовать [тип объединения.](https://www.typescriptlang.org/docs/handbook/unions-and-intersections.html)</span><span class="sxs-lookup"><span data-stu-id="27746-122">If you are uncertain about the type of a variable, you can use a [union type](https://www.typescriptlang.org/docs/handbook/unions-and-intersections.html).</span></span> <span data-ttu-id="27746-123">Это может быть полезно для переменных, которые удерживают значения, которые могут иметь тип , или (тип для значений это `Range` `string` объединение из `number` `boolean` `Range` них: `string | number | boolean` ).</span><span class="sxs-lookup"><span data-stu-id="27746-123">This can be useful for variables that hold `Range` values, which can be of type `string`, `number`, or `boolean` (the type for `Range` values is a union of those: `string | number | boolean`).</span></span>

### <a name="implicit-any"></a><span data-ttu-id="27746-124">Неявный `any`</span><span class="sxs-lookup"><span data-stu-id="27746-124">Implicit `any`</span></span>

<span data-ttu-id="27746-125">Типы переменных TypeScript можно определить неявно. [](https://www.typescriptlang.org/docs/handbook/type-inference.html)</span><span class="sxs-lookup"><span data-stu-id="27746-125">TypeScript variable types can be [implicitly](https://www.typescriptlang.org/docs/handbook/type-inference.html) defined.</span></span> <span data-ttu-id="27746-126">Если компилятору TypeScript не удается определить тип переменной (либо из-за того, что тип явно не определен, либо вывод типа невозможен), то это неявный параметр, и вы получите ошибку времени `any` компиляции.</span><span class="sxs-lookup"><span data-stu-id="27746-126">If the TypeScript compiler is unable to determine the type of a variable (either because type is not defined explicitly or type inference isn't possible), then it's an implicit `any` and you will receive a compilation-time error.</span></span>

<span data-ttu-id="27746-127">Наиболее распространенный случай для любого неявного `any` параметра — объявление переменной, например `let value;` .</span><span class="sxs-lookup"><span data-stu-id="27746-127">The most common case on any implicit `any` is in a variable declaration, such as `let value;`.</span></span> <span data-ttu-id="27746-128">Этого можно избежать двумя способами:</span><span class="sxs-lookup"><span data-stu-id="27746-128">There are two ways to avoid this:</span></span>

* <span data-ttu-id="27746-129">Назначьте переменную неявно идентифицируемого типа `let value = 5;` `let value = workbook.getWorksheet();` (или).</span><span class="sxs-lookup"><span data-stu-id="27746-129">Assign the variable to an implicitly identifiable type (`let value = 5;` or `let value = workbook.getWorksheet();`).</span></span>
* <span data-ttu-id="27746-130">Явно введите переменную ( `let value: number;` )</span><span class="sxs-lookup"><span data-stu-id="27746-130">Explicitly type the variable (`let value: number;`)</span></span>

## <a name="no-inheriting-office-script-classes-or-interfaces"></a><span data-ttu-id="27746-131">Без наследования классов и интерфейсов сценариев Office</span><span class="sxs-lookup"><span data-stu-id="27746-131">No inheriting Office Script classes or interfaces</span></span>

<span data-ttu-id="27746-132">Классы и интерфейсы, созданные в скрипте Office, не могут расширять или реализовывать классы или интерфейсы сценариев [Office.](https://www.typescriptlang.org/docs/handbook/classes.html#inheritance)</span><span class="sxs-lookup"><span data-stu-id="27746-132">Classes and interfaces that are created in your Office Script cannot [extend or implement](https://www.typescriptlang.org/docs/handbook/classes.html#inheritance) Office Scripts classes or interfaces.</span></span> <span data-ttu-id="27746-133">Другими словами, в пространстве имен не могут быть подклассы `ExcelScript` или подучества.</span><span class="sxs-lookup"><span data-stu-id="27746-133">In other words, nothing in the `ExcelScript` namespace can have subclasses or subinterfaces.</span></span>

## <a name="incompatible-typescript-functions"></a><span data-ttu-id="27746-134">Несовместимые функции TypeScript</span><span class="sxs-lookup"><span data-stu-id="27746-134">Incompatible TypeScript functions</span></span>

<span data-ttu-id="27746-135">API сценариев Office нельзя использовать в следующих сценариях:</span><span class="sxs-lookup"><span data-stu-id="27746-135">Office Scripts APIs cannot be used in the following:</span></span>

* [<span data-ttu-id="27746-136">Функции генератора</span><span class="sxs-lookup"><span data-stu-id="27746-136">Generator functions</span></span>](https://developer.mozilla.org/docs/Web/JavaScript/Guide/Iterators_and_Generators#generator_functions)
* [<span data-ttu-id="27746-137">Array.sort</span><span class="sxs-lookup"><span data-stu-id="27746-137">Array.sort</span></span>](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Array/sort)

## <a name="eval-is-not-supported"></a><span data-ttu-id="27746-138">`eval` не поддерживается</span><span class="sxs-lookup"><span data-stu-id="27746-138">`eval` is not supported</span></span>

<span data-ttu-id="27746-139">Функция [eval](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/eval) JavaScript не поддерживается из соображений безопасности.</span><span class="sxs-lookup"><span data-stu-id="27746-139">The JavaScript [eval function](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/eval) is not supported for security reasons.</span></span>

## <a name="restricted-identifers"></a><span data-ttu-id="27746-140">Ограниченные отступы</span><span class="sxs-lookup"><span data-stu-id="27746-140">Restricted identifers</span></span>

<span data-ttu-id="27746-141">Следующие слова нельзя использовать в качестве идентификаторов в сценарии.</span><span class="sxs-lookup"><span data-stu-id="27746-141">The following words can't be used as identifiers in a script.</span></span> <span data-ttu-id="27746-142">Это зарезервированные термины.</span><span class="sxs-lookup"><span data-stu-id="27746-142">They are reserved terms.</span></span>

* `Excel`
* `ExcelScript`
* `console`

## <a name="performance-warnings"></a><span data-ttu-id="27746-143">Предупреждения о производительности</span><span class="sxs-lookup"><span data-stu-id="27746-143">Performance warnings</span></span>

<span data-ttu-id="27746-144">Линтер [редактора](https://wikipedia.org/wiki/Lint_(software)) кода выдает предупреждения, если у скрипта могут возникнуть проблемы с производительностью.</span><span class="sxs-lookup"><span data-stu-id="27746-144">The Code Editor's [linter](https://wikipedia.org/wiki/Lint_(software)) gives warnings if the script might have performance issues.</span></span> <span data-ttu-id="27746-145">Сценарии и их работа описаны в документе о повышении производительности [сценариев Office.](web-client-performance.md)</span><span class="sxs-lookup"><span data-stu-id="27746-145">The cases and how to work around them are documented in [Improve the performance of your Office Scripts](web-client-performance.md).</span></span>

## <a name="external-api-calls"></a><span data-ttu-id="27746-146">Вызовы внешних API</span><span class="sxs-lookup"><span data-stu-id="27746-146">External API calls</span></span>

<span data-ttu-id="27746-147">Дополнительные сведения см. в службе поддержки вызовов внешнего [API в сценариях Office.](external-calls.md)</span><span class="sxs-lookup"><span data-stu-id="27746-147">See [External API call support in Office Scripts](external-calls.md) for more information.</span></span>

## <a name="see-also"></a><span data-ttu-id="27746-148">См. также</span><span class="sxs-lookup"><span data-stu-id="27746-148">See also</span></span>

* [<span data-ttu-id="27746-149">Основные сведения о сценариях Office в Excel в Интернете</span><span class="sxs-lookup"><span data-stu-id="27746-149">Scripting fundamentals for Office Scripts in Excel on the web</span></span>](scripting-fundamentals.md)
* [<span data-ttu-id="27746-150">Повышение производительности сценариев Office</span><span class="sxs-lookup"><span data-stu-id="27746-150">Improve the performance of your Office Scripts</span></span>](web-client-performance.md)
