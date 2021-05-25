---
title: Ограничения TypeScript в Office скриптах
description: Особенности компиляторов и подкладок TypeScript, используемых редактором кода Office скриптов.
ms.date: 05/24/2021
localization_priority: Normal
ms.openlocfilehash: 449a8abbcfdcfde53d0c9b96106f73259de368b1
ms.sourcegitcommit: 90ca8cdf30f2065f63938f6bb6780d024c128467
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 05/25/2021
ms.locfileid: "52639869"
---
# <a name="typescript-restrictions-in-office-scripts"></a><span data-ttu-id="a9705-103">Ограничения TypeScript в Office скриптах</span><span class="sxs-lookup"><span data-stu-id="a9705-103">TypeScript restrictions in Office Scripts</span></span>

<span data-ttu-id="a9705-104">Office Скрипты используют язык TypeScript.</span><span class="sxs-lookup"><span data-stu-id="a9705-104">Office Scripts use the TypeScript language.</span></span> <span data-ttu-id="a9705-105">По большей части любой код TypeScript или JavaScript будет работать в Office скриптах.</span><span class="sxs-lookup"><span data-stu-id="a9705-105">For the most part, any TypeScript or JavaScript code will work in Office Scripts.</span></span> <span data-ttu-id="a9705-106">Однако редактор кода соблюдает несколько ограничений, чтобы гарантировать, что сценарий работает последовательно и по назначению с Excel книгой.</span><span class="sxs-lookup"><span data-stu-id="a9705-106">However, there are a few restrictions enforced by the Code Editor to ensure your script works consistently and as intended with your Excel workbook.</span></span>

## <a name="no-any-type-in-office-scripts"></a><span data-ttu-id="a9705-107">Нет типа "любой" в Office скриптах</span><span class="sxs-lookup"><span data-stu-id="a9705-107">No 'any' type in Office Scripts</span></span>

<span data-ttu-id="a9705-108">Типы [записи](https://www.typescriptlang.org/docs/handbook/typescript-in-5-minutes.html) необязательны в TypeScript, так как эти типы можно сделать вывод.</span><span class="sxs-lookup"><span data-stu-id="a9705-108">Writing [types](https://www.typescriptlang.org/docs/handbook/typescript-in-5-minutes.html) is optional in TypeScript, because the types can be inferred.</span></span> <span data-ttu-id="a9705-109">Однако для Office скриптов требуется, чтобы переменная не была [типной.](https://www.typescriptlang.org/docs/handbook/basic-types.html#any)</span><span class="sxs-lookup"><span data-stu-id="a9705-109">However, Office Scripts requires that a variable can't be of [type any](https://www.typescriptlang.org/docs/handbook/basic-types.html#any).</span></span> <span data-ttu-id="a9705-110">Явные и неявные не допускаются `any` в Office скриптах.</span><span class="sxs-lookup"><span data-stu-id="a9705-110">Both explicit and implicit `any` are not allowed in Office Scripts.</span></span> <span data-ttu-id="a9705-111">Эти случаи сообщаются как ошибки.</span><span class="sxs-lookup"><span data-stu-id="a9705-111">These cases are reported as errors.</span></span>

### <a name="explicit-any"></a><span data-ttu-id="a9705-112">Явный `any`</span><span class="sxs-lookup"><span data-stu-id="a9705-112">Explicit `any`</span></span>

<span data-ttu-id="a9705-113">Нельзя явно объявить переменную типом в `any` Office Скрипты (то `let value: any;` есть).</span><span class="sxs-lookup"><span data-stu-id="a9705-113">You cannot explicitly declare a variable to be of type `any` in Office Scripts (that is, `let value: any;`).</span></span> <span data-ttu-id="a9705-114">Тип `any` вызывает проблемы при обработке Excel.</span><span class="sxs-lookup"><span data-stu-id="a9705-114">The `any` type causes issues when processed by Excel.</span></span> <span data-ttu-id="a9705-115">Например, необходимо знать, что значение `Range` является `string` значением , или `number` `boolean` .</span><span class="sxs-lookup"><span data-stu-id="a9705-115">For example, a `Range` needs to know that a value is a `string`, `number`, or `boolean`.</span></span> <span data-ttu-id="a9705-116">Вы получите ошибку времени компиляции (ошибка перед запуском скрипта), если любая переменная явно определена как `any` тип сценария.</span><span class="sxs-lookup"><span data-stu-id="a9705-116">You will receive a compile-time error (an error prior to running the script) if any variable is explicitly defined as the `any` type in the script.</span></span>

:::image type="content" source="../images/explicit-any-editor-message.png" alt-text="Явное сообщение &quot;любое&quot; в тексте наведении редактора кода":::

:::image type="content" source="../images/explicit-any-error-message.png" alt-text="Явная ошибка &quot;любая&quot; в окне консоли":::

<span data-ttu-id="a9705-119">На предыдущем скриншоте `[2, 14] Explicit Any is not allowed` указывается, что строка #2, столбец #14 определяет `any` тип.</span><span class="sxs-lookup"><span data-stu-id="a9705-119">In the previous screenshot, `[2, 14] Explicit Any is not allowed` indicates that line #2, column #14 defines `any` type.</span></span> <span data-ttu-id="a9705-120">Это поможет найти ошибку.</span><span class="sxs-lookup"><span data-stu-id="a9705-120">This helps you locate the error.</span></span>

<span data-ttu-id="a9705-121">Чтобы обойти эту проблему, всегда определите тип переменной.</span><span class="sxs-lookup"><span data-stu-id="a9705-121">To get around this issue, always define the type of the variable.</span></span> <span data-ttu-id="a9705-122">Если вы не уверены в типе переменной, можно использовать [тип union.](https://www.typescriptlang.org/docs/handbook/unions-and-intersections.html)</span><span class="sxs-lookup"><span data-stu-id="a9705-122">If you are uncertain about the type of a variable, you can use a [union type](https://www.typescriptlang.org/docs/handbook/unions-and-intersections.html).</span></span> <span data-ttu-id="a9705-123">Это может быть полезно для переменных, которые держат значения, которые могут быть типа , или (тип для значений является `Range` `string` `number` `boolean` `Range` союзом из них: `string | number | boolean` ).</span><span class="sxs-lookup"><span data-stu-id="a9705-123">This can be useful for variables that hold `Range` values, which can be of type `string`, `number`, or `boolean` (the type for `Range` values is a union of those: `string | number | boolean`).</span></span>

### <a name="implicit-any"></a><span data-ttu-id="a9705-124">Неявный `any`</span><span class="sxs-lookup"><span data-stu-id="a9705-124">Implicit `any`</span></span>

<span data-ttu-id="a9705-125">Типы переменных TypeScript можно [неявно](https://www.typescriptlang.org/docs/handbook/type-inference.html) определить.</span><span class="sxs-lookup"><span data-stu-id="a9705-125">TypeScript variable types can be [implicitly](https://www.typescriptlang.org/docs/handbook/type-inference.html) defined.</span></span> <span data-ttu-id="a9705-126">Если компилятор TypeScript не может определить тип переменной (либо из-за того, что тип явно не определен, либо вывод типа невозможен), то это неявное значение, и вы получите ошибку времени `any` компиляции.</span><span class="sxs-lookup"><span data-stu-id="a9705-126">If the TypeScript compiler is unable to determine the type of a variable (either because type is not defined explicitly or type inference isn't possible), then it's an implicit `any` and you will receive a compilation-time error.</span></span>

:::image type="content" source="../images/implicit-any-editor-message.png" alt-text="Неявное сообщение &quot;любое&quot; в тексте наведении редактора кода":::

<span data-ttu-id="a9705-128">Наиболее распространенный случай для любого неявного `any` находится в переменной декларации, например `let value;` .</span><span class="sxs-lookup"><span data-stu-id="a9705-128">The most common case on any implicit `any` is in a variable declaration, such as `let value;`.</span></span> <span data-ttu-id="a9705-129">Существует два способа избежать этого:</span><span class="sxs-lookup"><span data-stu-id="a9705-129">There are two ways to avoid this:</span></span>

* <span data-ttu-id="a9705-130">Назначение переменной неявно идентифицируемого типа `let value = 5;` `let value = workbook.getWorksheet();` (или).</span><span class="sxs-lookup"><span data-stu-id="a9705-130">Assign the variable to an implicitly identifiable type (`let value = 5;` or `let value = workbook.getWorksheet();`).</span></span>
* <span data-ttu-id="a9705-131">Явно введите переменную ( `let value: number;` )</span><span class="sxs-lookup"><span data-stu-id="a9705-131">Explicitly type the variable (`let value: number;`)</span></span>

## <a name="no-inheriting-office-script-classes-or-interfaces"></a><span data-ttu-id="a9705-132">Нет наследующих Office классов или интерфейсов скриптов</span><span class="sxs-lookup"><span data-stu-id="a9705-132">No inheriting Office Script classes or interfaces</span></span>

<span data-ttu-id="a9705-133">Классы и интерфейсы, созданные в Office скрипта, не могут расширять или [внедрять](https://www.typescriptlang.org/docs/handbook/classes.html#inheritance) Office скрипты или интерфейсы.</span><span class="sxs-lookup"><span data-stu-id="a9705-133">Classes and interfaces that are created in your Office Script cannot [extend or implement](https://www.typescriptlang.org/docs/handbook/classes.html#inheritance) Office Scripts classes or interfaces.</span></span> <span data-ttu-id="a9705-134">Другими словами, ничто в пространстве имен не может `ExcelScript` иметь подклассов или подинтерфейсов.</span><span class="sxs-lookup"><span data-stu-id="a9705-134">In other words, nothing in the `ExcelScript` namespace can have subclasses or subinterfaces.</span></span>

## <a name="incompatible-typescript-functions"></a><span data-ttu-id="a9705-135">Несовместимые функции TypeScript</span><span class="sxs-lookup"><span data-stu-id="a9705-135">Incompatible TypeScript functions</span></span>

<span data-ttu-id="a9705-136">Office API скриптов нельзя использовать в следующих следующих сценариях:</span><span class="sxs-lookup"><span data-stu-id="a9705-136">Office Scripts APIs cannot be used in the following:</span></span>

* [<span data-ttu-id="a9705-137">Функции генератора</span><span class="sxs-lookup"><span data-stu-id="a9705-137">Generator functions</span></span>](https://developer.mozilla.org/docs/Web/JavaScript/Guide/Iterators_and_Generators#generator_functions)
* [<span data-ttu-id="a9705-138">Array.sort</span><span class="sxs-lookup"><span data-stu-id="a9705-138">Array.sort</span></span>](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Array/sort)

## <a name="eval-is-not-supported"></a><span data-ttu-id="a9705-139">`eval` не поддерживается</span><span class="sxs-lookup"><span data-stu-id="a9705-139">`eval` is not supported</span></span>

<span data-ttu-id="a9705-140">Функция [eval JavaScript](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/eval) не поддерживается из соображений безопасности.</span><span class="sxs-lookup"><span data-stu-id="a9705-140">The JavaScript [eval function](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/eval) is not supported for security reasons.</span></span>

## <a name="restricted-identifers"></a><span data-ttu-id="a9705-141">Ограниченные identifers</span><span class="sxs-lookup"><span data-stu-id="a9705-141">Restricted identifers</span></span>

<span data-ttu-id="a9705-142">Следующие слова нельзя использовать в качестве идентификаторов в скрипте.</span><span class="sxs-lookup"><span data-stu-id="a9705-142">The following words can't be used as identifiers in a script.</span></span> <span data-ttu-id="a9705-143">Это зарезервированные условия.</span><span class="sxs-lookup"><span data-stu-id="a9705-143">They are reserved terms.</span></span>

* `Excel`
* `ExcelScript`
* `console`

## <a name="only-arrow-functions-in-array-callbacks"></a><span data-ttu-id="a9705-144">Только функции стрелки в вызовах массива</span><span class="sxs-lookup"><span data-stu-id="a9705-144">Only arrow functions in array callbacks</span></span>

<span data-ttu-id="a9705-145">В скриптах можно использовать функции [стрелки](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Functions/Arrow_functions) только при предоставлении аргументов вызова для [методов Array.](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Array)</span><span class="sxs-lookup"><span data-stu-id="a9705-145">Your scripts can only use [arrow functions](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Functions/Arrow_functions) when providing callback arguments for [Array](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Array) methods.</span></span> <span data-ttu-id="a9705-146">Эти методы не могут передавать какие-либо идентификаторы или "традиционные" функции.</span><span class="sxs-lookup"><span data-stu-id="a9705-146">You cannot pass any sort of identifier or "traditional" function to these methods.</span></span>

```TypeScript
const myArray = [1, 2, 3, 4, 5, 6];
let filteredArray = myArray.filter((x) => {
  return x % 2 === 0;
});
/*
  The following code generates a compiler error in the Office Scripts Code Editor.
  filteredArray = myArray.filter(function (x) {
    return x % 2 === 0;
  });
*/
```

## <a name="performance-warnings"></a><span data-ttu-id="a9705-147">Предупреждения о производительности</span><span class="sxs-lookup"><span data-stu-id="a9705-147">Performance warnings</span></span>

<span data-ttu-id="a9705-148">Подкладка редактора кода [дает](https://wikipedia.org/wiki/Lint_(software)) предупреждения, если у скрипта могут возникнуть проблемы с производительностью.</span><span class="sxs-lookup"><span data-stu-id="a9705-148">The Code Editor's [linter](https://wikipedia.org/wiki/Lint_(software)) gives warnings if the script might have performance issues.</span></span> <span data-ttu-id="a9705-149">Случаи и их работа описаны в документе Улучшение производительности [Office скриптов.](web-client-performance.md)</span><span class="sxs-lookup"><span data-stu-id="a9705-149">The cases and how to work around them are documented in [Improve the performance of your Office Scripts](web-client-performance.md).</span></span>

## <a name="external-api-calls"></a><span data-ttu-id="a9705-150">Внешние вызовы API</span><span class="sxs-lookup"><span data-stu-id="a9705-150">External API calls</span></span>

<span data-ttu-id="a9705-151">Дополнительные сведения см. в Office службе поддержки [вызовов API.](external-calls.md)</span><span class="sxs-lookup"><span data-stu-id="a9705-151">See [External API call support in Office Scripts](external-calls.md) for more information.</span></span>

## <a name="see-also"></a><span data-ttu-id="a9705-152">См. также</span><span class="sxs-lookup"><span data-stu-id="a9705-152">See also</span></span>

* [<span data-ttu-id="a9705-153">Основные сведения о сценариях Office в Excel в Интернете</span><span class="sxs-lookup"><span data-stu-id="a9705-153">Scripting fundamentals for Office Scripts in Excel on the web</span></span>](scripting-fundamentals.md)
* [<span data-ttu-id="a9705-154">Повышение производительности Office скриптов</span><span class="sxs-lookup"><span data-stu-id="a9705-154">Improve the performance of your Office Scripts</span></span>](web-client-performance.md)
