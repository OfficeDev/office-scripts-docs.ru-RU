---
title: Ограничения TypeScript в сценариях Office
description: Особенности компиляторов TypeScript и linter, используемых редактором кода сценариев Office.
ms.date: 02/05/2021
localization_priority: Normal
ms.openlocfilehash: 87a070b9f342fa5a1f5109fa647bba591832e0cf
ms.sourcegitcommit: 345f1dd96d80471b246044b199fe11126a192a88
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 02/14/2021
ms.locfileid: "50242020"
---
# <a name="typescript-restrictions-in-office-scripts"></a><span data-ttu-id="27671-103">Ограничения TypeScript в сценариях Office</span><span class="sxs-lookup"><span data-stu-id="27671-103">TypeScript restrictions in Office Scripts</span></span>

<span data-ttu-id="27671-104">Сценарии Office используют язык TypeScript.</span><span class="sxs-lookup"><span data-stu-id="27671-104">Office Scripts use the TypeScript language.</span></span> <span data-ttu-id="27671-105">По большей части любой код TypeScript или JavaScript будет работать в сценарии Office.</span><span class="sxs-lookup"><span data-stu-id="27671-105">For the most part, any TypeScript or JavaScript code will work in an Office Script.</span></span> <span data-ttu-id="27671-106">Однако редактор кода на принудительное применение нескольких ограничений гарантирует, что ваш сценарий работает согласованно и в том же отношении с книгой Excel.</span><span class="sxs-lookup"><span data-stu-id="27671-106">However, there are a few restrictions enforced by the Code Editor to ensure your script works consistently and as intended with your Excel workbook.</span></span>

## <a name="no-any-type-in-office-scripts"></a><span data-ttu-id="27671-107">Нет типа "any" в сценариях Office</span><span class="sxs-lookup"><span data-stu-id="27671-107">No 'any' type in Office Scripts</span></span>

<span data-ttu-id="27671-108">Типы [записи](https://www.typescriptlang.org/docs/handbook/typescript-in-5-minutes.html) не являются обязательными в TypeScript, так как эти типы могут быть высмеяны.</span><span class="sxs-lookup"><span data-stu-id="27671-108">Writing [types](https://www.typescriptlang.org/docs/handbook/typescript-in-5-minutes.html) is optional in TypeScript, because the types can be inferred.</span></span> <span data-ttu-id="27671-109">Однако для сценария Office требуется, чтобы переменная не была [типом.](https://www.typescriptlang.org/docs/handbook/basic-types.html#any)</span><span class="sxs-lookup"><span data-stu-id="27671-109">However, Office Script requires that a variable can't be of [type any](https://www.typescriptlang.org/docs/handbook/basic-types.html#any).</span></span> <span data-ttu-id="27671-110">Явные и `any` неявные не допускаются в сценарии Office.</span><span class="sxs-lookup"><span data-stu-id="27671-110">Both explicit and implicit `any` are not allowed in an Office Script.</span></span> <span data-ttu-id="27671-111">Эти случаи сообщаются как ошибки.</span><span class="sxs-lookup"><span data-stu-id="27671-111">These cases are reported as errors.</span></span>

### <a name="explicit-any"></a><span data-ttu-id="27671-112">Explicit `any`</span><span class="sxs-lookup"><span data-stu-id="27671-112">Explicit `any`</span></span>

<span data-ttu-id="27671-113">Нельзя явно объявить переменную типа в `any` скриптах Office (то `let someVariable: any;` есть).</span><span class="sxs-lookup"><span data-stu-id="27671-113">You cannot explicitly declare a variable to be of type `any` in Office Scripts (that is, `let someVariable: any;`).</span></span> <span data-ttu-id="27671-114">Тип `any` вызывает проблемы при обработке Excel.</span><span class="sxs-lookup"><span data-stu-id="27671-114">The `any` type causes issues when processed by Excel.</span></span> <span data-ttu-id="27671-115">Например, необходимо `Range` знать, что значением является `string` , или `number` `boolean` .</span><span class="sxs-lookup"><span data-stu-id="27671-115">For example, a `Range` needs to know that a value is a `string`, `number`, or `boolean`.</span></span> <span data-ttu-id="27671-116">Вы получите ошибку времени компиляции (ошибку перед запуском сценария), если любая переменная явно определена как тип `any` в сценарии.</span><span class="sxs-lookup"><span data-stu-id="27671-116">You will receive a compile-time error (an error prior to running the script) if any variable is explicitly defined as the `any` type in the script.</span></span>

![Явное сообщение в тексте наведении курсоров редактора кода](../images/explicit-any-editor-message.png)

![Явное сообщение об ошибке в окне консоли](../images/explicit-any-error-message.png)

<span data-ttu-id="27671-119">На снимке экрана `[5, 16] Explicit Any is not allowed` выше по указано, что #5 строка, #16 определяет `any` тип.</span><span class="sxs-lookup"><span data-stu-id="27671-119">In the above screenshot `[5, 16] Explicit Any is not allowed` indicates that line #5, column #16 defines `any` type.</span></span> <span data-ttu-id="27671-120">Это помогает найти ошибку.</span><span class="sxs-lookup"><span data-stu-id="27671-120">This helps you locate the error.</span></span>

<span data-ttu-id="27671-121">Чтобы обойти эту проблему, всегда определите тип переменной.</span><span class="sxs-lookup"><span data-stu-id="27671-121">To get around this issue, always define the type of the variable.</span></span> <span data-ttu-id="27671-122">Если вы не уверены в типе переменной, можно использовать [тип объединения.](https://www.typescriptlang.org/docs/handbook/unions-and-intersections.html)</span><span class="sxs-lookup"><span data-stu-id="27671-122">If you are uncertain about the type of a variable, you can use a [union type](https://www.typescriptlang.org/docs/handbook/unions-and-intersections.html).</span></span> <span data-ttu-id="27671-123">Это может быть полезно для переменных, которые удерживают значения, которые могут иметь тип , или (тип для значений это объединение `Range` `string` из `number` `boolean` `Range` них: `string | number | boolean` ).</span><span class="sxs-lookup"><span data-stu-id="27671-123">This can be useful for variables that hold `Range` values, which can be of type `string`, `number`, or `boolean` (the type for `Range` values is a union of those: `string | number | boolean`).</span></span>

### <a name="implicit-any"></a><span data-ttu-id="27671-124">Неявный `any`</span><span class="sxs-lookup"><span data-stu-id="27671-124">Implicit `any`</span></span>

<span data-ttu-id="27671-125">Типы переменных TypeScript можно [определить неявно.](https://www.typescriptlang.org/docs/handbook/type-inference.html)</span><span class="sxs-lookup"><span data-stu-id="27671-125">TypeScript variable types can be [implicitly](https://www.typescriptlang.org/docs/handbook/type-inference.html) defined.</span></span> <span data-ttu-id="27671-126">Если компилятору TypeScript не удается определить тип переменной (либо из-за того, что тип явно не определен, либо вывод типа невозможен), то это неявный параметр, и вы получите ошибку времени `any` компиляции.</span><span class="sxs-lookup"><span data-stu-id="27671-126">If the TypeScript compiler is unable to determine the type of a variable (either because type is not defined explicitly or type inference isn't possible), then it's an implicit `any` and you will receive a compilation-time error.</span></span>

<span data-ttu-id="27671-127">Наиболее распространенный случай для любого неявного `any` параметра — объявление переменной, например `let value;` .</span><span class="sxs-lookup"><span data-stu-id="27671-127">The most common case on any implicit `any` is in a variable declaration, such as `let value;`.</span></span> <span data-ttu-id="27671-128">Этого можно избежать двумя способами:</span><span class="sxs-lookup"><span data-stu-id="27671-128">There are two ways to avoid this:</span></span>

* <span data-ttu-id="27671-129">Назначьте переменную неявно идентифицируемого типа `let value = 5;` `let value = workbook.getWorksheet();` (или).</span><span class="sxs-lookup"><span data-stu-id="27671-129">Assign the variable to an implicitly identifiable type (`let value = 5;` or `let value = workbook.getWorksheet();`).</span></span>
* <span data-ttu-id="27671-130">Явно введите переменную ( `let value: number;` )</span><span class="sxs-lookup"><span data-stu-id="27671-130">Explicitly type the variable (`let value: number;`)</span></span>

## <a name="no-inheriting-office-script-classes-or-interfaces"></a><span data-ttu-id="27671-131">Без наследования классов и интерфейсов сценариев Office</span><span class="sxs-lookup"><span data-stu-id="27671-131">No inheriting Office Script classes or interfaces</span></span>

<span data-ttu-id="27671-132">Классы и интерфейсы, созданные в скрипте Office, не могут расширять или реализовывать классы или интерфейсы сценариев [Office.](https://www.typescriptlang.org/docs/handbook/classes.html#inheritance)</span><span class="sxs-lookup"><span data-stu-id="27671-132">Classes and interfaces that are created in your Office Script cannot [extend or implement](https://www.typescriptlang.org/docs/handbook/classes.html#inheritance) Office Scripts classes or interfaces.</span></span> <span data-ttu-id="27671-133">Другими словами, в пространстве имен не могут быть подклассы `ExcelScript` или подучества.</span><span class="sxs-lookup"><span data-stu-id="27671-133">In other words, nothing in the `ExcelScript` namespace can have subclasses or subinterfaces.</span></span>

## <a name="incompatible-typescript-functions"></a><span data-ttu-id="27671-134">Несовместимые функции TypeScript</span><span class="sxs-lookup"><span data-stu-id="27671-134">Incompatible TypeScript functions</span></span>

<span data-ttu-id="27671-135">API сценариев Office нельзя использовать в следующих сценариях:</span><span class="sxs-lookup"><span data-stu-id="27671-135">Office Scripts APIs cannot be used in the following:</span></span>

* [<span data-ttu-id="27671-136">Функции генератора</span><span class="sxs-lookup"><span data-stu-id="27671-136">Generator functions</span></span>](https://developer.mozilla.org/docs/Web/JavaScript/Guide/Iterators_and_Generators#generator_functions)
* [<span data-ttu-id="27671-137">Array.sort</span><span class="sxs-lookup"><span data-stu-id="27671-137">Array.sort</span></span>](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Array/sort)

## <a name="eval-is-not-supported"></a><span data-ttu-id="27671-138">`eval` не поддерживается</span><span class="sxs-lookup"><span data-stu-id="27671-138">`eval` is not supported</span></span>

<span data-ttu-id="27671-139">Функция [eval](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/eval) JavaScript не поддерживается из соображений безопасности.</span><span class="sxs-lookup"><span data-stu-id="27671-139">The JavaScript [eval function](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/eval) is not supported for security reasons.</span></span>

## <a name="restricted-identifers"></a><span data-ttu-id="27671-140">Ограниченные отступы</span><span class="sxs-lookup"><span data-stu-id="27671-140">Restricted identifers</span></span>

<span data-ttu-id="27671-141">Следующие слова нельзя использовать в качестве идентификаторов в сценарии.</span><span class="sxs-lookup"><span data-stu-id="27671-141">The following words can't be used as identifiers in a script.</span></span> <span data-ttu-id="27671-142">Это зарезервированные термины.</span><span class="sxs-lookup"><span data-stu-id="27671-142">They are reserved terms.</span></span>

* `Excel`
* `ExcelScript`
* `console`

## <a name="only-arrow-functions-in-array-callbacks"></a><span data-ttu-id="27671-143">Функции только со стрелками в вызовах массива</span><span class="sxs-lookup"><span data-stu-id="27671-143">Only arrow functions in array callbacks</span></span>

<span data-ttu-id="27671-144">Ваши сценарии могут использовать функции [со стрелками](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Functions/Arrow_functions) только при предоставлении аргументов вызова для [методов Array.](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Array)</span><span class="sxs-lookup"><span data-stu-id="27671-144">Your scripts can only use [arrow functions](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Functions/Arrow_functions) when providing callback arguments for [Array](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Array) methods.</span></span> <span data-ttu-id="27671-145">Эти методы не могут передавать какой-либо идентификатор или "традиционную" функцию.</span><span class="sxs-lookup"><span data-stu-id="27671-145">You cannot pass any sort of identifier or "traditional" function to these methods.</span></span>

```typescript
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

## <a name="performance-warnings"></a><span data-ttu-id="27671-146">Предупреждения о производительности</span><span class="sxs-lookup"><span data-stu-id="27671-146">Performance warnings</span></span>

<span data-ttu-id="27671-147">Линтер [редактора](https://wikipedia.org/wiki/Lint_(software)) кода выдает предупреждения, если у скрипта могут возникнуть проблемы с производительностью.</span><span class="sxs-lookup"><span data-stu-id="27671-147">The Code Editor's [linter](https://wikipedia.org/wiki/Lint_(software)) gives warnings if the script might have performance issues.</span></span> <span data-ttu-id="27671-148">Сценарии и их работа описаны в документе о повышении производительности [сценариев Office.](web-client-performance.md)</span><span class="sxs-lookup"><span data-stu-id="27671-148">The cases and how to work around them are documented in [Improve the performance of your Office Scripts](web-client-performance.md).</span></span>

## <a name="external-api-calls"></a><span data-ttu-id="27671-149">Вызовы внешних API</span><span class="sxs-lookup"><span data-stu-id="27671-149">External API calls</span></span>

<span data-ttu-id="27671-150">Дополнительные сведения см. в службе поддержки вызовов внешнего [API в сценариях Office.](external-calls.md)</span><span class="sxs-lookup"><span data-stu-id="27671-150">See [External API call support in Office Scripts](external-calls.md) for more information.</span></span>

## <a name="see-also"></a><span data-ttu-id="27671-151">См. также</span><span class="sxs-lookup"><span data-stu-id="27671-151">See also</span></span>

* [<span data-ttu-id="27671-152">Основные сведения о сценариях Office в Excel в Интернете</span><span class="sxs-lookup"><span data-stu-id="27671-152">Scripting fundamentals for Office Scripts in Excel on the web</span></span>](scripting-fundamentals.md)
* [<span data-ttu-id="27671-153">Повышение производительности сценариев Office</span><span class="sxs-lookup"><span data-stu-id="27671-153">Improve the performance of your Office Scripts</span></span>](web-client-performance.md)
