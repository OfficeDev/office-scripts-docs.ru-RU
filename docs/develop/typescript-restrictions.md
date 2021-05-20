---
title: Ограничения TypeScript в Office скриптах
description: Специфика компилятора TypeScript и линтера, используемого Office Code Scripts.
ms.date: 02/05/2021
localization_priority: Normal
ms.openlocfilehash: a4198e0e56224ac5da89e89c43c8d2f3ef44d6d7
ms.sourcegitcommit: 4687693f02fc90a57ba30c461f35046e02e6f5fb
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 05/19/2021
ms.locfileid: "52545021"
---
# <a name="typescript-restrictions-in-office-scripts"></a><span data-ttu-id="20e3e-103">Ограничения TypeScript в Office скриптах</span><span class="sxs-lookup"><span data-stu-id="20e3e-103">TypeScript restrictions in Office Scripts</span></span>

<span data-ttu-id="20e3e-104">Office Скрипты используют язык TypeScript.</span><span class="sxs-lookup"><span data-stu-id="20e3e-104">Office Scripts use the TypeScript language.</span></span> <span data-ttu-id="20e3e-105">По большей части любой код TypeScript или JavaScript будет работать в Office скриптах.</span><span class="sxs-lookup"><span data-stu-id="20e3e-105">For the most part, any TypeScript or JavaScript code will work in Office Scripts.</span></span> <span data-ttu-id="20e3e-106">Тем не менее, редактор Кода применяет несколько ограничений для обеспечения того, чтобы ваш скрипт работал последовательно и по назначению с вашей Excel книгой.</span><span class="sxs-lookup"><span data-stu-id="20e3e-106">However, there are a few restrictions enforced by the Code Editor to ensure your script works consistently and as intended with your Excel workbook.</span></span>

## <a name="no-any-type-in-office-scripts"></a><span data-ttu-id="20e3e-107">Нет типа "любой" в Office скриптах</span><span class="sxs-lookup"><span data-stu-id="20e3e-107">No 'any' type in Office Scripts</span></span>

<span data-ttu-id="20e3e-108">Типы [писания](https://www.typescriptlang.org/docs/handbook/typescript-in-5-minutes.html) не являются обязательными в TypeScript, потому что типы могут быть выведены.</span><span class="sxs-lookup"><span data-stu-id="20e3e-108">Writing [types](https://www.typescriptlang.org/docs/handbook/typescript-in-5-minutes.html) is optional in TypeScript, because the types can be inferred.</span></span> <span data-ttu-id="20e3e-109">Тем не Office, что скрипты требуют, чтобы переменная не может быть [типа любого](https://www.typescriptlang.org/docs/handbook/basic-types.html#any).</span><span class="sxs-lookup"><span data-stu-id="20e3e-109">However, Office Scripts requires that a variable can't be of [type any](https://www.typescriptlang.org/docs/handbook/basic-types.html#any).</span></span> <span data-ttu-id="20e3e-110">Как явные, так `any` и неявные не допускаются Office скриптов.</span><span class="sxs-lookup"><span data-stu-id="20e3e-110">Both explicit and implicit `any` are not allowed in Office Scripts.</span></span> <span data-ttu-id="20e3e-111">Эти случаи сообщаются как ошибки.</span><span class="sxs-lookup"><span data-stu-id="20e3e-111">These cases are reported as errors.</span></span>

### <a name="explicit-any"></a><span data-ttu-id="20e3e-112">явный `any`</span><span class="sxs-lookup"><span data-stu-id="20e3e-112">Explicit `any`</span></span>

<span data-ttu-id="20e3e-113">Вы не можете прямо объявить переменную `any` типом в Office (то `let someVariable: any;` есть).</span><span class="sxs-lookup"><span data-stu-id="20e3e-113">You cannot explicitly declare a variable to be of type `any` in Office Scripts (that is, `let someVariable: any;`).</span></span> <span data-ttu-id="20e3e-114">Тип `any` вызывает проблемы при обработке Excel.</span><span class="sxs-lookup"><span data-stu-id="20e3e-114">The `any` type causes issues when processed by Excel.</span></span> <span data-ttu-id="20e3e-115">Например, `Range` нужно знать, что значение `string` `number` является, или `boolean` .</span><span class="sxs-lookup"><span data-stu-id="20e3e-115">For example, a `Range` needs to know that a value is a `string`, `number`, or `boolean`.</span></span> <span data-ttu-id="20e3e-116">Вы получите ошибку времени компиляции (ошибка до запуска скрипта), если какая-либо переменная явно `any` определена как тип скрипта.</span><span class="sxs-lookup"><span data-stu-id="20e3e-116">You will receive a compile-time error (an error prior to running the script) if any variable is explicitly defined as the `any` type in the script.</span></span>

:::image type="content" source="../images/explicit-any-editor-message.png" alt-text="Явное сообщение «любого» в тексте наведении редактора Кода":::

:::image type="content" source="../images/explicit-any-error-message.png" alt-text="Явная ошибка «любого» в окне консоли":::

<span data-ttu-id="20e3e-119">На предыдущем скриншоте `[5, 16] Explicit Any is not allowed` указывается, что #5, #16 определяет `any` тип.</span><span class="sxs-lookup"><span data-stu-id="20e3e-119">In the previous screenshot `[5, 16] Explicit Any is not allowed` indicates that line #5, column #16 defines `any` type.</span></span> <span data-ttu-id="20e3e-120">Это поможет вам найти ошибку.</span><span class="sxs-lookup"><span data-stu-id="20e3e-120">This helps you locate the error.</span></span>

<span data-ttu-id="20e3e-121">Чтобы обойти эту проблему, всегда определите тип переменной.</span><span class="sxs-lookup"><span data-stu-id="20e3e-121">To get around this issue, always define the type of the variable.</span></span> <span data-ttu-id="20e3e-122">Если вы не уверены в типе переменной, можно использовать [тип профсоюза.](https://www.typescriptlang.org/docs/handbook/unions-and-intersections.html)</span><span class="sxs-lookup"><span data-stu-id="20e3e-122">If you are uncertain about the type of a variable, you can use a [union type](https://www.typescriptlang.org/docs/handbook/unions-and-intersections.html).</span></span> <span data-ttu-id="20e3e-123">Это может быть полезно для переменных, которые держат значения, которые могут быть типа `Range` , или `string` `number` `boolean` `Range` (тип для значений является объединение `string | number | boolean` тех:).</span><span class="sxs-lookup"><span data-stu-id="20e3e-123">This can be useful for variables that hold `Range` values, which can be of type `string`, `number`, or `boolean` (the type for `Range` values is a union of those: `string | number | boolean`).</span></span>

### <a name="implicit-any"></a><span data-ttu-id="20e3e-124">подразумеваемый `any`</span><span class="sxs-lookup"><span data-stu-id="20e3e-124">Implicit `any`</span></span>

<span data-ttu-id="20e3e-125">ТипОписные переменные типы [могут быть неявно](https://www.typescriptlang.org/docs/handbook/type-inference.html) определены.</span><span class="sxs-lookup"><span data-stu-id="20e3e-125">TypeScript variable types can be [implicitly](https://www.typescriptlang.org/docs/handbook/type-inference.html) defined.</span></span> <span data-ttu-id="20e3e-126">Если компилятор TypeScript не в состоянии определить тип переменной (либо потому, что тип не определен явно или вывод типа не является возможным), то это неявное, `any` и вы получите ошибку компиляции времени.</span><span class="sxs-lookup"><span data-stu-id="20e3e-126">If the TypeScript compiler is unable to determine the type of a variable (either because type is not defined explicitly or type inference isn't possible), then it's an implicit `any` and you will receive a compilation-time error.</span></span>

<span data-ttu-id="20e3e-127">Наиболее распространенный случай на любой неявной `any` находится в переменной декларации, например `let value;` .</span><span class="sxs-lookup"><span data-stu-id="20e3e-127">The most common case on any implicit `any` is in a variable declaration, such as `let value;`.</span></span> <span data-ttu-id="20e3e-128">Есть два способа избежать этого:</span><span class="sxs-lookup"><span data-stu-id="20e3e-128">There are two ways to avoid this:</span></span>

* <span data-ttu-id="20e3e-129">Присвоить переменную неявно идентифицируемому типу `let value = 5;` `let value = workbook.getWorksheet();` (или).</span><span class="sxs-lookup"><span data-stu-id="20e3e-129">Assign the variable to an implicitly identifiable type (`let value = 5;` or `let value = workbook.getWorksheet();`).</span></span>
* <span data-ttu-id="20e3e-130">Явно ввех ввех переменной `let value: number;` ( )</span><span class="sxs-lookup"><span data-stu-id="20e3e-130">Explicitly type the variable (`let value: number;`)</span></span>

## <a name="no-inheriting-office-script-classes-or-interfaces"></a><span data-ttu-id="20e3e-131">Отсутствие Office классов или интерфейсов скрипта</span><span class="sxs-lookup"><span data-stu-id="20e3e-131">No inheriting Office Script classes or interfaces</span></span>

<span data-ttu-id="20e3e-132">Классы и интерфейсы, созданные в вашем Office, [не могут](https://www.typescriptlang.org/docs/handbook/classes.html#inheritance) расширить или реализовать Office классов или интерфейсов скриптов.</span><span class="sxs-lookup"><span data-stu-id="20e3e-132">Classes and interfaces that are created in your Office Script cannot [extend or implement](https://www.typescriptlang.org/docs/handbook/classes.html#inheritance) Office Scripts classes or interfaces.</span></span> <span data-ttu-id="20e3e-133">Другими словами, ничто в `ExcelScript` пространстве имен не может иметь подклассов или подповерхностных данных.</span><span class="sxs-lookup"><span data-stu-id="20e3e-133">In other words, nothing in the `ExcelScript` namespace can have subclasses or subinterfaces.</span></span>

## <a name="incompatible-typescript-functions"></a><span data-ttu-id="20e3e-134">Несовместимые функции TypeScript</span><span class="sxs-lookup"><span data-stu-id="20e3e-134">Incompatible TypeScript functions</span></span>

<span data-ttu-id="20e3e-135">Office API-файлы скриптов не могут быть использованы в следующих:</span><span class="sxs-lookup"><span data-stu-id="20e3e-135">Office Scripts APIs cannot be used in the following:</span></span>

* [<span data-ttu-id="20e3e-136">Функции генератора</span><span class="sxs-lookup"><span data-stu-id="20e3e-136">Generator functions</span></span>](https://developer.mozilla.org/docs/Web/JavaScript/Guide/Iterators_and_Generators#generator_functions)
* [<span data-ttu-id="20e3e-137">Array.sort</span><span class="sxs-lookup"><span data-stu-id="20e3e-137">Array.sort</span></span>](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Array/sort)

## <a name="eval-is-not-supported"></a><span data-ttu-id="20e3e-138">`eval` не поддерживается</span><span class="sxs-lookup"><span data-stu-id="20e3e-138">`eval` is not supported</span></span>

<span data-ttu-id="20e3e-139">Функция JavaScript [eval не](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/eval) поддерживается по соображениям безопасности.</span><span class="sxs-lookup"><span data-stu-id="20e3e-139">The JavaScript [eval function](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/eval) is not supported for security reasons.</span></span>

## <a name="restricted-identifers"></a><span data-ttu-id="20e3e-140">Ограниченные идентификаторы</span><span class="sxs-lookup"><span data-stu-id="20e3e-140">Restricted identifers</span></span>

<span data-ttu-id="20e3e-141">Следующие слова не могут быть использованы в качестве идентификаторов в скрипте.</span><span class="sxs-lookup"><span data-stu-id="20e3e-141">The following words can't be used as identifiers in a script.</span></span> <span data-ttu-id="20e3e-142">Это зарезервированные условия.</span><span class="sxs-lookup"><span data-stu-id="20e3e-142">They are reserved terms.</span></span>

* `Excel`
* `ExcelScript`
* `console`

## <a name="only-arrow-functions-in-array-callbacks"></a><span data-ttu-id="20e3e-143">Функции только стрелки в возвратах массивов</span><span class="sxs-lookup"><span data-stu-id="20e3e-143">Only arrow functions in array callbacks</span></span>

<span data-ttu-id="20e3e-144">Скрипты могут использовать функции [стрелки только при предоставлении](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Functions/Arrow_functions) аргументов обратного вызова для [методов Array.](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Array)</span><span class="sxs-lookup"><span data-stu-id="20e3e-144">Your scripts can only use [arrow functions](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Functions/Arrow_functions) when providing callback arguments for [Array](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Array) methods.</span></span> <span data-ttu-id="20e3e-145">Вы не можете передать эти методы какой-либо идентификатор или «традиционную» функцию.</span><span class="sxs-lookup"><span data-stu-id="20e3e-145">You cannot pass any sort of identifier or "traditional" function to these methods.</span></span>

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

## <a name="performance-warnings"></a><span data-ttu-id="20e3e-146">Предупреждения о производительности</span><span class="sxs-lookup"><span data-stu-id="20e3e-146">Performance warnings</span></span>

<span data-ttu-id="20e3e-147">Линтер редактора кода [предупреждает, если](https://wikipedia.org/wiki/Lint_(software)) у скрипта могут возникнуть проблемы с производительностью.</span><span class="sxs-lookup"><span data-stu-id="20e3e-147">The Code Editor's [linter](https://wikipedia.org/wiki/Lint_(software)) gives warnings if the script might have performance issues.</span></span> <span data-ttu-id="20e3e-148">Случаи и как обойти их задокументированы в [Улучшение производительности ваших Office скриптов](web-client-performance.md).</span><span class="sxs-lookup"><span data-stu-id="20e3e-148">The cases and how to work around them are documented in [Improve the performance of your Office Scripts](web-client-performance.md).</span></span>

## <a name="external-api-calls"></a><span data-ttu-id="20e3e-149">Внешние вызовы API</span><span class="sxs-lookup"><span data-stu-id="20e3e-149">External API calls</span></span>

<span data-ttu-id="20e3e-150">Дополнительную [информацию можно получить в Office внешних API-поддержки](external-calls.md) в скриптах.</span><span class="sxs-lookup"><span data-stu-id="20e3e-150">See [External API call support in Office Scripts](external-calls.md) for more information.</span></span>

## <a name="see-also"></a><span data-ttu-id="20e3e-151">См. также</span><span class="sxs-lookup"><span data-stu-id="20e3e-151">See also</span></span>

* [<span data-ttu-id="20e3e-152">Основные сведения о сценариях Office в Excel в Интернете</span><span class="sxs-lookup"><span data-stu-id="20e3e-152">Scripting fundamentals for Office Scripts in Excel on the web</span></span>](scripting-fundamentals.md)
* [<span data-ttu-id="20e3e-153">Улучшение производительности ваших Office скриптов</span><span class="sxs-lookup"><span data-stu-id="20e3e-153">Improve the performance of your Office Scripts</span></span>](web-client-performance.md)
