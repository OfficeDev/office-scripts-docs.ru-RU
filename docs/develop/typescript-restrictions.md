---
title: Ограничения TypeScript в Office скриптах
description: Особенности компиляторов и подкладок TypeScript, используемых редактором кода Office скриптов.
ms.date: 02/05/2021
localization_priority: Normal
ms.openlocfilehash: 40eb6923d7b0c47dfeb4c846cdcc745e5d893c13
ms.sourcegitcommit: f7a7aebfb687f2a35dbed07ed62ff352a114525a
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 05/06/2021
ms.locfileid: "52232461"
---
# <a name="typescript-restrictions-in-office-scripts"></a>Ограничения TypeScript в Office скриптах

Office Скрипты используют язык TypeScript. По большей части любой код TypeScript или JavaScript будет работать в Office скрипта. Однако редактор кода соблюдает несколько ограничений, чтобы гарантировать, что сценарий работает последовательно и по назначению с Excel книгой.

## <a name="no-any-type-in-office-scripts"></a>Нет типа "любой" в Office скриптах

Типы [записи](https://www.typescriptlang.org/docs/handbook/typescript-in-5-minutes.html) необязательны в TypeScript, так как эти типы можно сделать вывод. Однако для Office скрипта требуется, чтобы переменная не была [типной.](https://www.typescriptlang.org/docs/handbook/basic-types.html#any) Явные и неявные не допускаются в `any` Office скрипте. Эти случаи сообщаются как ошибки.

### <a name="explicit-any"></a>Явный `any`

Нельзя явно объявить переменную типом в `any` Office Скрипты (то `let someVariable: any;` есть). Тип `any` вызывает проблемы при обработке Excel. Например, необходимо знать, что значение `Range` является `string` значением , или `number` `boolean` . Вы получите ошибку времени компиляции (ошибка перед запуском скрипта), если любая переменная явно определена как `any` тип сценария.

:::image type="content" source="../images/explicit-any-editor-message.png" alt-text="Явное сообщение &quot;любое&quot; в тексте наведении редактора кода":::

:::image type="content" source="../images/explicit-any-error-message.png" alt-text="Явные ошибки в окне консоли":::

На предыдущем `[5, 16] Explicit Any is not allowed` скриншоте указывается, что строка #5, столбец #16 определяет `any` тип. Это поможет найти ошибку.

Чтобы обойти эту проблему, всегда определите тип переменной. Если вы не уверены в типе переменной, можно использовать [тип union.](https://www.typescriptlang.org/docs/handbook/unions-and-intersections.html) Это может быть полезно для переменных, которые держат значения, которые могут быть типа , или (тип для значений является `Range` `string` `number` `boolean` `Range` союзом из них: `string | number | boolean` ).

### <a name="implicit-any"></a>Неявный `any`

Типы переменных TypeScript можно [неявно](https://www.typescriptlang.org/docs/handbook/type-inference.html) определить. Если компилятор TypeScript не может определить тип переменной (либо из-за того, что тип явно не определен, либо вывод типа невозможен), то это неявное значение, и вы получите ошибку времени `any` компиляции.

Наиболее распространенный случай для любого неявного `any` находится в переменной декларации, например `let value;` . Существует два способа избежать этого:

* Назначение переменной неявно идентифицируемого типа `let value = 5;` `let value = workbook.getWorksheet();` (или).
* Явно введите переменную ( `let value: number;` )

## <a name="no-inheriting-office-script-classes-or-interfaces"></a>Нет наследующих Office классов или интерфейсов скриптов

Классы и интерфейсы, созданные в Office скрипта, не могут расширять или [внедрять](https://www.typescriptlang.org/docs/handbook/classes.html#inheritance) Office скрипты или интерфейсы. Другими словами, ничто в пространстве имен не может `ExcelScript` иметь подклассов или подинтерфейсов.

## <a name="incompatible-typescript-functions"></a>Несовместимые функции TypeScript

Office API скриптов нельзя использовать в следующих следующих сценариях:

* [Функции генератора](https://developer.mozilla.org/docs/Web/JavaScript/Guide/Iterators_and_Generators#generator_functions)
* [Array.sort](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Array/sort)

## <a name="eval-is-not-supported"></a>`eval` не поддерживается

Функция [eval JavaScript](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/eval) не поддерживается из соображений безопасности.

## <a name="restricted-identifers"></a>Ограниченные identifers

Следующие слова нельзя использовать в качестве идентификаторов в скрипте. Это зарезервированные условия.

* `Excel`
* `ExcelScript`
* `console`

## <a name="only-arrow-functions-in-array-callbacks"></a>Только функции стрелки в вызовах массива

В скриптах можно использовать функции [стрелки](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Functions/Arrow_functions) только при предоставлении аргументов вызова для [методов Array.](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Array) Эти методы не могут передавать какие-либо идентификаторы или "традиционные" функции.

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

## <a name="performance-warnings"></a>Предупреждения о производительности

Подкладка редактора кода [дает](https://wikipedia.org/wiki/Lint_(software)) предупреждения, если у скрипта могут возникнуть проблемы с производительностью. Случаи и их работа описаны в документе Улучшение производительности [Office скриптов.](web-client-performance.md)

## <a name="external-api-calls"></a>Внешние вызовы API

Дополнительные сведения см. в Office службе поддержки [вызовов API.](external-calls.md)

## <a name="see-also"></a>См. также

* [Основные сведения о сценариях Office в Excel в Интернете](scripting-fundamentals.md)
* [Повышение производительности Office скриптов](web-client-performance.md)
