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
# <a name="typescript-restrictions-in-office-scripts"></a>Ограничения TypeScript в Office скриптах

Office Скрипты используют язык TypeScript. По большей части любой код TypeScript или JavaScript будет работать в Office скриптах. Тем не менее, редактор Кода применяет несколько ограничений для обеспечения того, чтобы ваш скрипт работал последовательно и по назначению с вашей Excel книгой.

## <a name="no-any-type-in-office-scripts"></a>Нет типа "любой" в Office скриптах

Типы [писания](https://www.typescriptlang.org/docs/handbook/typescript-in-5-minutes.html) не являются обязательными в TypeScript, потому что типы могут быть выведены. Тем не Office, что скрипты требуют, чтобы переменная не может быть [типа любого](https://www.typescriptlang.org/docs/handbook/basic-types.html#any). Как явные, так `any` и неявные не допускаются Office скриптов. Эти случаи сообщаются как ошибки.

### <a name="explicit-any"></a>явный `any`

Вы не можете прямо объявить переменную `any` типом в Office (то `let someVariable: any;` есть). Тип `any` вызывает проблемы при обработке Excel. Например, `Range` нужно знать, что значение `string` `number` является, или `boolean` . Вы получите ошибку времени компиляции (ошибка до запуска скрипта), если какая-либо переменная явно `any` определена как тип скрипта.

:::image type="content" source="../images/explicit-any-editor-message.png" alt-text="Явное сообщение «любого» в тексте наведении редактора Кода":::

:::image type="content" source="../images/explicit-any-error-message.png" alt-text="Явная ошибка «любого» в окне консоли":::

На предыдущем скриншоте `[5, 16] Explicit Any is not allowed` указывается, что #5, #16 определяет `any` тип. Это поможет вам найти ошибку.

Чтобы обойти эту проблему, всегда определите тип переменной. Если вы не уверены в типе переменной, можно использовать [тип профсоюза.](https://www.typescriptlang.org/docs/handbook/unions-and-intersections.html) Это может быть полезно для переменных, которые держат значения, которые могут быть типа `Range` , или `string` `number` `boolean` `Range` (тип для значений является объединение `string | number | boolean` тех:).

### <a name="implicit-any"></a>подразумеваемый `any`

ТипОписные переменные типы [могут быть неявно](https://www.typescriptlang.org/docs/handbook/type-inference.html) определены. Если компилятор TypeScript не в состоянии определить тип переменной (либо потому, что тип не определен явно или вывод типа не является возможным), то это неявное, `any` и вы получите ошибку компиляции времени.

Наиболее распространенный случай на любой неявной `any` находится в переменной декларации, например `let value;` . Есть два способа избежать этого:

* Присвоить переменную неявно идентифицируемому типу `let value = 5;` `let value = workbook.getWorksheet();` (или).
* Явно ввех ввех переменной `let value: number;` ( )

## <a name="no-inheriting-office-script-classes-or-interfaces"></a>Отсутствие Office классов или интерфейсов скрипта

Классы и интерфейсы, созданные в вашем Office, [не могут](https://www.typescriptlang.org/docs/handbook/classes.html#inheritance) расширить или реализовать Office классов или интерфейсов скриптов. Другими словами, ничто в `ExcelScript` пространстве имен не может иметь подклассов или подповерхностных данных.

## <a name="incompatible-typescript-functions"></a>Несовместимые функции TypeScript

Office API-файлы скриптов не могут быть использованы в следующих:

* [Функции генератора](https://developer.mozilla.org/docs/Web/JavaScript/Guide/Iterators_and_Generators#generator_functions)
* [Array.sort](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Array/sort)

## <a name="eval-is-not-supported"></a>`eval` не поддерживается

Функция JavaScript [eval не](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/eval) поддерживается по соображениям безопасности.

## <a name="restricted-identifers"></a>Ограниченные идентификаторы

Следующие слова не могут быть использованы в качестве идентификаторов в скрипте. Это зарезервированные условия.

* `Excel`
* `ExcelScript`
* `console`

## <a name="only-arrow-functions-in-array-callbacks"></a>Функции только стрелки в возвратах массивов

Скрипты могут использовать функции [стрелки только при предоставлении](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Functions/Arrow_functions) аргументов обратного вызова для [методов Array.](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Array) Вы не можете передать эти методы какой-либо идентификатор или «традиционную» функцию.

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

Линтер редактора кода [предупреждает, если](https://wikipedia.org/wiki/Lint_(software)) у скрипта могут возникнуть проблемы с производительностью. Случаи и как обойти их задокументированы в [Улучшение производительности ваших Office скриптов](web-client-performance.md).

## <a name="external-api-calls"></a>Внешние вызовы API

Дополнительную [информацию можно получить в Office внешних API-поддержки](external-calls.md) в скриптах.

## <a name="see-also"></a>См. также

* [Основные сведения о сценариях Office в Excel в Интернете](scripting-fundamentals.md)
* [Улучшение производительности ваших Office скриптов](web-client-performance.md)
