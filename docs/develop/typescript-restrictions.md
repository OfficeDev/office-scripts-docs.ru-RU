---
title: Ограничения TypeScript в сценариях Office
description: Особенности компиляторов TypeScript и linter, используемых редактором кода сценариев Office.
ms.date: 01/29/2021
localization_priority: Normal
ms.openlocfilehash: 41584ff23b333d17b2e267fdb3b0ec8741f3d203
ms.sourcegitcommit: df2b64603f91acb37bf95230efd538db0fbf9206
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 02/04/2021
ms.locfileid: "50099910"
---
# <a name="typescript-restrictions-in-office-scripts"></a>Ограничения TypeScript в сценариях Office

Сценарии Office используют язык TypeScript. По большей части любой код TypeScript или JavaScript будет работать в сценарии Office. Однако редактор кода на принудительное применение нескольких ограничений гарантирует, что ваш сценарий работает согласованно и по назначению с книгой Excel.

## <a name="no-any-type-in-office-scripts"></a>Нет типа "any" в сценариях Office

Типы [записи](https://www.typescriptlang.org/docs/handbook/typescript-in-5-minutes.html) необязательны в TypeScript, так как эти типы могут быть высмеяны. Однако сценарий Office требует, чтобы переменная не была [типом.](https://www.typescriptlang.org/docs/handbook/basic-types.html#any) Явные и `any` неявные не допускаются в сценарии Office. Эти случаи сообщаются как ошибки.

### <a name="explicit-any"></a>Explicit `any`

Нельзя явно объявить переменную типа в `any` сценариях Office (то `let someVariable: any;` есть). Тип `any` вызывает проблемы при обработке Excel. Например, необходимо `Range` знать, что значением является `string` , или `number` `boolean` . Вы получите ошибку времени компиляции (ошибку перед запуском сценария), если любая переменная явно определена в качестве типа `any` в сценарии.

![Явное сообщение в тексте наведении курсоров редактора кода](../images/explicit-any-editor-message.png)

![Явная любая ошибка в окне консоли](../images/explicit-any-error-message.png)

На снимке экрана выше `[5, 16] Explicit Any is not allowed` по указано, что #5 строка, #16 определяет `any` тип. Это помогает найти ошибку.

Чтобы обойти эту проблему, всегда определите тип переменной. Если вы не уверены в типе переменной, можно использовать [тип объединения.](https://www.typescriptlang.org/docs/handbook/unions-and-intersections.html) Это может быть полезно для переменных, которые удерживают значения, которые могут иметь тип , или (тип для значений это `Range` `string` объединение из `number` `boolean` `Range` них: `string | number | boolean` ).

### <a name="implicit-any"></a>Неявный `any`

Типы переменных TypeScript могут быть [неявно](( https://www.typescriptlang.org/docs/handbook/type-inference.html) определено. Если компилятору TypeScript не удается определить тип переменной (из-за того, что тип не определен явным образом или вывод типа невозможен), то это неявный параметр, и вы получите ошибку времени `any` компиляции.

Наиболее распространенный случай для любого неявного `any` параметра — объявление переменной, например `let value;` . Этого можно избежать двумя способами:

* Назначьте переменную неявно идентифицируемого типа `let value = 5;` `let value = workbook.getWorksheet();` (или).
* Явно введите переменную ( `let value: number;` )

## <a name="no-inheriting-office-script-classes-or-interfaces"></a>Нет наследования классов и интерфейсов сценариев Office

Классы и интерфейсы, созданные в скрипте Office, не могут расширять или реализовывать классы или интерфейсы сценариев [Office.](https://www.typescriptlang.org/docs/handbook/classes.html#inheritance) Другими словами, в пространстве имен не могут быть подклассы `ExcelScript` или подучества.

## <a name="incompatible-typescript-functions"></a>Несовместимые функции TypeScript

API сценариев Office нельзя использовать в следующих сценариях:

* [Функции генератора](https://developer.mozilla.org/docs/Web/JavaScript/Guide/Iterators_and_Generators#generator_functions)
* [Array.sort](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Array/sort)

## <a name="eval-is-not-supported"></a>`eval` не поддерживается

Функция [eval](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/eval) JavaScript не поддерживается из соображений безопасности.

## <a name="restricted-identifers"></a>Ограниченные отступы

Следующие слова нельзя использовать в качестве идентификаторов в скрипте. Это зарезервированные термины.

* `Excel`
* `ExcelScript`
* `console`

## <a name="performance-warnings"></a>Предупреждения о производительности

Линтер [редактора](https://wikipedia.org/wiki/Lint_(software)) кода выдает предупреждения, если у скрипта могут возникнуть проблемы с производительностью. Сценарии и их работа описаны в документе о повышении производительности [сценариев Office.](web-client-performance.md)

## <a name="external-api-calls"></a>Вызовы внешних API

Дополнительные сведения см. в службе поддержки вызовов внешнего [API в скриптах Office.](external-calls.md)

## <a name="see-also"></a>См. также

* [Основные сведения о сценариях Office в Excel в Интернете](scripting-fundamentals.md)
* [Повышение производительности сценариев Office](web-client-performance.md)
