---
title: Ограничения TypeScript в Office скриптах
description: Особенности компиляторов и подкладок TypeScript, используемых редактором кода Office скриптов.
ms.date: 11/09/2021
ms.localizationpriority: medium
ms.openlocfilehash: 7b67ccb4898823100e890aa5c8c0332d28a4522b
ms.sourcegitcommit: ddbb1c66d627ffabbfc3b938d6e25cf6fe3cc13f
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 11/12/2021
ms.locfileid: "60924105"
---
# <a name="typescript-restrictions-in-office-scripts"></a>Ограничения TypeScript в Office скриптах

Office скрипты используют язык TypeScript. По большей части любой код TypeScript или JavaScript будет работать в Office скриптах. Однако редактор кода соблюдает несколько ограничений, чтобы гарантировать, что сценарий работает последовательно и по назначению с Excel книгой.

## <a name="no-any-type-in-office-scripts"></a>Нет типа "любой" в Office скриптах

Типы [записи](https://www.typescriptlang.org/docs/handbook/typescript-in-5-minutes.html) необязательны в TypeScript, так как эти типы можно сделать вывод. Однако для Office скриптов требуется, чтобы переменная не была [типной.](https://www.typescriptlang.org/docs/handbook/basic-types.html#any) Явные и неявные не допускаются `any` в Office скриптах. Эти случаи сообщаются как ошибки.

### <a name="explicit-any"></a>Явный `any`

Нельзя явно объявить переменную типом в `any` Office Скрипты (то `let value: any;` есть). Тип `any` вызывает проблемы при обработке Excel. Например, необходимо знать, что значение `Range` является `string` значением , или `number` `boolean` . Вы получите ошибку времени компиляции (ошибка перед запуском скрипта), если любая переменная явно определена как `any` тип сценария.

:::image type="content" source="../images/explicit-any-editor-message.png" alt-text="Явное сообщение &quot;любое&quot; в тексте наведении редактора кода.":::

:::image type="content" source="../images/explicit-any-error-message.png" alt-text="Явная ошибка &quot;любая&quot; в окне консоли.":::

На предыдущем скриншоте `[2, 14] Explicit Any is not allowed` указывается, что строка #2 столбец #14 определяет `any` тип. Это поможет найти ошибку.

Чтобы обойти эту проблему, всегда определите тип переменной. Если вы не уверены в типе переменной, можно использовать [тип union.](https://www.typescriptlang.org/docs/handbook/unions-and-intersections.html) Это может быть полезно для переменных, которые держат значения, которые могут быть типа , или (тип для значений является `Range` `string` `number` `boolean` `Range` союзом из них: `string | number | boolean` ).

### <a name="implicit-any"></a>Неявный `any`

Типы переменных TypeScript можно [неявно](https://www.typescriptlang.org/docs/handbook/type-inference.html) определить. Если компилятор TypeScript не может определить тип переменной (либо из-за того, что тип явно не определен, либо вывод типа невозможен), то это неявное значение, и вы получите ошибку времени `any` компиляции.

:::image type="content" source="../images/implicit-any-editor-message.png" alt-text="Неявное сообщение &quot;любое&quot; в тексте наведении редактора кода.":::

Наиболее распространенный случай для любого неявного `any` находится в переменной декларации, например `let value;` . Существует два способа избежать этого:

* Назначение переменной неявно идентифицируемого типа `let value = 5;` `let value = workbook.getWorksheet();` (или).
* Явно введите переменную ( `let value: number;` )

## <a name="no-inheriting-office-script-classes-or-interfaces"></a>Нет наследующих Office классов или интерфейсов скриптов

Классы и интерфейсы, созданные в Office скрипта, не могут расширять или [внедрять](https://www.typescriptlang.org/docs/handbook/classes.html#inheritance) Office скрипты или интерфейсы. Другими словами, ничто в пространстве имен не может `ExcelScript` иметь подклассов или подинтерфейсов.

## <a name="incompatible-typescript-functions"></a>Несовместимые функции TypeScript

Office API скриптов нельзя использовать в следующих ниже.

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

## <a name="unions-of-excelscript-types-and-user-defined-types-arent-supported"></a>Союзы типов и типов, определенных `ExcelScript` пользователем, не поддерживаются

Office скрипты преобразуются во время запуска из синхронных в асинхронные блоки кода. Связь с книгой с помощью [обещаний](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise) скрыта от создателя сценария. Это преобразование не поддерживает типы [профсоюзов,](https://www.typescriptlang.org/docs/handbook/2/everyday-types.html#union-types) которые включают типы и `ExcelScript` типы, определенные пользователем. В этом случае сценарий возвращается в сценарий, но компилятор Office скрипта не ожидает его, и создатель сценария не может взаимодействовать с `Promise` `Promise` .

В следующем примере кода показан неподтверченный союз между `ExcelScript.Table` пользовательским `MyTable` интерфейсом.

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  const selectedSheet = workbook.getActiveWorksheet();

  // This union is not supported.
  const tableOrMyTable: ExcelScript.Table | MyTable = selectedSheet.getTables()[0];

  // `getName` returns a promise that can't be resolved by the script.
  const name = tableOrMyTable.getName();

  // This logs "{}" instead of the table name.
  console.log(name);
}

interface MyTable {
  getName(): string
}
```

## <a name="constructors-dont-support-office-scripts-apis-and-console-statements"></a>Конструкторы не поддерживают API Office скриптов и `console` утверждения

`console`утверждения и многие API Office скриптов требуют синхронизации с Excel книгой. Эти синхронизации используют `await` утверждения в скомпилизированной версии сценария. `await`не поддерживается в [конструкторах.](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Classes/constructor) Если вам нужны классы с конструкторами, не Office API или утверждения скриптов `console` в этих блоках кода.

В следующем примере кода демонстрируется этот сценарий. Он создает ошибку, которая говорит `failed to load [code] [library]` .

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  class MyClass {
    constructor() {
      // Console statements and Office Scripts APIs aren't supported in constructors.
      console.log("This won't print.");
    }
  }

  let test = new MyClass();
}
```

## <a name="performance-warnings"></a>Предупреждения о производительности

Подкладка редактора кода [дает](https://wikipedia.org/wiki/Lint_(software)) предупреждения, если у скрипта могут возникнуть проблемы с производительностью. Случаи и их работа описаны в документе Улучшение производительности [Office скриптов.](web-client-performance.md)

## <a name="external-api-calls"></a>Внешние вызовы API

Дополнительные сведения см. в Office службе поддержки [вызовов API.](external-calls.md)

## <a name="see-also"></a>См. также

* [Основные сведения о сценариях Office в Excel для Интернета](scripting-fundamentals.md)
* [Повышение производительности Office скриптов](web-client-performance.md)
