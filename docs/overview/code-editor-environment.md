---
title: Office Среда редактора кода скриптов
description: Предпосылки и информация об окружающей среде для Office скриптов в Excel в Интернете.
ms.date: 05/10/2021
localization_priority: Normal
ms.openlocfilehash: aa54939826f8dda2a068df0f3fabf0fd3a2c842b
ms.sourcegitcommit: 4687693f02fc90a57ba30c461f35046e02e6f5fb
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 05/19/2021
ms.locfileid: "52545824"
---
# <a name="office-scripts-code-editor-environment"></a>Office Среда редактора кода скриптов

Office Сценарии написаны в TypeScript или JavaScript и используют Office Api JavaScript для взаимодействия с Excel книгой. Редактор кода основан на Visual Studio Code, так что если вы использовали эту среду раньше, вы будете чувствовать себя как дома.

## <a name="scripting-language-typescript-or-javascript"></a>Язык сценариев: TypeScript или JavaScript

Office Сценарии написаны в [TypeScript](https://www.typescriptlang.org/docs/home.html), который является суперсет [javaScript](https://developer.mozilla.org/docs/Web/JavaScript). Action Recorder генерирует код в TypeScript, а Office Scripts использует TypeScript. Поскольку TypeScript является суперсетом JavaScript, любой скрипт-код, который вы пишете в JavaScript, будет работать просто отлично.

Office Сценарии в основном являются автономными фрагментами кода. Используется лишь малая часть функциональности TypeScript. Таким образом, вы можете редактировать скрипты без необходимости изучать тонкости TypeScript. Редактор кода также обрабатывает установку, компиляцию и выполнение кода, так что вам не нужно беспокоиться ни о чем, кроме самого скрипта. Можно изучать язык и создавать скрипты без предыдущих знаний программирования. Однако, если вы нов для программирования, мы рекомендуем изузнать некоторые основы, прежде чем приступить к Office скриптов:

[!INCLUDE [Preview note](../includes/coding-basics-references.md)]

## <a name="office-scripts-javascript-api"></a>Office Сценарии JavaScript API

Office Скрипты используют специализированную версию Office JavaScript [для Office дополнительных приложений.](/office/dev/add-ins/overview/index) Хотя в двух API есть сходство, не следует предполагать, что код может быть портирован между двумя платформами. Различия между двумя платформами описаны в [различиях между Office и Office надстройки](../resources/add-ins-differences.md#apis) статьи. Вы можете просмотреть все API, доступные для вашего [скрипта, в справочной документации Office скриптов API.](/javascript/api/office-scripts/overview)

## <a name="external-library-support"></a>Внешняя поддержка библиотеки

Office Скрипты не поддерживают использование внешних сторонних библиотек JavaScript. В настоящее время нельзя звонить в какую-либо библиотеку, кроме Office API-скриптов из сценария. У вас все еще есть доступ к [любому встроенному объекту JavaScript,](../develop/javascript-objects.md)например, [к математике.](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Math)

## <a name="intellisense"></a>IntelliSense

IntelliSense функция редактора кода, которая помогает предотвратить опечатки и ошибки синтаксиса при редактировании сценария. Он отображает возможные имена объектов и поля при ввеся, а также вко— вко— вко— вко— в соответствующую документацию для каждого API.

Редактор Excel кода использует тот же движок IntelliSense, что и Visual Studio Code. Чтобы узнать больше об этой функции, [Visual Studio Code в IntelliSense функции](https://code.visualstudio.com/docs/editor/intellisense#_intellisense-features).

## <a name="keyboard-shortcuts"></a>Сочетания клавиш

Большинство ярлыков клавиатуры для Visual Studio Code также работают в Office Code Editor. Используйте следующие PDF-файлы, чтобы узнать о доступных опциях и получить большую часть из редактора кода:

- [Клавишные ярлыки для macOS](https://code.visualstudio.com/shortcuts/keyboard-shortcuts-macos.pdf).
- [Клавиатура ярлыки для Windows](https://code.visualstudio.com/shortcuts/keyboard-shortcuts-windows.pdf).

## <a name="see-also"></a>См. также

- [Справочник API для сценариев Office](/javascript/api/office-scripts/overview)
- [Устранение неполадок в сценариях Office](../testing/troubleshooting.md)
- [Использование встроенных объектов JavaScript в сценариях Office](../develop/javascript-objects.md)
