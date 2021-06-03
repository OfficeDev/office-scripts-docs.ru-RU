---
title: Office Среда редактора кода скриптов
description: Необходимые условия и сведения об среде для Office скриптов в Excel в Интернете.
ms.date: 05/27/2021
localization_priority: Normal
ms.openlocfilehash: 2334b0f98dfe03d97c35e6d1f54eeb45c06a134c
ms.sourcegitcommit: c75f71b8abde962e922927a18145dd1d9b361b05
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 06/03/2021
ms.locfileid: "52731143"
---
# <a name="office-scripts-code-editor-environment"></a>Office Среда редактора кода скриптов

Office Скрипты написаны в TypeScript или JavaScript и используют API javaScript Office скриптов для взаимодействия с Excel книгой. Редактор кода основан на Visual Studio Code, поэтому если вы использовали эту среду раньше, вы будете чувствовать себя как дома.

## <a name="scripting-language-typescript-or-javascript"></a>Язык скриптов: TypeScript или JavaScript

Сценарии Office написаны на языке [TypeScript](https://www.typescriptlang.org/docs/home.html), который является супермножеством [JavaScript](https://developer.mozilla.org/docs/Web/JavaScript). Регистратор действий создает код в TypeScript, а документация Office скриптов использует TypeScript. Так как TypeScript является суперсетью JavaScript, любой код скриптов, который вы пишете в JavaScript, будет работать нормально.

Office Скрипты — это в основном автономные фрагменты кода. Используется только малая часть функциональных возможностей TypeScript. Таким образом, вы можете изменить сценарии, не изучив тонкости TypeScript. Редактор кода также обрабатывает установку, компиляцию и выполнение кода, поэтому вам не нужно беспокоиться ни о чем, кроме самого сценария. Можно изучать язык и создавать сценарии без предыдущих знаний о программировании. Однако, если вы не знаете программирования, мы рекомендуем изучать некоторые основы перед тем, как Office скрипты:

[!INCLUDE [Recommended coding resources](../includes/coding-basics-references.md)]

## <a name="office-scripts-javascript-api"></a>Office API скриптов JavaScript

Office Скрипты используют специализированную версию API Office JavaScript для Office [надстройки](/office/dev/add-ins/overview/index). Хотя в двух API имеются сходства, не следует предполагать, что код можно портировать между двумя платформами. Различия между двумя платформами описаны в статье Различия между Office скриптами и [Office надстройки.](../resources/add-ins-differences.md#apis) Все API, доступные вашему сценарию, можно просмотреть в справочной документации [Office скриптов.](/javascript/api/office-scripts/overview)

## <a name="external-library-support"></a>Поддержка внешней библиотеки

Office Скрипты не поддерживают использование внешних сторонних библиотек JavaScript. В настоящее время нельзя вызывать любую библиотеку, кроме API Office скриптов. У вас по-прежнему есть доступ к любому встроенного [объекта JavaScript,](../develop/javascript-objects.md)например [Math](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Math).

## <a name="intellisense"></a>IntelliSense

IntelliSense — это набор функций редактора кода, которые помогают вам писать код. Он предоставляет автозаполнение, выделение ошибок синтаксиса и документацию по API.

IntelliSense при введите предложения, похожие на предложенный текст в Excel. Нажатие клавиши Tab или Enter вставляет предложенный элемент. Триггер IntelliSense в текущем расположении курсора, нажав клавиши Ctrl+Space. Эти предложения особенно полезны при выполнении метода. Подпись метода, отображаемая IntelliSense, содержит список необходимых ему аргументов, тип каждого аргумента, является ли данный аргумент обязательным или необязательным, а также тип возврата метода.

Наведите курсор на метод, класс или другой объект кода, чтобы увидеть дополнительные сведения. Наведите курсор над синтаксисной ошибкой или предложением кода, представленным красной или желтой строкой squiggly, чтобы увидеть предложения по устранению проблемы. Часто IntelliSense предоставляет параметр "Быстрое исправление", чтобы автоматически изменить код.

:::image type="content" source="../images/implicit-any-editor-message.png" alt-text="Сообщение об ошибке в тексте наведении редактора кода с кнопкой &quot;Быстрое исправление&quot;":::

Редактор Office скриптов использует тот же IntelliSense, что и Visual Studio Code. Дополнительные возможности этой функции [Visual Studio Code в IntelliSense.](https://code.visualstudio.com/docs/editor/intellisense#_intellisense-features)

## <a name="keyboard-shortcuts"></a>Сочетания клавиш

Большинство клавиш для Visual Studio Code также работают в редакторе Office скриптов. С помощью следующих PDF-адресов вы узнаете о доступных вариантах и получите большую часть редактора кода:

- [Клавиши клавиш для macOS](https://code.visualstudio.com/shortcuts/keyboard-shortcuts-macos.pdf).
- [Клавиши клавиш для Windows](https://code.visualstudio.com/shortcuts/keyboard-shortcuts-windows.pdf).

## <a name="see-also"></a>См. также

- [Справочник API для сценариев Office](/javascript/api/office-scripts/overview)
- [Устранение неполадок в сценариях Office](../testing/troubleshooting.md)
- [Использование встроенных объектов JavaScript в сценариях Office](../develop/javascript-objects.md)
