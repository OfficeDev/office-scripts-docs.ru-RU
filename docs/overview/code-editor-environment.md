---
title: Среда редактора кода сценариев Office
description: Предварительные требования и сведения о среде для сценариев Office в Excel в Интернете.
ms.date: 11/08/2022
ms.localizationpriority: medium
ms.openlocfilehash: a5a7601285553b1da4001a1870b6120f21bf5f2c
ms.sourcegitcommit: 7cadf2b637bf62874e43b6e595286101816662aa
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 11/09/2022
ms.locfileid: "68891255"
---
# <a name="office-scripts-code-editor-environment"></a>Среда редактора кода сценариев Office

Скрипты Office написаны на языке TypeScript или JavaScript и используют API JavaScript для сценариев Office для взаимодействия с книгой Excel. Редактор кода основан на Visual Studio Code, поэтому, если вы уже использовали эту среду, вы будете чувствовать себя как дома.

> [!TIP]
> Если вы знакомы с Visual Studio Code, теперь вы можете использовать его для написания скриптов. Чтобы опробовать эту функцию[, посетите Visual Studio Code для сценариев Office (предварительная версия).](../develop/vscode-for-scripts.md)

## <a name="scripting-language-typescript-or-javascript"></a>Язык сценариев: TypeScript или JavaScript

Сценарии Office написаны на языке [TypeScript](https://www.typescriptlang.org/docs/home.html), который является супермножеством [JavaScript](https://developer.mozilla.org/docs/Web/JavaScript). Средство записи действий создает код в TypeScript, а в документации по сценариям Office используется TypeScript. Так как TypeScript является надмножеством JavaScript, любой код скрипта, который вы пишете в JavaScript, будет работать нормально.

Скрипты Office в основном являются автономными фрагментами кода. Используется только небольшая часть функциональных возможностей TypeScript. Таким образом, вы можете редактировать скрипты, не изучая тонкости TypeScript. Редактор кода также обрабатывает установку, компиляцию и выполнение кода, поэтому вам не нужно беспокоиться ни о чем, кроме самого скрипта. Вы можете изучать язык и создавать скрипты без знаний программирования. Тем не менее, если вы не знакомы с программированием, рекомендуем ознакомиться с некоторыми основами, прежде чем приступить к работе со сценариями Office:

[!INCLUDE [Recommended coding resources](../includes/coding-basics-references.md)]

## <a name="office-scripts-javascript-api"></a>API JavaScript для сценариев Office

Скрипты Office используют специализированную версию API JavaScript для Office для [надстроек Office](/office/dev/add-ins/overview/index). Хотя в двух API есть сходства, не следует предполагать, что код может быть перенесен между двумя платформами. Различия между двумя платформами описаны в статье [Различия между сценариями Office и надстройками Office](../resources/add-ins-differences.md#apis) . Все API, доступные для скрипта, можно просмотреть в [справочной документации по API сценариев Office](/javascript/api/office-scripts/overview).

## <a name="external-library-support"></a>Поддержка внешней библиотеки

Скрипты Office не поддерживают использование внешних сторонних библиотек JavaScript. В настоящее время из скрипта нельзя вызывать библиотеку, кроме API сценариев Office. У вас по-прежнему есть доступ к любому [встроенному объекту JavaScript](../develop/javascript-objects.md), например [Math](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Math).

## <a name="intellisense"></a>Intellisense

IntelliSense — это набор функций редактора кода, которые помогают писать код. Он предоставляет документацию по автоматическому заполнению, выделению синтаксических ошибок и встроенному API.

IntelliSense предоставляет предложения по мере ввода, как в предложенном тексте в Excel. При нажатии клавиши TAB или ВВОД вставляется предложенный элемент. Запустите IntelliSense в текущем расположении курсора, нажав клавиши CTRL+ПРОБЕЛ. Эти рекомендации особенно полезны при выполнении метода. Сигнатура метода, отображаемая IntelliSense, содержит список необходимых аргументов, тип каждого аргумента, является ли заданный аргумент обязательным или необязательным, а также тип возвращаемого значения метода.

Наведите указатель мыши на метод, класс или другой объект кода, чтобы просмотреть дополнительные сведения. Наведите указатель мыши на синтаксическую ошибку или предложение кода, представленные красной или желтой полосой, чтобы увидеть рекомендации по устранению проблемы. Часто IntelliSense предоставляет параметр "Быстрое исправление" для автоматического изменения кода.

:::image type="content" source="../images/implicit-any-editor-message.png" alt-text="Сообщение об ошибке в тексте указателя мыши редактора кода с кнопкой &quot;Быстрое исправление&quot;.":::

Редактор кода сценариев Office использует тот же модуль IntelliSense, что и Visual Studio Code. Дополнительные сведения о функции см. в статье [Функции IntelliSense Visual Studio Code](https://code.visualstudio.com/docs/editor/intellisense#_intellisense-features).

## <a name="keyboard-shortcuts"></a>Сочетания клавиш

Большинство сочетаний клавиш для Visual Studio Code также работают в редакторе кода сценариев Office. Используйте следующие PDF-файлы, чтобы узнать о доступных параметрах и максимально эффективно использовать редактор кода:

- [Сочетания клавиш для macOS](https://code.visualstudio.com/shortcuts/keyboard-shortcuts-macos.pdf).
- [Сочетания клавиш для Windows](https://code.visualstudio.com/shortcuts/keyboard-shortcuts-windows.pdf).

## <a name="see-also"></a>См. также

- [Справочник API для сценариев Office](/javascript/api/office-scripts/overview)
- [Устранение неполадок в сценариях Office](../testing/troubleshooting.md)
- [Использование встроенных объектов JavaScript в сценариях Office](../develop/javascript-objects.md)
- [Visual Studio Code для сценариев Office (предварительная версия)](../develop/vscode-for-scripts.md)
