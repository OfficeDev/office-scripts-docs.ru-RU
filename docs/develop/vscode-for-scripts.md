---
title: Visual Studio Code для сценариев Office (предварительная версия)
description: Настройка редактора кода сценариев Office для подключения к VS Code в Интернете.
ms.date: 11/04/2022
ms.localizationpriority: medium
ms.openlocfilehash: fd9dd417610c8ad64fbd3fc50048ce56afdb4e28
ms.sourcegitcommit: 7cadf2b637bf62874e43b6e595286101816662aa
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 11/09/2022
ms.locfileid: "68892046"
---
# <a name="visual-studio-code-for-office-scripts-preview"></a>Visual Studio Code для сценариев Office (предварительная версия)

[Visual Studio Code для Интернета](https://vscode.dev/) позволяет пользователям редактировать что-либо из любого места. Подключите интерфейс сценариев Office к этому популярному редактору кода, чтобы начать выполнение скриптов за пределами книги.

:::image type="content" source="../images/vscode-script-editor.png" alt-text="Окно Excel в Интернете с открытым редактором кода рядом с окном VS Code в Интернете с открытым скриптом.":::

Visual Studio Code имеет ряд преимуществ по сравнению со встроенным редактором кода.

- Полноэкранное редактирование! Вашему скрипту больше не нужно делиться пространством на экране с книгой.
- Одновременное редактирование нескольких скриптов! Быстрое переключение между скриптами для совместного использования кода из других служб автоматизации.
- Расширения! Используйте избранные расширения VS Code для проверки орфографии, форматирования и других действий, которые помогут вам выполнить работу.

> [!NOTE]
> Эта функция находится в состоянии предварительной версии. Он может быть изменен на основе отзывов. Если у вас возникли проблемы, сообщите о них с помощью кнопки **Отзыв** в Excel. Ниже приведены известные проблемы с текущей версией компонента.
>
> - Visual Studio Code можно подключить к сценариям Office только через Excel в Интернете.
> - Это подключение к скриптам Office доступно только для клиентов Excel на английском языке.

## <a name="connect-visual-studio-code-to-office-scripts"></a>Подключение Visual Studio Code к сценариям Office

Выполните эти однократные действия, чтобы подключить Visual Studio Code и Excel в Интернете.

1. Откройте **редактор кода** сценариев Office.
2. В меню **Дополнительные параметры (...)** выберите **Параметры редактора**.
3. Выберите **(предварительная версия) Visual Studio Code подключение**.

:::image type="content" source="../images/vscode-enable-option.png" alt-text="Область задач &quot;Параметры редактора&quot; с флажком Visual Studio Code подключения.":::

Теперь вы можете изменять и запускать скрипты из Visual Studio Code. В любом сценарии перейдите в меню **Дополнительные параметры (...)** и выберите **Открыть в VS Code**.

:::image type="content" source="../images/vscode-open-option.png" alt-text="Параметр Открыть в VS Code выбирается в списке рядом с открытым скриптом.":::

## <a name="see-also"></a>См. также

- [Среда редактора кода сценариев Office](../overview/code-editor-environment.md)
- [Visual Studio Code для Интернета (документация)](https://code.visualstudio.com/docs/editor/vscode-web)
