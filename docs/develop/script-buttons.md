---
title: Запуск Office сценариев в Excel с помощью кнопок
description: Добавьте кнопки в книги, которые Office скрипты в Excel.
ms.topic: overview
ms.date: 05/09/2022
ms.localizationpriority: medium
ms.openlocfilehash: fde34d62f9abe897a8b93195ab37a75cfc73f619
ms.sourcegitcommit: 34c7740c9bff0e4c7426e01029f967724bfee566
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 05/13/2022
ms.locfileid: "65393686"
---
# <a name="run-office-scripts-in-excel-with-buttons"></a>Запуск Office сценариев в Excel с помощью кнопок

Помогите коллегам находить и запускать сценарии путем добавления кнопок сценариев в книгу.

:::image type="content" source="../images/run-from-button.png" alt-text="Кнопка на листе, которая запускает сценарий при нажатии.":::

## <a name="create-script-buttons"></a>Создание кнопок сценариев

В любом сценарии перейдите в меню "Дополнительные параметры" **(...)** на странице сведений о скрипте или в области задач редактора кода и нажмите кнопку **"Добавить"**. В результате в книге будет создана кнопка, которая запускает связанный сценарий при нажатии. Она также предоставляет общий доступ к сценарию в книге, поэтому каждый пользователь, у кого есть разрешения на запись в книге, может использовать вашу удобную автоматизацию.

На следующем снимке экрана показана страница сведений о скрипте для скрипта сводная таблица с выделенным параметром "**Добавить** кнопку" в меню "Дополнительные параметры" **(...)** 

:::image type="content" source="../images/add-button.png" alt-text="Параметр &quot;Добавить кнопку&quot; в меню страницы сведений о скрипте.":::

## <a name="remove-script-buttons"></a>Удаление кнопок скрипта

Чтобы прекратить общий доступ к скрипту с помощью кнопки, перейдите в меню "Дополнительные **параметры" (...)** на странице сведений о скрипте и выберите " **Остановить общий доступ"**. Это удалит все кнопки, которые запускают сценарий. Удаление одной кнопки удаляет сценарий из этой одной кнопки, даже если операция отменена или кнопка вырезана и вставлена.

## <a name="script-buttons-with-excel-on-windows"></a>Кнопки скрипта с Excel на Windows

Эти кнопки сценариев также работают в Windows. Создайте кнопку в Excel в Интернете и пользователи на Windows могут запустить сценарий, нажав кнопку. Обратите внимание, что вы не можете изменять скрипты в Excel на Windows. Изменять скрипты можно только в Excel в Интернете.

Некоторые Office скриптов могут не поддерживаться Excel в Windows, особенно в более старых сборках. К ним относятся новые API и API для функций, доступных только в Интернете. Если скрипт содержит неподдерживаемые API, скрипт не выполняется, а вместо этого в области задач  "Состояние выполнения скрипта" отображается предупреждающее сообщение: "Этот сценарий в настоящее время должен выполняться в Excel для Интернета. Откройте книгу в браузере и повторите попытку или обратитесь за помощью к владельцу скрипта".  

> [!IMPORTANT]
> Для работы с [кнопками сценариев требуется webView2](/deployoffice/webview2-install) Excel на Windows. Он устанавливается по умолчанию с последними версиями Excel desktop, но если вы не можете нажать кнопки сценариев, перейдите на страницу скачивания среды [выполнения WebView2](https://developer.microsoft.com/en-us/microsoft-edge/webview2/#download-section) и скачайте подсистему браузера.