---
title: Использование макрофайлов в Power Automate потоках
description: Узнайте, как использовать макрофайлы или xlsm-файлы в Power Automate потоках.
ms.date: 09/01/2021
localization_priority: Normal
ms.openlocfilehash: 204eb8315f90c0ab0e20a34ec64517fbfbf304b1
ms.sourcegitcommit: 9472e78eb186ceffdaaceb2718d5074ce55a0e74
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/02/2021
ms.locfileid: "58866546"
---
# <a name="how-to-use-macro-files-in-power-automate-flows"></a>Использование макрофайлов в Power Automate потоках

[Соединитель Excel Online (Business)](https://flow.microsoft.com/connectors/shared_excelonlinebusiness/excel-online-business/) [](https://flow.microsoft.com/) в Power Automate обычно работает только с файлами в формате Microsoft Excel XML (.xlsx). Браузер файлов ограничивает ваш выбор .xlsx файлами внутри соединитетеля. Тем не менее, макрофайлы совместимы с действием скрипта **Run соединиттеля,** если используются метаданные файла.

В потоке используйте действие **Get File Metadata** из соединители OneDrive для бизнеса [или SharePoint](https://flow.microsoft.com/connectors/shared_sharepointonline/sharepoint/) файлов. [](https://flow.microsoft.com/connectors/shared_onedriveforbusiness/onedrive-for-business/) Действие **Сценарий Run** принимает эти метаданные как допустимый файл. Используйте *динамическое содержимое ID,* возвращаемое из действия метаданных **get file** в качестве аргумента "Файл" при запуске сценария. На следующем скриншоте показан поток, предоставляющий метаданные для файла под названием "Test Macro File.xlsm" для действия **скрипта Run.**

:::image type="content" source="../images/xlsm-in-power-automate.png" alt-text="Поток с действием метаданных Get file, передав метаданные макрофайла действию сценария Run.":::

> [!WARNING]
> Некоторые файлы xlsm, особенно файлы с ActiveX или элементами управления формами, могут не работать в Excel соединители. Убедитесь, что перед развертыванием решения необходимо протестировать.

## <a name="other-resources"></a>Другие ресурсы

[Просмотрите видео Sudhi Ramamurthy](https://youtu.be/o-H9BbywJQQ)на YouTube о том, как использовать файл xlsm в действии Run Script.