---
title: Использование файлов с поддержкой макроса в Power Automate потоках
description: Узнайте, как использовать файлы с поддержкой макроса или Xlsm-файлы в Power Automate потоках.
ms.date: 03/24/2022
ms.localizationpriority: medium
ms.openlocfilehash: 9f2ecefe9fb97d1c5514ddb52c3cbcd0596df426
ms.sourcegitcommit: 7023b9e23499806901a5ecf8ebc460b76887cca6
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/31/2022
ms.locfileid: "64585746"
---
# <a name="how-to-use-macro-enabled-files-in-power-automate-flows"></a>Использование файлов с поддержкой макроса в Power Automate потоках

Вы можете интегрировать файлы xlsm с потоком Power Automate. Это позволяет приступить к преобразованию существующих решений автоматизации в веб-форматы. Обратите внимание, что макрос, содержащиеся в файлах XSLM, невозможно выполнить через Power Automate. Только Office скрипты включены там.

[Соединителю Excel Online (Business)](https://flow.microsoft.com/connectors/shared_excelonlinebusiness/excel-online-business/) в Power Automate обычно ограничиваются файлы в формате Microsoft Excel таблицы open XML (.xlsx).[](https://flow.microsoft.com/) Браузер файлов позволяет выбирать только .xlsx файлы. Однако файлы с поддержкой макроса совместимы с действием скрипта **Run соединители** , если используются метаданные файла.

В потоке используйте действие **Get File Metadata** с OneDrive для бизнеса [или SharePoint](https://flow.microsoft.com/connectors/shared_sharepointonline/sharepoint/) соединители.[](https://flow.microsoft.com/connectors/shared_onedriveforbusiness/onedrive-for-business/) Действие **Сценарий Run** принимает эти метаданные как допустимый файл. Используйте *динамическое содержимое ID* , возвращаемое из действия метаданных **get file** в качестве аргумента "Файл" при запуске сценария. На следующем скриншоте показан поток, предоставляющий метаданные для файла под названием "Test Macro File.xlsm" для действия **скрипта Run** .

:::image type="content" source="../images/xlsm-in-power-automate.png" alt-text="Поток с действием метаданных Get file, передав метаданные макрофайла действию сценария Run.":::

> [!WARNING]
> Некоторые файлы xlsm, особенно файлы с ActiveX или элементами управления формами, могут не работать в Excel сетевом соединителене. Убедитесь, что перед развертыванием решения необходимо протестировать.

## <a name="other-resources"></a>Другие ресурсы

[Просмотрите видео Sudhi Ramamurthy на YouTube](https://youtu.be/o-H9BbywJQQ) о том, как использовать файл xlsm в действии Run Script.
