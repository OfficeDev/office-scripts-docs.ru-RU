---
title: Office Хранение и владение файлами скриптов
description: Сведения о том, Office скрипты хранятся в Microsoft OneDrive и передаются между владельцами.
ms.date: 11/13/2020
localization_priority: Normal
ms.openlocfilehash: 47b732399c3068bea78b027e01324bbd73a83bc7
ms.sourcegitcommit: f7a7aebfb687f2a35dbed07ed62ff352a114525a
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 05/06/2021
ms.locfileid: "52232531"
---
# <a name="office-scripts-file-storage-and-ownership"></a>Office Хранение и владение файлами скриптов

Office Скрипты хранятся как **файлы .osts** в Microsoft OneDrive. Это позволяет скриптам существовать вне какой-либо конкретной книги. Параметры OneDrive управления общим доступом и разрешениями для всех файлов **скрипта .osts;** независимо от Excel параметров.

## <a name="file-storage"></a>Хранилище файлов

Вы Office, что скрипты хранятся в OneDrive. Файлы **.osts** находятся в **папке /Documents/Office scripts/.** Любые изменения, сделанные в этих **файлах .osts,** такие как переименование или удаление файлов, будут отражены в редакторе кода и галерее скриптов.

Скрипты, общие для одной из ваших книг, остаются в OneDrive. Они не копируется ни в одной из локальных или OneDrive папок при запуске общего скрипта в Excel. Кнопка **Make a Copy** редактора кода сохраняет отдельную копию сценария в OneDrive. Изменения в копии не влияют на исходный сценарий.

### <a name="script-folders"></a>Папки скриптов

Добавление папок в OneDrive помогает организовать скрипты. Все папки в **разделе /Documents/Office Scripts/отображаются** в разделе **Мои** скрипты редактора кода. Обратите внимание, что эти папки невозможно создать или удалить с помощью редактора кода. Кроме того, скрипты не могут размещаться в папках или перемещаться между папками с помощью редактора кода.

:::image type="content" source="../images/script-folders.png" alt-text="Диалоговое окно New Script в редакторе кода, отображаемом сценарии, содержащиеся в папках, как отображается в области задач":::

## <a name="file-ownership-and-retention"></a>Владение и хранение файлов

Office Скрипты хранятся в OneDrive. Они следуют политикам хранения и удаления, указанным Microsoft OneDrive. Сведения о том, как обрабатывать сценарии, созданные и предоставленные пользователем, удаляемым из вашей организации, см. в статье [Хранение и удаление в OneDrive](/onedrive/retention-and-deletion).

## <a name="see-also"></a>См. также

- [Общий доступ к сценариям Office в веб-программе Excel](https://support.microsoft.com/office/sharing-office-scripts-in-excel-for-the-web-226eddbc-3a44-4540-acfe-fccda3d1122b)
- [Устранение неполадок в сценариях Office](../testing/troubleshooting.md)
- [Параметры сценариев Office в M365](https://support.office.com/article/office-scripts-settings-in-m365-19d3c51a-6ca2-40ab-978d-60fa49554dcf)
- [Отмена эффектов сценария Office](../testing/undo.md)
