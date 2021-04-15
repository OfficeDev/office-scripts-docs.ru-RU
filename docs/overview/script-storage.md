---
title: Хранение и владение файлами Office Scripts
description: Сведения о том, как скрипты Office хранятся в Microsoft OneDrive и передаются между владельцами.
ms.date: 11/13/2020
localization_priority: Normal
ms.openlocfilehash: bd868c1dbfd0b33d3cd9fc4ee774c654d86f9b07
ms.sourcegitcommit: 45ffe3dbd2c834b78592ad35928cf8096f5e80bc
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 04/14/2021
ms.locfileid: "51755107"
---
# <a name="office-scripts-file-storage-and-ownership"></a>Хранение и владение файлами Office Scripts

Скрипты Office хранятся как **файлы .osts** в Microsoft OneDrive. Это позволяет скриптам существовать вне какой-либо конкретной книги. Параметры OneDrive контролируют общий доступ и разрешения для всех **файлов скрипта .osts;** независимо от параметров Excel.

## <a name="file-storage"></a>Хранилище файлов

Скрипты Office хранятся в OneDrive. Файлы **.osts** находятся в **папке /Documents/Office Scripts/.** Любые изменения, сделанные в этих **файлах .osts,** такие как переименование или удаление файлов, будут отражены в редакторе кода и галерее скриптов.

Скрипты, общие для одной из ваших книг, остаются в OneDrive создателя сценария. При запуске общего скрипта в Excel они не копируется ни в одну из локальных или папок OneDrive. Кнопка **Make a Copy** редактора кода сохраняет отдельную копию скрипта в OneDrive. Изменения в копии не влияют на исходный сценарий.

### <a name="script-folders"></a>Папки скриптов

Добавление папок в OneDrive помогает организовывать скрипты. Все папки **в разделе /Documents/Office Scripts/** отображаются в разделе **Мои скрипты** редактора кода. Обратите внимание, что эти папки невозможно создать или удалить с помощью редактора кода. Кроме того, скрипты не могут размещаться в папках или перемещаться между папками с помощью редактора кода.

:::image type="content" source="../images/script-folders.png" alt-text="Диалоговое окно New Script в редакторе кода показывает скрипты, содержащиеся в папках, как отображается в области задач.":::

## <a name="file-ownership-and-retention"></a>Владение и хранение файлов

Скрипты Office хранятся в OneDrive пользователя. Они следуют политикам хранения и удаления, указанным Microsoft OneDrive. Сведения о том, как обрабатывать сценарии, созданные и предоставленные пользователем, удаляемым из вашей организации, см. в статье [Хранение и удаление в OneDrive](/onedrive/retention-and-deletion).

## <a name="see-also"></a>См. также

- [Общий доступ к сценариям Office в веб-программе Excel](https://support.microsoft.com/office/sharing-office-scripts-in-excel-for-the-web-226eddbc-3a44-4540-acfe-fccda3d1122b)
- [Устранение неполадок в сценариях Office](../testing/troubleshooting.md)
- [Параметры сценариев Office в M365](https://support.office.com/article/office-scripts-settings-in-m365-19d3c51a-6ca2-40ab-978d-60fa49554dcf)
- [Отмена эффектов сценария Office](../testing/undo.md)
