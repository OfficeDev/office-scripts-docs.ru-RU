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
# <a name="office-scripts-file-storage-and-ownership"></a><span data-ttu-id="38f85-103">Хранение и владение файлами Office Scripts</span><span class="sxs-lookup"><span data-stu-id="38f85-103">Office Scripts file storage and ownership</span></span>

<span data-ttu-id="38f85-104">Скрипты Office хранятся как **файлы .osts** в Microsoft OneDrive.</span><span class="sxs-lookup"><span data-stu-id="38f85-104">Office Scripts are stored as **.osts** files in your Microsoft OneDrive.</span></span> <span data-ttu-id="38f85-105">Это позволяет скриптам существовать вне какой-либо конкретной книги.</span><span class="sxs-lookup"><span data-stu-id="38f85-105">This allows your scripts to exist outside any particular workbook.</span></span> <span data-ttu-id="38f85-106">Параметры OneDrive контролируют общий доступ и разрешения для всех **файлов скрипта .osts;** независимо от параметров Excel.</span><span class="sxs-lookup"><span data-stu-id="38f85-106">Your OneDrive settings control the shared access and permissions for all script **.osts** files; independent of any Excel settings.</span></span>

## <a name="file-storage"></a><span data-ttu-id="38f85-107">Хранилище файлов</span><span class="sxs-lookup"><span data-stu-id="38f85-107">File storage</span></span>

<span data-ttu-id="38f85-108">Скрипты Office хранятся в OneDrive.</span><span class="sxs-lookup"><span data-stu-id="38f85-108">You Office Scripts are stored in your OneDrive.</span></span> <span data-ttu-id="38f85-109">Файлы **.osts** находятся в **папке /Documents/Office Scripts/.**</span><span class="sxs-lookup"><span data-stu-id="38f85-109">The **.osts** files are found in the **/Documents/Office Scripts/** folder.</span></span> <span data-ttu-id="38f85-110">Любые изменения, сделанные в этих **файлах .osts,** такие как переименование или удаление файлов, будут отражены в редакторе кода и галерее скриптов.</span><span class="sxs-lookup"><span data-stu-id="38f85-110">Any edits made to these **.osts** files, such as renaming or deleting files, will be reflected in the Code Editor and Script Gallery.</span></span>

<span data-ttu-id="38f85-111">Скрипты, общие для одной из ваших книг, остаются в OneDrive создателя сценария.</span><span class="sxs-lookup"><span data-stu-id="38f85-111">Scripts that are shared with one of your workbooks remain in the script creator's OneDrive.</span></span> <span data-ttu-id="38f85-112">При запуске общего скрипта в Excel они не копируется ни в одну из локальных или папок OneDrive.</span><span class="sxs-lookup"><span data-stu-id="38f85-112">They are not copied to any of your local or OneDrive folders when you run the shared script in Excel.</span></span> <span data-ttu-id="38f85-113">Кнопка **Make a Copy** редактора кода сохраняет отдельную копию скрипта в OneDrive.</span><span class="sxs-lookup"><span data-stu-id="38f85-113">The **Make a Copy** button of the Code Editor saves a separate copy of the script in your OneDrive.</span></span> <span data-ttu-id="38f85-114">Изменения в копии не влияют на исходный сценарий.</span><span class="sxs-lookup"><span data-stu-id="38f85-114">Changes to the copy don't affect the original script.</span></span>

### <a name="script-folders"></a><span data-ttu-id="38f85-115">Папки скриптов</span><span class="sxs-lookup"><span data-stu-id="38f85-115">Script folders</span></span>

<span data-ttu-id="38f85-116">Добавление папок в OneDrive помогает организовывать скрипты.</span><span class="sxs-lookup"><span data-stu-id="38f85-116">Adding folders to your OneDrive helps keep your scripts organized.</span></span> <span data-ttu-id="38f85-117">Все папки **в разделе /Documents/Office Scripts/** отображаются в разделе **Мои скрипты** редактора кода.</span><span class="sxs-lookup"><span data-stu-id="38f85-117">Any folders under **/Documents/Office Scripts/** are displayed under the **My Scripts** section of the Code Editor.</span></span> <span data-ttu-id="38f85-118">Обратите внимание, что эти папки невозможно создать или удалить с помощью редактора кода.</span><span class="sxs-lookup"><span data-stu-id="38f85-118">Please note that these folders cannot be created or deleted by using the Code Editor.</span></span> <span data-ttu-id="38f85-119">Кроме того, скрипты не могут размещаться в папках или перемещаться между папками с помощью редактора кода.</span><span class="sxs-lookup"><span data-stu-id="38f85-119">Likewise, scripts cannot be placed in folders, or moved across folders by using the Code Editor.</span></span>

:::image type="content" source="../images/script-folders.png" alt-text="Диалоговое окно New Script в редакторе кода показывает скрипты, содержащиеся в папках, как отображается в области задач.":::

## <a name="file-ownership-and-retention"></a><span data-ttu-id="38f85-121">Владение и хранение файлов</span><span class="sxs-lookup"><span data-stu-id="38f85-121">File ownership and retention</span></span>

<span data-ttu-id="38f85-122">Скрипты Office хранятся в OneDrive пользователя.</span><span class="sxs-lookup"><span data-stu-id="38f85-122">Office Scripts are stored in a user's OneDrive.</span></span> <span data-ttu-id="38f85-123">Они следуют политикам хранения и удаления, указанным Microsoft OneDrive.</span><span class="sxs-lookup"><span data-stu-id="38f85-123">They follow the retention and deletion policies specified by Microsoft OneDrive.</span></span> <span data-ttu-id="38f85-124">Сведения о том, как обрабатывать сценарии, созданные и предоставленные пользователем, удаляемым из вашей организации, см. в статье [Хранение и удаление в OneDrive](/onedrive/retention-and-deletion).</span><span class="sxs-lookup"><span data-stu-id="38f85-124">To learn how to handle scripts that were created and shared by a user being removed from your organization, see [OneDrive retention and deletion](/onedrive/retention-and-deletion).</span></span>

## <a name="see-also"></a><span data-ttu-id="38f85-125">См. также</span><span class="sxs-lookup"><span data-stu-id="38f85-125">See also</span></span>

- [<span data-ttu-id="38f85-126">Общий доступ к сценариям Office в веб-программе Excel</span><span class="sxs-lookup"><span data-stu-id="38f85-126">Sharing Office Scripts in Excel for the Web</span></span>](https://support.microsoft.com/office/sharing-office-scripts-in-excel-for-the-web-226eddbc-3a44-4540-acfe-fccda3d1122b)
- [<span data-ttu-id="38f85-127">Устранение неполадок в сценариях Office</span><span class="sxs-lookup"><span data-stu-id="38f85-127">Troubleshooting Office Scripts</span></span>](../testing/troubleshooting.md)
- [<span data-ttu-id="38f85-128">Параметры сценариев Office в M365</span><span class="sxs-lookup"><span data-stu-id="38f85-128">Office Scripts settings in M365</span></span>](https://support.office.com/article/office-scripts-settings-in-m365-19d3c51a-6ca2-40ab-978d-60fa49554dcf)
- [<span data-ttu-id="38f85-129">Отмена эффектов сценария Office</span><span class="sxs-lookup"><span data-stu-id="38f85-129">Undo the effects of an Office Script</span></span>](../testing/undo.md)
