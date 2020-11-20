---
title: Хранение и владение файлами сценариев Office
description: Сведения о том, как скрипты Office хранятся в Microsoft OneDrive и передаются между владельцами.
ms.date: 11/13/2020
localization_priority: Normal
ms.openlocfilehash: 648f3b2cf7e7d8d3bab2cf07a090e116e267a99a
ms.sourcegitcommit: 82d3c0ef1e187bcdeceb2b5fc3411186674fe150
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 11/18/2020
ms.locfileid: "49346872"
---
# <a name="office-scripts-file-storage-and-ownership"></a><span data-ttu-id="14b30-103">Хранение и владение файлами сценариев Office</span><span class="sxs-lookup"><span data-stu-id="14b30-103">Office Scripts file storage and ownership</span></span>

<span data-ttu-id="14b30-104">Сценарии Office хранятся в виде файлов **. остс** в Microsoft OneDrive.</span><span class="sxs-lookup"><span data-stu-id="14b30-104">Office Scripts are stored as **.osts** files in your Microsoft OneDrive.</span></span> <span data-ttu-id="14b30-105">Это позволит вашим скриптам существовать за прев определенной книге.</span><span class="sxs-lookup"><span data-stu-id="14b30-105">This allows your scripts to exist outside any particular workbook.</span></span> <span data-ttu-id="14b30-106">Параметры OneDrive управляют общим доступом и разрешениями для всех файлов script **. остс** ; не зависит от параметров Excel.</span><span class="sxs-lookup"><span data-stu-id="14b30-106">Your OneDrive settings control the shared access and permissions for all script **.osts** files; independent of any Excel settings.</span></span>

## <a name="file-storage"></a><span data-ttu-id="14b30-107">Хранение файлов</span><span class="sxs-lookup"><span data-stu-id="14b30-107">File storage</span></span>

<span data-ttu-id="14b30-108">Сценарии Office хранятся в вашем хранилище OneDrive.</span><span class="sxs-lookup"><span data-stu-id="14b30-108">You Office Scripts are stored in your OneDrive.</span></span> <span data-ttu-id="14b30-109">**Остс** -файлы находятся в папке **Scripts//документс/оффице** .</span><span class="sxs-lookup"><span data-stu-id="14b30-109">The **.osts** files are found in the **/Documents/Office Scripts/** folder.</span></span> <span data-ttu-id="14b30-110">Любые изменения, внесенные в эти файлы **. остс** , такие как переименование или удаление файлов, будут отражены в редакторе кода и в коллекции скриптов.</span><span class="sxs-lookup"><span data-stu-id="14b30-110">Any edits made to these **.osts** files, such as renaming or deleting files, will be reflected in the Code Editor and Script Gallery.</span></span>

<span data-ttu-id="14b30-111">Сценарии, совместно используемые с одной из книг, хранятся в хранилище OneDrive создателя скриптов.</span><span class="sxs-lookup"><span data-stu-id="14b30-111">Scripts that are shared with one of your workbooks remain in the script creator's OneDrive.</span></span> <span data-ttu-id="14b30-112">Они не копируются в локальные папки или папки OneDrive при запуске общего сценария в Excel.</span><span class="sxs-lookup"><span data-stu-id="14b30-112">They are not copied to any of your local or OneDrive folders when you run the shared script in Excel.</span></span> <span data-ttu-id="14b30-113">Кнопка **сделать копию** в редакторе кода сохраняет отдельную копию скрипта в OneDrive.</span><span class="sxs-lookup"><span data-stu-id="14b30-113">The **Make a Copy** button of the Code Editor saves a separate copy of the script in your OneDrive.</span></span> <span data-ttu-id="14b30-114">Изменения в копии не повлияют на исходный сценарий.</span><span class="sxs-lookup"><span data-stu-id="14b30-114">Changes to the copy don't affect the original script.</span></span>

### <a name="script-folders"></a><span data-ttu-id="14b30-115">Папки сценариев</span><span class="sxs-lookup"><span data-stu-id="14b30-115">Script folders</span></span>

<span data-ttu-id="14b30-116">Добавление папок в OneDrive поможет обеспечить упорядоченность сценариев.</span><span class="sxs-lookup"><span data-stu-id="14b30-116">Adding folders to your OneDrive helps keep your scripts organized.</span></span> <span data-ttu-id="14b30-117">Все папки в разделе **сценарии/документс/оффице/** отображаются в разделе **Мои сценарии** редактора кода.</span><span class="sxs-lookup"><span data-stu-id="14b30-117">Any folders under **/Documents/Office Scripts/** are displayed under the **My Scripts** section of the Code Editor.</span></span> <span data-ttu-id="14b30-118">Обратите внимание, что эти папки невозможно создать или удалить с помощью редактора кода.</span><span class="sxs-lookup"><span data-stu-id="14b30-118">Please note that these folders cannot be created or deleted by using the Code Editor.</span></span> <span data-ttu-id="14b30-119">Аналогичным образом, сценарии не могут быть помещены в папки или перемещаться по папкам с помощью редактора кода.</span><span class="sxs-lookup"><span data-stu-id="14b30-119">Likewise, scripts cannot be placed in folders, or moved across folders by using the Code Editor.</span></span>

![Некоторые сценарии, которые входят в папки, как показано в области задач редактор кода](../images/script-folders.png)

## <a name="file-ownership-and-retention"></a><span data-ttu-id="14b30-121">Владение файлами и хранение</span><span class="sxs-lookup"><span data-stu-id="14b30-121">File ownership and retention</span></span>

<span data-ttu-id="14b30-122">Скрипты Office хранятся в хранилище OneDrive пользователя.</span><span class="sxs-lookup"><span data-stu-id="14b30-122">Office Scripts are stored in a user's OneDrive.</span></span> <span data-ttu-id="14b30-123">Они следуют политикам хранения и удаления, указанным в Microsoft OneDrive.</span><span class="sxs-lookup"><span data-stu-id="14b30-123">They follow the retention and deletion policies specified by Microsoft OneDrive.</span></span> <span data-ttu-id="14b30-124">Сведения о том, как обрабатывать сценарии, созданные и предоставленные пользователем, удаляемым из вашей организации, см. в статье [Хранение и удаление в OneDrive](/onedrive/retention-and-deletion).</span><span class="sxs-lookup"><span data-stu-id="14b30-124">To learn how to handle scripts that were created and shared by a user being removed from your organization, see [OneDrive retention and deletion](/onedrive/retention-and-deletion).</span></span>

## <a name="see-also"></a><span data-ttu-id="14b30-125">См. также</span><span class="sxs-lookup"><span data-stu-id="14b30-125">See also</span></span>

- [<span data-ttu-id="14b30-126">Общий доступ к сценариям Office в веб-программе Excel</span><span class="sxs-lookup"><span data-stu-id="14b30-126">Sharing Office Scripts in Excel for the Web</span></span>](https://support.microsoft.com/office/sharing-office-scripts-in-excel-for-the-web-226eddbc-3a44-4540-acfe-fccda3d1122b)
- [<span data-ttu-id="14b30-127">Устранение неполадок в сценариях Office</span><span class="sxs-lookup"><span data-stu-id="14b30-127">Troubleshooting Office Scripts</span></span>](../testing/troubleshooting.md)
- [<span data-ttu-id="14b30-128">Параметры сценариев Office в M365</span><span class="sxs-lookup"><span data-stu-id="14b30-128">Office Scripts settings in M365</span></span>](https://support.office.com/article/office-scripts-settings-in-m365-19d3c51a-6ca2-40ab-978d-60fa49554dcf)
- [<span data-ttu-id="14b30-129">Отмена эффектов сценария Office</span><span class="sxs-lookup"><span data-stu-id="14b30-129">Undo the effects of an Office Script</span></span>](../testing/undo.md)
