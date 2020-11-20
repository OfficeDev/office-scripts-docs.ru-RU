---
title: Устранение неполадок в сценариях Office
description: Советы и методы отладки сценариев Office, а также справочные ресурсы.
ms.date: 10/30/2020
localization_priority: Normal
ms.openlocfilehash: b45957bd336edce527397253cacec8cb09df715a
ms.sourcegitcommit: 82d3c0ef1e187bcdeceb2b5fc3411186674fe150
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 11/18/2020
ms.locfileid: "49342880"
---
# <a name="troubleshooting-office-scripts"></a><span data-ttu-id="e2ea2-103">Устранение неполадок в сценариях Office</span><span class="sxs-lookup"><span data-stu-id="e2ea2-103">Troubleshooting Office Scripts</span></span>

<span data-ttu-id="e2ea2-104">Когда вы разрабатываете сценарии Office, вы можете делать ошибки.</span><span class="sxs-lookup"><span data-stu-id="e2ea2-104">As you develop Office Scripts, you may make mistakes.</span></span> <span data-ttu-id="e2ea2-105">Всё в порядке.</span><span class="sxs-lookup"><span data-stu-id="e2ea2-105">It's okay.</span></span> <span data-ttu-id="e2ea2-106">У нас есть инструменты, которые помогают находить проблемы и работать с ними в идеальном состоянии.</span><span class="sxs-lookup"><span data-stu-id="e2ea2-106">We have tools that help find the problems and get your scripts working perfectly.</span></span>

## <a name="console-logs"></a><span data-ttu-id="e2ea2-107">Журналы консоли</span><span class="sxs-lookup"><span data-stu-id="e2ea2-107">Console logs</span></span>

<span data-ttu-id="e2ea2-108">Иногда при устранении неполадок необходимо напечатать сообщения на экране.</span><span class="sxs-lookup"><span data-stu-id="e2ea2-108">Sometimes while troubleshooting, you'll want to print messages to the screen.</span></span> <span data-ttu-id="e2ea2-109">Здесь можно отобразить текущее значение переменных или триггеров, которые вызываются.</span><span class="sxs-lookup"><span data-stu-id="e2ea2-109">These can show you the current value of variables or which code paths are being triggered.</span></span> <span data-ttu-id="e2ea2-110">Для этого зарегистрируете текст в консоли.</span><span class="sxs-lookup"><span data-stu-id="e2ea2-110">To do this, log text to the console.</span></span>

```TypeScript
console.log("Logging myRange's address.");
console.log(myRange.getAddress());
```

<span data-ttu-id="e2ea2-111">Строки, в которые передаются данные `console.log` , отображаются в консоли ведения журнала редактора кода.</span><span class="sxs-lookup"><span data-stu-id="e2ea2-111">Strings passed to `console.log` will be displayed in the Code Editor's logging console.</span></span> <span data-ttu-id="e2ea2-112">Чтобы включить консоль, нажмите кнопку **с многоточием** и выберите пункт **журналы...**</span><span class="sxs-lookup"><span data-stu-id="e2ea2-112">To turn on the console, press the **Ellipses** button and select **Logs...**</span></span>

<span data-ttu-id="e2ea2-113">Журналы не влияют на книгу.</span><span class="sxs-lookup"><span data-stu-id="e2ea2-113">Logs do not affect the workbook.</span></span>

## <a name="error-messages"></a><span data-ttu-id="e2ea2-114">Сообщения об ошибках</span><span class="sxs-lookup"><span data-stu-id="e2ea2-114">Error messages</span></span>

<span data-ttu-id="e2ea2-115">Если при выполнении сценария Excel возникла проблема, она выдает ошибку.</span><span class="sxs-lookup"><span data-stu-id="e2ea2-115">When your Excel Script encounters a problem running, it produces an error.</span></span> <span data-ttu-id="e2ea2-116">Появится всплывающее сообщение с вопросом, следует ли **просматривать журналы**.</span><span class="sxs-lookup"><span data-stu-id="e2ea2-116">You'll see a prompt pop-up asking if you want to **View Logs**.</span></span> <span data-ttu-id="e2ea2-117">Нажмите эту кнопку, чтобы открыть консоль и отобразить все ошибки.</span><span class="sxs-lookup"><span data-stu-id="e2ea2-117">Press that button to open the console and display any errors.</span></span>

## <a name="automate-tab-not-appearing-or-office-scripts-unavailable"></a><span data-ttu-id="e2ea2-118">Автоматизация вкладок не отображается или сценарии Office недоступны</span><span class="sxs-lookup"><span data-stu-id="e2ea2-118">Automate tab not appearing or Office Scripts unavailable</span></span>

<span data-ttu-id="e2ea2-119">Выполните следующие действия, чтобы устранить все неполадки, связанные с вкладкой **Автоматизация** , в Excel в Интернете.</span><span class="sxs-lookup"><span data-stu-id="e2ea2-119">The following steps should help troubleshoot any problems related to the **Automate** tab not appearing in Excel on the web.</span></span>

1. <span data-ttu-id="e2ea2-120">[Убедитесь, что лицензия microsoft 365 включает сценарии Office](../overview/excel.md#requirements).</span><span class="sxs-lookup"><span data-stu-id="e2ea2-120">[Make sure your Microsoft 365 license includes Office Scripts](../overview/excel.md#requirements).</span></span>
1. <span data-ttu-id="e2ea2-121">[Убедитесь, что ваш браузер поддерживается](platform-limits.md#browser-support).</span><span class="sxs-lookup"><span data-stu-id="e2ea2-121">[Check that your browser is supported](platform-limits.md#browser-support).</span></span>
1. <span data-ttu-id="e2ea2-122">[Убедитесь, что сторонние файлы Cookie включены](platform-limits.md#third-party-cookies).</span><span class="sxs-lookup"><span data-stu-id="e2ea2-122">[Ensure third-party cookies are enabled](platform-limits.md#third-party-cookies).</span></span>
1. <span data-ttu-id="e2ea2-123">[Убедитесь, что администратор не отключил сценарии Office в центре администрирования Microsoft 365](/microsoft-365/admin/manage/manage-office-scripts-settings).</span><span class="sxs-lookup"><span data-stu-id="e2ea2-123">[Ensure that your admin has not disabled Office Scripts in the Microsoft 365 admin center](/microsoft-365/admin/manage/manage-office-scripts-settings).</span></span>

[!INCLUDE [Teams support note](../includes/teams-support-note.md)]

## <a name="help-resources"></a><span data-ttu-id="e2ea2-124">Справочные материалы</span><span class="sxs-lookup"><span data-stu-id="e2ea2-124">Help resources</span></span>

<span data-ttu-id="e2ea2-125">[Переполнение стека](https://stackoverflow.com/questions/tagged/office-scripts) — это сообщество разработчиков, которые могут помочь при возникновении проблем с написанием кода.</span><span class="sxs-lookup"><span data-stu-id="e2ea2-125">[Stack Overflow](https://stackoverflow.com/questions/tagged/office-scripts) is a community of developers willing to help with coding problems.</span></span> <span data-ttu-id="e2ea2-126">Часто вы сможете найти решение проблемы с помощью быстрого поиска переполнения стека.</span><span class="sxs-lookup"><span data-stu-id="e2ea2-126">Often, you'll be able to find the solution to your problem through a quick Stack Overflow search.</span></span> <span data-ttu-id="e2ea2-127">Если это не так, задайте свой вопрос и пометьте его тегом "Office – Scripts".</span><span class="sxs-lookup"><span data-stu-id="e2ea2-127">If not, ask your question and tag it with the "office-scripts" tag.</span></span> <span data-ttu-id="e2ea2-128">Запомните, что вы создаете *сценарий* Office, а не *надстройку* Office.</span><span class="sxs-lookup"><span data-stu-id="e2ea2-128">Be sure to mention you're creating an Office *Script*, not an Office *Add-in*.</span></span>

<span data-ttu-id="e2ea2-129">Если возникла проблема с API JavaScript для Office, создайте проблему в репозитории GitHub [OfficeDev/Office-JS](https://github.com/OfficeDev/office-js) .</span><span class="sxs-lookup"><span data-stu-id="e2ea2-129">If you encounter a problem with the Office JavaScript API, create an issue in the [OfficeDev/office-js](https://github.com/OfficeDev/office-js) GitHub repository.</span></span> <span data-ttu-id="e2ea2-130">Участники группы продукта будут отвечать на проблемы и предоставлять дальнейшую помощь.</span><span class="sxs-lookup"><span data-stu-id="e2ea2-130">Members of the product team will respond to issues and provide further assistance.</span></span> <span data-ttu-id="e2ea2-131">Создание ошибки в репозитории **OfficeDev/Office-JS** свидетельствует о том, что в библиотеке API JavaScript для Office обнаружен изъян, который должен быть адрес группой разработки продуктов.</span><span class="sxs-lookup"><span data-stu-id="e2ea2-131">Creating an issue in the **OfficeDev/office-js** repository indicates you have found a flaw in the Office JavaScript API library that the product team should address.</span></span>

<span data-ttu-id="e2ea2-132">При возникновении проблем с регистратором действий или редактором отправьте отзыв с помощью кнопки " **справка > отзыв** " в Excel.</span><span class="sxs-lookup"><span data-stu-id="e2ea2-132">If there is a problem with the Action Recorder or Editor, send feedback through the **Help > Feedback** button in Excel.</span></span>

## <a name="see-also"></a><span data-ttu-id="e2ea2-133">См. также</span><span class="sxs-lookup"><span data-stu-id="e2ea2-133">See also</span></span>

- [<span data-ttu-id="e2ea2-134">Сценарии Office в Excel в Интернете</span><span class="sxs-lookup"><span data-stu-id="e2ea2-134">Office Scripts in Excel on the web</span></span>](../overview/excel.md)
- [<span data-ttu-id="e2ea2-135">Основные сведения о сценариях для сценариев Office в Excel в Интернете</span><span class="sxs-lookup"><span data-stu-id="e2ea2-135">Scripting Fundamentals for Office Scripts in Excel on the web</span></span>](../develop/scripting-fundamentals.md)
- [<span data-ttu-id="e2ea2-136">Пределы платформы с помощью сценариев Office</span><span class="sxs-lookup"><span data-stu-id="e2ea2-136">Platform Limits with Office Scripts</span></span>](platform-limits.md)
- [<span data-ttu-id="e2ea2-137">Повышение производительности сценариев Office</span><span class="sxs-lookup"><span data-stu-id="e2ea2-137">Improve the performance of your Office Scripts</span></span>](../develop/web-client-performance.md)
- [<span data-ttu-id="e2ea2-138">Отмена эффектов сценария Office</span><span class="sxs-lookup"><span data-stu-id="e2ea2-138">Undo the effects of an Office Script</span></span>](undo.md)
