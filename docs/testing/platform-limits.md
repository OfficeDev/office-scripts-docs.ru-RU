---
title: Ограничения и требования платформы с помощью скриптов Office
description: Ограничения ресурсов и поддержка браузера для скриптов Office при работе с Excel в Интернете
ms.date: 03/12/2021
localization_priority: Normal
ms.openlocfilehash: ef733562fb3caa8261fbbd8382923927a46cb7d4
ms.sourcegitcommit: 5ca286615a11d282e3f80023d22d36a039800eed
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 04/13/2021
ms.locfileid: "51689768"
---
# <a name="platform-limits-and-requirements-with-office-scripts"></a><span data-ttu-id="8cdbf-103">Ограничения и требования платформы с помощью скриптов Office</span><span class="sxs-lookup"><span data-stu-id="8cdbf-103">Platform limits and requirements with Office Scripts</span></span>

<span data-ttu-id="8cdbf-104">Существуют некоторые ограничения платформы, о которых следует помнить при разработке сценариев Office.</span><span class="sxs-lookup"><span data-stu-id="8cdbf-104">There are some platform limitations of which you should be aware when developing Office Scripts.</span></span> <span data-ttu-id="8cdbf-105">В этой статье подробно извесятся о поддержке браузера и ограничениях данных для Office Scripts for Excel в Интернете.</span><span class="sxs-lookup"><span data-stu-id="8cdbf-105">This article details the browser support and data limits for Office Scripts for Excel on the web.</span></span>

## <a name="browser-support"></a><span data-ttu-id="8cdbf-106">Поддержка браузеров</span><span class="sxs-lookup"><span data-stu-id="8cdbf-106">Browser support</span></span>

<span data-ttu-id="8cdbf-107">Скрипты Office работают в любом [браузере, который поддерживает Office для интернета.](https://support.microsoft.com/office/ad1303e0-a318-47aa-b409-d3a5eb44e452)</span><span class="sxs-lookup"><span data-stu-id="8cdbf-107">Office Scripts work in any browser that [supports Office for the web](https://support.microsoft.com/office/ad1303e0-a318-47aa-b409-d3a5eb44e452).</span></span> <span data-ttu-id="8cdbf-108">Однако некоторые функции JavaScript не поддерживаются в Internet Explorer 11 (IE 11).</span><span class="sxs-lookup"><span data-stu-id="8cdbf-108">However, some JavaScript features aren't supported in Internet Explorer 11 (IE 11).</span></span> <span data-ttu-id="8cdbf-109">Любые функции, [введенные в ES6 или](https://www.w3schools.com/Js/js_es6.asp) более поздней, не будут работать с IE 11.</span><span class="sxs-lookup"><span data-stu-id="8cdbf-109">Any features introduced in [ES6 or later](https://www.w3schools.com/Js/js_es6.asp) won't work with IE 11.</span></span> <span data-ttu-id="8cdbf-110">Если люди в организации по-прежнему используют этот браузер, обязательно проверьте свои скрипты в этой среде при их совместном использовании.</span><span class="sxs-lookup"><span data-stu-id="8cdbf-110">If people in your organization still use that browser, be sure to test your scripts in that environment when sharing them.</span></span>

[!INCLUDE [Teams support note](../includes/teams-support-note.md)]

### <a name="third-party-cookies"></a><span data-ttu-id="8cdbf-111">Сторонние файлы cookie</span><span class="sxs-lookup"><span data-stu-id="8cdbf-111">Third-party cookies</span></span>

<span data-ttu-id="8cdbf-112">Вашему браузеру необходимы сторонние файлы cookie, включенные для показа вкладки **Automate** в Excel в Интернете.</span><span class="sxs-lookup"><span data-stu-id="8cdbf-112">Your browser needs third-party cookies enabled to show the **Automate** tab in Excel on the web.</span></span> <span data-ttu-id="8cdbf-113">Проверьте параметры браузера, если вкладка не отображается.</span><span class="sxs-lookup"><span data-stu-id="8cdbf-113">Check your browser settings if the tab isn't being displayed.</span></span> <span data-ttu-id="8cdbf-114">При использовании закрытого сеанса браузера может потребоваться каждый раз повторно включить этот параметр.</span><span class="sxs-lookup"><span data-stu-id="8cdbf-114">If you're using a private browser session, you may need to re-enable this setting each time.</span></span>

> [!NOTE]
> <span data-ttu-id="8cdbf-115">Некоторые браузеры ссылаются на этот параметр как на "все файлы cookie", а не на "сторонние файлы cookie".</span><span class="sxs-lookup"><span data-stu-id="8cdbf-115">Some browsers refer to this setting as "all cookies", instead of "third-party cookies".</span></span>

#### <a name="instructions-for-adjusting-cookie-settings-in-popular-browsers"></a><span data-ttu-id="8cdbf-116">Инструкции по настройке параметров cookie в популярных браузерах</span><span class="sxs-lookup"><span data-stu-id="8cdbf-116">Instructions for adjusting cookie settings in popular browsers</span></span>

- [<span data-ttu-id="8cdbf-117">Chrome</span><span class="sxs-lookup"><span data-stu-id="8cdbf-117">Chrome</span></span>](https://support.google.com/chrome/answer/95647)
- [<span data-ttu-id="8cdbf-118">Microsoft Edge</span><span class="sxs-lookup"><span data-stu-id="8cdbf-118">Edge</span></span>](https://support.microsoft.com/microsoft-edge/temporarily-allow-cookies-and-site-data-in-microsoft-edge-597f04f2-c0ce-f08c-7c2b-541086362bd2)
- [<span data-ttu-id="8cdbf-119">Firefox</span><span class="sxs-lookup"><span data-stu-id="8cdbf-119">Firefox</span></span>](https://support.mozilla.org/kb/disable-third-party-cookies)
- [<span data-ttu-id="8cdbf-120">Safari</span><span class="sxs-lookup"><span data-stu-id="8cdbf-120">Safari</span></span>](https://support.apple.com/guide/safari/manage-cookies-and-website-data-sfri11471/mac)

## <a name="data-limits"></a><span data-ttu-id="8cdbf-121">Ограничения данных</span><span class="sxs-lookup"><span data-stu-id="8cdbf-121">Data limits</span></span>

<span data-ttu-id="8cdbf-122">Существуют ограничения на то, сколько данных Excel можно передавать одновременно и сколько отдельных транзакций power Automate можно проводить.</span><span class="sxs-lookup"><span data-stu-id="8cdbf-122">There are limits on how much Excel data can be transferred at once and how many individual Power Automate transactions can be conducted.</span></span>

### <a name="excel"></a><span data-ttu-id="8cdbf-123">Excel</span><span class="sxs-lookup"><span data-stu-id="8cdbf-123">Excel</span></span>

<span data-ttu-id="8cdbf-124">Excel для веб-сайта имеет следующие ограничения при вызове в книгу с помощью скрипта:</span><span class="sxs-lookup"><span data-stu-id="8cdbf-124">Excel for the web has the following limitations when making calls to the workbook through a script:</span></span>

- <span data-ttu-id="8cdbf-125">Запросы и ответы ограничены **5МБ.**</span><span class="sxs-lookup"><span data-stu-id="8cdbf-125">Requests and responses are limited to **5MB**.</span></span>
- <span data-ttu-id="8cdbf-126">Диапазон ограничен пятью **миллионами ячеек.**</span><span class="sxs-lookup"><span data-stu-id="8cdbf-126">A range is limited to **five million cells**.</span></span>

<span data-ttu-id="8cdbf-127">Если вы сталкиваетесь с ошибками при работе с большими наборами данных, попробуйте использовать несколько меньших диапазонов вместо больших диапазонов.</span><span class="sxs-lookup"><span data-stu-id="8cdbf-127">If you're encountering errors when dealing with large datasets, try using multiple smaller ranges instead of larger ranges.</span></span> <span data-ttu-id="8cdbf-128">Вы также можете использовать API, такие как [Range.getSpecialCells,](/javascript/api/office-scripts/excelscript/excelscript.range#getspecialcells-celltype--cellvaluetype-) для ориентации определенных ячеек вместо больших диапазонов.</span><span class="sxs-lookup"><span data-stu-id="8cdbf-128">You can also APIs like [Range.getSpecialCells](/javascript/api/office-scripts/excelscript/excelscript.range#getspecialcells-celltype--cellvaluetype-) to target specific cells instead of large ranges.</span></span>

### <a name="power-automate"></a><span data-ttu-id="8cdbf-129">Power Automate</span><span class="sxs-lookup"><span data-stu-id="8cdbf-129">Power Automate</span></span>

<span data-ttu-id="8cdbf-130">При использовании скриптов Office с помощью power Automate каждый пользователь может использовать только **400** вызовов к действию Сценарий запуска в день.</span><span class="sxs-lookup"><span data-stu-id="8cdbf-130">When using Office Scripts with Power Automate, each user is limited to **400 calls to the Run Script action per day**.</span></span> <span data-ttu-id="8cdbf-131">Это ограничение сбрасывается в 12:00 утра по UTC.</span><span class="sxs-lookup"><span data-stu-id="8cdbf-131">This limit resets at 12:00 AM UTC.</span></span>

<span data-ttu-id="8cdbf-132">Платформа Power Automate также имеет ограничения использования, которые можно найти в следующих статьях:</span><span class="sxs-lookup"><span data-stu-id="8cdbf-132">The Power Automate platform also has usage limitations, which can be found in the following articles:</span></span>

- [<span data-ttu-id="8cdbf-133">Ограничения и конфигурация в Power Automate</span><span class="sxs-lookup"><span data-stu-id="8cdbf-133">Limits and configuration in Power Automate</span></span>](/power-automate/limits-and-config)
- [<span data-ttu-id="8cdbf-134">Известные проблемы и ограничения для соединиттеля Excel Online (Бизнес)</span><span class="sxs-lookup"><span data-stu-id="8cdbf-134">Known issues and limitations for the Excel Online (Business) connector</span></span>](/connectors/excelonlinebusiness/#known-issues-and-limitations)

## <a name="see-also"></a><span data-ttu-id="8cdbf-135">См. также</span><span class="sxs-lookup"><span data-stu-id="8cdbf-135">See also</span></span>

- [<span data-ttu-id="8cdbf-136">Устранение неполадок в сценариях Office</span><span class="sxs-lookup"><span data-stu-id="8cdbf-136">Troubleshooting Office Scripts</span></span>](troubleshooting.md)
- [<span data-ttu-id="8cdbf-137">Отмена эффектов сценария Office</span><span class="sxs-lookup"><span data-stu-id="8cdbf-137">Undo the effects of an Office Script</span></span>](undo.md)
- [<span data-ttu-id="8cdbf-138">Повышение производительности скриптов Office</span><span class="sxs-lookup"><span data-stu-id="8cdbf-138">Improve the performance of your Office Scripts</span></span>](../develop/web-client-performance.md)
- [<span data-ttu-id="8cdbf-139">Основы скриптов для Office Scripts в Excel в Интернете</span><span class="sxs-lookup"><span data-stu-id="8cdbf-139">Scripting Fundamentals for Office Scripts in Excel on the web</span></span>](../develop/scripting-fundamentals.md)
