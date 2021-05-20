---
title: Ограничения и требования платформы с Office скриптами
description: Ограничения ресурсов и поддержка браузера для Office скриптов при использовании с Excel в Интернете
ms.date: 05/17/2021
localization_priority: Normal
ms.openlocfilehash: 7e81aaf2f96faeb67c815814fe3b7f1795651318
ms.sourcegitcommit: 4687693f02fc90a57ba30c461f35046e02e6f5fb
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 05/19/2021
ms.locfileid: "52545583"
---
# <a name="platform-limits-and-requirements-with-office-scripts"></a><span data-ttu-id="454d6-103">Ограничения и требования платформы с Office скриптами</span><span class="sxs-lookup"><span data-stu-id="454d6-103">Platform limits and requirements with Office Scripts</span></span>

<span data-ttu-id="454d6-104">Есть некоторые ограничения платформы, о которых вы должны знать при разработке Office скриптов.</span><span class="sxs-lookup"><span data-stu-id="454d6-104">There are some platform limitations of which you should be aware when developing Office Scripts.</span></span> <span data-ttu-id="454d6-105">В этой статье подробно подробно поддержки браузера и ограничения данных для Office скриптов для Excel в Интернете.</span><span class="sxs-lookup"><span data-stu-id="454d6-105">This article details the browser support and data limits for Office Scripts for Excel on the web.</span></span>

## <a name="browser-support"></a><span data-ttu-id="454d6-106">Поддержка браузеров</span><span class="sxs-lookup"><span data-stu-id="454d6-106">Browser support</span></span>

<span data-ttu-id="454d6-107">Office Скрипты работают в любом [браузере, который Office для Интернета.](https://support.microsoft.com/office/ad1303e0-a318-47aa-b409-d3a5eb44e452)</span><span class="sxs-lookup"><span data-stu-id="454d6-107">Office Scripts work in any browser that [supports Office for the web](https://support.microsoft.com/office/ad1303e0-a318-47aa-b409-d3a5eb44e452).</span></span> <span data-ttu-id="454d6-108">Однако некоторые функции JavaScript не поддерживаются в Internet Explorer 11 (IE 11).</span><span class="sxs-lookup"><span data-stu-id="454d6-108">However, some JavaScript features aren't supported in Internet Explorer 11 (IE 11).</span></span> <span data-ttu-id="454d6-109">Любые функции, [введенные в ES6 или позже,](https://www.w3schools.com/Js/js_es6.asp) не будут работать с IE 11.</span><span class="sxs-lookup"><span data-stu-id="454d6-109">Any features introduced in [ES6 or later](https://www.w3schools.com/Js/js_es6.asp) won't work with IE 11.</span></span> <span data-ttu-id="454d6-110">Если люди в вашей организации по-прежнему используют этот браузер, обязательно проверьте ваши скрипты в этой среде при их совместном использовании.</span><span class="sxs-lookup"><span data-stu-id="454d6-110">If people in your organization still use that browser, be sure to test your scripts in that environment when sharing them.</span></span>

[!INCLUDE [Teams support note](../includes/teams-support-note.md)]

### <a name="third-party-cookies"></a><span data-ttu-id="454d6-111">Сторонние файлы cookie</span><span class="sxs-lookup"><span data-stu-id="454d6-111">Third-party cookies</span></span>

<span data-ttu-id="454d6-112">Вашему браузеру нужны сторонние файлы cookie, которые могут показывать **вкладку Automate** в Excel в Интернете.</span><span class="sxs-lookup"><span data-stu-id="454d6-112">Your browser needs third-party cookies enabled to show the **Automate** tab in Excel on the web.</span></span> <span data-ttu-id="454d6-113">Проверьте настройки браузера, если вкладка не отображается.</span><span class="sxs-lookup"><span data-stu-id="454d6-113">Check your browser settings if the tab isn't being displayed.</span></span> <span data-ttu-id="454d6-114">Если вы используете сеанс частного браузера, возможно, потребуется каждый раз повторно включать эту настройку.</span><span class="sxs-lookup"><span data-stu-id="454d6-114">If you're using a private browser session, you may need to re-enable this setting each time.</span></span>

> [!NOTE]
> <span data-ttu-id="454d6-115">Некоторые браузеры называют эту настройку «всеми файлами cookie», а не «сторонними файлами cookie».</span><span class="sxs-lookup"><span data-stu-id="454d6-115">Some browsers refer to this setting as "all cookies", instead of "third-party cookies".</span></span>

#### <a name="instructions-for-adjusting-cookie-settings-in-popular-browsers"></a><span data-ttu-id="454d6-116">Инструкции по настройке настроек файлов cookie в популярных браузерах</span><span class="sxs-lookup"><span data-stu-id="454d6-116">Instructions for adjusting cookie settings in popular browsers</span></span>

- [<span data-ttu-id="454d6-117">Chrome</span><span class="sxs-lookup"><span data-stu-id="454d6-117">Chrome</span></span>](https://support.google.com/chrome/answer/95647)
- [<span data-ttu-id="454d6-118">Microsoft Edge</span><span class="sxs-lookup"><span data-stu-id="454d6-118">Edge</span></span>](https://support.microsoft.com/microsoft-edge/temporarily-allow-cookies-and-site-data-in-microsoft-edge-597f04f2-c0ce-f08c-7c2b-541086362bd2)
- [<span data-ttu-id="454d6-119">Firefox</span><span class="sxs-lookup"><span data-stu-id="454d6-119">Firefox</span></span>](https://support.mozilla.org/kb/disable-third-party-cookies)
- [<span data-ttu-id="454d6-120">Safari</span><span class="sxs-lookup"><span data-stu-id="454d6-120">Safari</span></span>](https://support.apple.com/guide/safari/manage-cookies-and-website-data-sfri11471/mac)

## <a name="data-limits"></a><span data-ttu-id="454d6-121">Ограничения данных</span><span class="sxs-lookup"><span data-stu-id="454d6-121">Data limits</span></span>

<span data-ttu-id="454d6-122">Существуют ограничения на то, Excel данные могут быть переданы сразу и сколько отдельных Power Automate транзакций могут быть проведены.</span><span class="sxs-lookup"><span data-stu-id="454d6-122">There are limits on how much Excel data can be transferred at once and how many individual Power Automate transactions can be conducted.</span></span>

### <a name="excel"></a><span data-ttu-id="454d6-123">Excel</span><span class="sxs-lookup"><span data-stu-id="454d6-123">Excel</span></span>

<span data-ttu-id="454d6-124">Excel для Интернета имеет следующие ограничения при звонках в трудовую книжку через скрипт:</span><span class="sxs-lookup"><span data-stu-id="454d6-124">Excel for the web has the following limitations when making calls to the workbook through a script:</span></span>

- <span data-ttu-id="454d6-125">Запросы и ответы ограничены **5MB**.</span><span class="sxs-lookup"><span data-stu-id="454d6-125">Requests and responses are limited to **5MB**.</span></span>
- <span data-ttu-id="454d6-126">Диапазон ограничен пятью **миллионами ячеек.**</span><span class="sxs-lookup"><span data-stu-id="454d6-126">A range is limited to **five million cells**.</span></span>

<span data-ttu-id="454d6-127">Если вы столкнулись с ошибками при работе с большими наборами данных, попробуйте использовать несколько меньших диапазонов вместо больших диапазонов.</span><span class="sxs-lookup"><span data-stu-id="454d6-127">If you're encountering errors when dealing with large datasets, try using multiple smaller ranges instead of larger ranges.</span></span> <span data-ttu-id="454d6-128">Например, [см.](../resources/samples/write-large-dataset.md)</span><span class="sxs-lookup"><span data-stu-id="454d6-128">For an example, see the [Write a large dataset](../resources/samples/write-large-dataset.md) sample.</span></span> <span data-ttu-id="454d6-129">Вы также можете использовать API, такие как [Range.getSpecialCells, для таргетинга](/javascript/api/office-scripts/excelscript/excelscript.range#getspecialcells-celltype--cellvaluetype-) на определенные ячейки вместо больших диапазонов.</span><span class="sxs-lookup"><span data-stu-id="454d6-129">You can also use APIs like [Range.getSpecialCells](/javascript/api/office-scripts/excelscript/excelscript.range#getspecialcells-celltype--cellvaluetype-) to target specific cells instead of large ranges.</span></span>

### <a name="power-automate"></a><span data-ttu-id="454d6-130">Power Automate</span><span class="sxs-lookup"><span data-stu-id="454d6-130">Power Automate</span></span>

<span data-ttu-id="454d6-131">При использовании Office скриптов Power Automate, каждый пользователь ограничен **400 вызовов на run Script действий в день**.</span><span class="sxs-lookup"><span data-stu-id="454d6-131">When using Office Scripts with Power Automate, each user is limited to **400 calls to the Run Script action per day**.</span></span> <span data-ttu-id="454d6-132">Этот лимит сбрасывается в 12:00 UTC.</span><span class="sxs-lookup"><span data-stu-id="454d6-132">This limit resets at 12:00 AM UTC.</span></span>

<span data-ttu-id="454d6-133">Платформа Power Automate также имеет ограничения использования, которые можно найти в следующих статьях:</span><span class="sxs-lookup"><span data-stu-id="454d6-133">The Power Automate platform also has usage limitations, which can be found in the following articles:</span></span>

- [<span data-ttu-id="454d6-134">Ограничения и конфигурация в Power Automate</span><span class="sxs-lookup"><span data-stu-id="454d6-134">Limits and configuration in Power Automate</span></span>](/power-automate/limits-and-config)
- [<span data-ttu-id="454d6-135">Известные проблемы и ограничения для Excel Online (Бизнес) разъем</span><span class="sxs-lookup"><span data-stu-id="454d6-135">Known issues and limitations for the Excel Online (Business) connector</span></span>](/connectors/excelonlinebusiness/#known-issues-and-limitations)

## <a name="see-also"></a><span data-ttu-id="454d6-136">См. также</span><span class="sxs-lookup"><span data-stu-id="454d6-136">See also</span></span>

- [<span data-ttu-id="454d6-137">Устранение неполадок Office скриптов</span><span class="sxs-lookup"><span data-stu-id="454d6-137">Troubleshoot Office Scripts</span></span>](troubleshooting.md)
- [<span data-ttu-id="454d6-138">Отмена эффектов сценариев Office</span><span class="sxs-lookup"><span data-stu-id="454d6-138">Undo the effects of Office Scripts</span></span>](undo.md)
- [<span data-ttu-id="454d6-139">Улучшение производительности ваших Office скриптов</span><span class="sxs-lookup"><span data-stu-id="454d6-139">Improve the performance of your Office Scripts</span></span>](../develop/web-client-performance.md)
- [<span data-ttu-id="454d6-140">Основы сценариев для Office сценариев в Excel в Интернете</span><span class="sxs-lookup"><span data-stu-id="454d6-140">Scripting Fundamentals for Office Scripts in Excel on the web</span></span>](../develop/scripting-fundamentals.md)
