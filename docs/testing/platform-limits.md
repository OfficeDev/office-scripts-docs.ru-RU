---
title: Ограничения и требования платформы с Office скриптами
description: Ограничения ресурсов и поддержка браузера для Office скриптов при Excel в Интернете
ms.date: 05/17/2021
localization_priority: Normal
ms.openlocfilehash: 7e81aaf2f96faeb67c815814fe3b7f1795651318
ms.sourcegitcommit: 4687693f02fc90a57ba30c461f35046e02e6f5fb
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 05/19/2021
ms.locfileid: "52545583"
---
# <a name="platform-limits-and-requirements-with-office-scripts"></a><span data-ttu-id="d02bf-103">Ограничения и требования платформы с Office скриптами</span><span class="sxs-lookup"><span data-stu-id="d02bf-103">Platform limits and requirements with Office Scripts</span></span>

<span data-ttu-id="d02bf-104">Существуют некоторые ограничения платформы, о которых следует знать при разработке Office скриптов.</span><span class="sxs-lookup"><span data-stu-id="d02bf-104">There are some platform limitations of which you should be aware when developing Office Scripts.</span></span> <span data-ttu-id="d02bf-105">В этой статье подробно извесятся о поддержке браузера и ограничениях данных для Office скриптов для Excel в Интернете.</span><span class="sxs-lookup"><span data-stu-id="d02bf-105">This article details the browser support and data limits for Office Scripts for Excel on the web.</span></span>

## <a name="browser-support"></a><span data-ttu-id="d02bf-106">Поддержка браузеров</span><span class="sxs-lookup"><span data-stu-id="d02bf-106">Browser support</span></span>

<span data-ttu-id="d02bf-107">Office Скрипты работают в любом [браузере, Office для веб-сайта.](https://support.microsoft.com/office/ad1303e0-a318-47aa-b409-d3a5eb44e452)</span><span class="sxs-lookup"><span data-stu-id="d02bf-107">Office Scripts work in any browser that [supports Office for the web](https://support.microsoft.com/office/ad1303e0-a318-47aa-b409-d3a5eb44e452).</span></span> <span data-ttu-id="d02bf-108">Однако некоторые функции JavaScript не поддерживаются в Internet Explorer 11 (IE 11).</span><span class="sxs-lookup"><span data-stu-id="d02bf-108">However, some JavaScript features aren't supported in Internet Explorer 11 (IE 11).</span></span> <span data-ttu-id="d02bf-109">Любые функции, [введенные в ES6 или](https://www.w3schools.com/Js/js_es6.asp) более поздней, не будут работать с IE 11.</span><span class="sxs-lookup"><span data-stu-id="d02bf-109">Any features introduced in [ES6 or later](https://www.w3schools.com/Js/js_es6.asp) won't work with IE 11.</span></span> <span data-ttu-id="d02bf-110">Если люди в организации по-прежнему используют этот браузер, обязательно проверьте свои скрипты в этой среде при их совместном использовании.</span><span class="sxs-lookup"><span data-stu-id="d02bf-110">If people in your organization still use that browser, be sure to test your scripts in that environment when sharing them.</span></span>

[!INCLUDE [Teams support note](../includes/teams-support-note.md)]

### <a name="third-party-cookies"></a><span data-ttu-id="d02bf-111">Сторонние файлы cookie</span><span class="sxs-lookup"><span data-stu-id="d02bf-111">Third-party cookies</span></span>

<span data-ttu-id="d02bf-112">Вашему браузеру нужны сторонние файлы cookie, включенные для показа вкладки **Automate** в Excel в Интернете.</span><span class="sxs-lookup"><span data-stu-id="d02bf-112">Your browser needs third-party cookies enabled to show the **Automate** tab in Excel on the web.</span></span> <span data-ttu-id="d02bf-113">Проверьте параметры браузера, если вкладка не отображается.</span><span class="sxs-lookup"><span data-stu-id="d02bf-113">Check your browser settings if the tab isn't being displayed.</span></span> <span data-ttu-id="d02bf-114">При использовании закрытого сеанса браузера может потребоваться каждый раз повторно включить этот параметр.</span><span class="sxs-lookup"><span data-stu-id="d02bf-114">If you're using a private browser session, you may need to re-enable this setting each time.</span></span>

> [!NOTE]
> <span data-ttu-id="d02bf-115">Некоторые браузеры ссылаются на этот параметр как на "все файлы cookie", а не на "сторонние файлы cookie".</span><span class="sxs-lookup"><span data-stu-id="d02bf-115">Some browsers refer to this setting as "all cookies", instead of "third-party cookies".</span></span>

#### <a name="instructions-for-adjusting-cookie-settings-in-popular-browsers"></a><span data-ttu-id="d02bf-116">Инструкции по настройке параметров cookie в популярных браузерах</span><span class="sxs-lookup"><span data-stu-id="d02bf-116">Instructions for adjusting cookie settings in popular browsers</span></span>

- [<span data-ttu-id="d02bf-117">Chrome</span><span class="sxs-lookup"><span data-stu-id="d02bf-117">Chrome</span></span>](https://support.google.com/chrome/answer/95647)
- [<span data-ttu-id="d02bf-118">Microsoft Edge</span><span class="sxs-lookup"><span data-stu-id="d02bf-118">Edge</span></span>](https://support.microsoft.com/microsoft-edge/temporarily-allow-cookies-and-site-data-in-microsoft-edge-597f04f2-c0ce-f08c-7c2b-541086362bd2)
- [<span data-ttu-id="d02bf-119">Firefox</span><span class="sxs-lookup"><span data-stu-id="d02bf-119">Firefox</span></span>](https://support.mozilla.org/kb/disable-third-party-cookies)
- [<span data-ttu-id="d02bf-120">Safari</span><span class="sxs-lookup"><span data-stu-id="d02bf-120">Safari</span></span>](https://support.apple.com/guide/safari/manage-cookies-and-website-data-sfri11471/mac)

## <a name="data-limits"></a><span data-ttu-id="d02bf-121">Ограничения данных</span><span class="sxs-lookup"><span data-stu-id="d02bf-121">Data limits</span></span>

<span data-ttu-id="d02bf-122">Существуют ограничения по объему Excel данных, которые могут быть переданы одновременно, и Power Automate отдельных транзакций.</span><span class="sxs-lookup"><span data-stu-id="d02bf-122">There are limits on how much Excel data can be transferred at once and how many individual Power Automate transactions can be conducted.</span></span>

### <a name="excel"></a><span data-ttu-id="d02bf-123">Excel</span><span class="sxs-lookup"><span data-stu-id="d02bf-123">Excel</span></span>

<span data-ttu-id="d02bf-124">Excel веб-страницы имеет следующие ограничения при вызове в книгу с помощью скрипта:</span><span class="sxs-lookup"><span data-stu-id="d02bf-124">Excel for the web has the following limitations when making calls to the workbook through a script:</span></span>

- <span data-ttu-id="d02bf-125">Запросы и ответы ограничены **5МБ.**</span><span class="sxs-lookup"><span data-stu-id="d02bf-125">Requests and responses are limited to **5MB**.</span></span>
- <span data-ttu-id="d02bf-126">Диапазон ограничен пятью **миллионами ячеек.**</span><span class="sxs-lookup"><span data-stu-id="d02bf-126">A range is limited to **five million cells**.</span></span>

<span data-ttu-id="d02bf-127">Если вы сталкиваетесь с ошибками при работе с большими наборами данных, попробуйте использовать несколько меньших диапазонов вместо больших диапазонов.</span><span class="sxs-lookup"><span data-stu-id="d02bf-127">If you're encountering errors when dealing with large datasets, try using multiple smaller ranges instead of larger ranges.</span></span> <span data-ttu-id="d02bf-128">Пример см. в [примере Write a large dataset](../resources/samples/write-large-dataset.md) sample.</span><span class="sxs-lookup"><span data-stu-id="d02bf-128">For an example, see the [Write a large dataset](../resources/samples/write-large-dataset.md) sample.</span></span> <span data-ttu-id="d02bf-129">Вы также можете использовать API, такие как [Range.getSpecialCells,](/javascript/api/office-scripts/excelscript/excelscript.range#getspecialcells-celltype--cellvaluetype-) чтобы нацелить определенные ячейки вместо больших диапазонов.</span><span class="sxs-lookup"><span data-stu-id="d02bf-129">You can also use APIs like [Range.getSpecialCells](/javascript/api/office-scripts/excelscript/excelscript.range#getspecialcells-celltype--cellvaluetype-) to target specific cells instead of large ranges.</span></span>

### <a name="power-automate"></a><span data-ttu-id="d02bf-130">Power Automate</span><span class="sxs-lookup"><span data-stu-id="d02bf-130">Power Automate</span></span>

<span data-ttu-id="d02bf-131">При использовании Office скриптов с Power Automate каждый пользователь может использовать **400** вызовов к действию Run Script в день.</span><span class="sxs-lookup"><span data-stu-id="d02bf-131">When using Office Scripts with Power Automate, each user is limited to **400 calls to the Run Script action per day**.</span></span> <span data-ttu-id="d02bf-132">Это ограничение сбрасывается в 12:00 утра по UTC.</span><span class="sxs-lookup"><span data-stu-id="d02bf-132">This limit resets at 12:00 AM UTC.</span></span>

<span data-ttu-id="d02bf-133">Платформа Power Automate также имеет ограничения использования, которые можно найти в следующих статьях:</span><span class="sxs-lookup"><span data-stu-id="d02bf-133">The Power Automate platform also has usage limitations, which can be found in the following articles:</span></span>

- [<span data-ttu-id="d02bf-134">Ограничения и конфигурация в Power Automate</span><span class="sxs-lookup"><span data-stu-id="d02bf-134">Limits and configuration in Power Automate</span></span>](/power-automate/limits-and-config)
- [<span data-ttu-id="d02bf-135">Известные проблемы и ограничения для соединиттеля Excel Online (Бизнес)</span><span class="sxs-lookup"><span data-stu-id="d02bf-135">Known issues and limitations for the Excel Online (Business) connector</span></span>](/connectors/excelonlinebusiness/#known-issues-and-limitations)

## <a name="see-also"></a><span data-ttu-id="d02bf-136">См. также</span><span class="sxs-lookup"><span data-stu-id="d02bf-136">See also</span></span>

- [<span data-ttu-id="d02bf-137">Устранение Office скриптов</span><span class="sxs-lookup"><span data-stu-id="d02bf-137">Troubleshoot Office Scripts</span></span>](troubleshooting.md)
- [<span data-ttu-id="d02bf-138">Отмена эффектов сценариев Office</span><span class="sxs-lookup"><span data-stu-id="d02bf-138">Undo the effects of Office Scripts</span></span>](undo.md)
- [<span data-ttu-id="d02bf-139">Повышение производительности Office скриптов</span><span class="sxs-lookup"><span data-stu-id="d02bf-139">Improve the performance of your Office Scripts</span></span>](../develop/web-client-performance.md)
- [<span data-ttu-id="d02bf-140">Основы сценариев для Office скриптов в Excel в Интернете</span><span class="sxs-lookup"><span data-stu-id="d02bf-140">Scripting Fundamentals for Office Scripts in Excel on the web</span></span>](../develop/scripting-fundamentals.md)
