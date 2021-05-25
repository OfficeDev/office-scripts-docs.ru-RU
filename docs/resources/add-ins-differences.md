---
title: Различия между сценариями Office и надстройками Office
description: Различия между поведением и API между Office скриптами и Office надстройки.
ms.date: 06/01/2020
localization_priority: Normal
ms.openlocfilehash: 5c30406867da05952dedda684f765df5e7a7e53f
ms.sourcegitcommit: 09d8859d5269ada8f1d0e141f6b5a4f96d95a739
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 05/24/2021
ms.locfileid: "52631680"
---
# <a name="differences-between-office-scripts-and-office-add-ins"></a><span data-ttu-id="fdb3a-103">Различия между сценариями Office и надстройками Office</span><span class="sxs-lookup"><span data-stu-id="fdb3a-103">Differences between Office Scripts and Office Add-ins</span></span>

<span data-ttu-id="fdb3a-104">Office Надстройки и Office скрипты имеют много общего.</span><span class="sxs-lookup"><span data-stu-id="fdb3a-104">Office Add-ins and Office Scripts have a lot in common.</span></span> <span data-ttu-id="fdb3a-105">Они оба предлагают автоматизированное управление Excel aPI JavaScript.</span><span class="sxs-lookup"><span data-stu-id="fdb3a-105">They both offer automated control of an Excel workbook a JavaScript API.</span></span> <span data-ttu-id="fdb3a-106">Однако API Office скриптов — это специализированная синхронная версия API Office JavaScript.</span><span class="sxs-lookup"><span data-stu-id="fdb3a-106">However, the Office Scripts APIs are a specialized, synchronous version of the Office JavaScript API.</span></span>

:::image type="content" source="../images/office-programmability-diagram.png" alt-text="Схема с четырьмя квадрантами, показывающая области фокусиза для различных решений Office разностоверных решений. Сценарии Office и Office веб-надстройки ориентированы на веб-сайты и совместную работу, но Office скрипты обслуживают конечных пользователей (в то время как Office веб-надстройки ориентированы на профессиональных разработчиков)":::

<span data-ttu-id="fdb3a-108">Office Сценарии запускаются до завершения с помощью ручного нажатия кнопки или в [Power Automate,](https://flow.microsoft.com/)в то время как Office надстройки сохраняются, пока их области задач открыты.</span><span class="sxs-lookup"><span data-stu-id="fdb3a-108">Office Scripts run to completion with a manual button press or as a step in [Power Automate](https://flow.microsoft.com/), whereas Office Add-ins persist while their task panes are open.</span></span> <span data-ttu-id="fdb3a-109">Это означает, что надстройки могут поддерживать состояние во время сеанса, в то время как Office скрипты не поддерживают внутреннее состояние между запусками.</span><span class="sxs-lookup"><span data-stu-id="fdb3a-109">This means the add-ins can maintain state during a session, whereas Office Scripts do not maintain an internal state between runs.</span></span> <span data-ttu-id="fdb3a-110">Если вы Excel, что расширение должно превышать возможности платформы скриптов, [](/office/dev/add-ins) посетите документацию Office надстройок, чтобы узнать больше о Office надстройки.</span><span class="sxs-lookup"><span data-stu-id="fdb3a-110">If you find that your Excel extension needs to exceed the scripting platform's capabilities, visit the [Office Add-ins documentation](/office/dev/add-ins) to learn more about Office Add-ins.</span></span>

<span data-ttu-id="fdb3a-111">В остальной части этой статьи описываются основные различия между Office надстройки и Office скриптами.</span><span class="sxs-lookup"><span data-stu-id="fdb3a-111">The rest of this article describes on the main differences between Office Add-ins and Office Scripts.</span></span>

## <a name="platform-support"></a><span data-ttu-id="fdb3a-112">Поддержка платформы</span><span class="sxs-lookup"><span data-stu-id="fdb3a-112">Platform Support</span></span>

<span data-ttu-id="fdb3a-113">Office Надстройки — это кроссплатформы.</span><span class="sxs-lookup"><span data-stu-id="fdb3a-113">Office Add-ins are cross-platform.</span></span> <span data-ttu-id="fdb3a-114">Они работают на Windows, Mac, iOS и веб-платформах и предоставляют одинаковый опыт для каждой из них.</span><span class="sxs-lookup"><span data-stu-id="fdb3a-114">They work across Windows desktop, Mac, iOS, and web platforms and provide the same experience on each.</span></span> <span data-ttu-id="fdb3a-115">Любое исключение из этого отмечено в документации отдельного API.</span><span class="sxs-lookup"><span data-stu-id="fdb3a-115">Any exception to this is noted in the documentation of the individual API.</span></span>

<span data-ttu-id="fdb3a-116">Office В настоящее время скрипты поддерживаются только для Excel в Интернете.</span><span class="sxs-lookup"><span data-stu-id="fdb3a-116">Office Scripts are currently only supported by for Excel on the web.</span></span> <span data-ttu-id="fdb3a-117">Вся запись, редактирование и запуск делаются на веб-платформе.</span><span class="sxs-lookup"><span data-stu-id="fdb3a-117">All recording, editing, and running is done on the web platform.</span></span>

## <a name="apis"></a><span data-ttu-id="fdb3a-118">Интерфейсы API</span><span class="sxs-lookup"><span data-stu-id="fdb3a-118">APIs</span></span>

<span data-ttu-id="fdb3a-119">Хотя Office API JavaScript для Office надстройки и API Office скриптов имеют некоторые функциональные возможности, они являются различными платформами.</span><span class="sxs-lookup"><span data-stu-id="fdb3a-119">While the Office JavaScript APIs for Office Add-ins and the Office Scripts APIs share some functionality, they are different platforms.</span></span> <span data-ttu-id="fdb3a-120">API Office скриптов — оптимизированная синхронная версия Excel API JavaScript.</span><span class="sxs-lookup"><span data-stu-id="fdb3a-120">The Office Scripts APIs are an optimized, synchronous version of the Excel JavaScript API model.</span></span> <span data-ttu-id="fdb3a-121">Главное отличие заключается в использовании `load` / `sync` парадигмы с надстройки. Кроме того, надстройки предлагают API для событий и более широкий набор функций за пределами Excel, известных как общие API.</span><span class="sxs-lookup"><span data-stu-id="fdb3a-121">The major difference is usage of the `load`/`sync` paradigm with add-ins. Additionally, add-ins offer APIs for events and a broader set of functionality outside of Excel, known as the Common APIs.</span></span>

### <a name="events"></a><span data-ttu-id="fdb3a-122">События</span><span class="sxs-lookup"><span data-stu-id="fdb3a-122">Events</span></span>

<span data-ttu-id="fdb3a-123">Office Сценарии не поддерживают [события.](/office/dev/add-ins/excel/excel-add-ins-events)</span><span class="sxs-lookup"><span data-stu-id="fdb3a-123">Office Scripts do not support [events](/office/dev/add-ins/excel/excel-add-ins-events).</span></span> <span data-ttu-id="fdb3a-124">Каждый скрипт запускает код одним `main` методом, а затем заканчивается.</span><span class="sxs-lookup"><span data-stu-id="fdb3a-124">Every script runs the code in a single `main` method, then ends.</span></span> <span data-ttu-id="fdb3a-125">Он не активируется при запуске событий и, следовательно, не может зарегистрировать события.</span><span class="sxs-lookup"><span data-stu-id="fdb3a-125">It does not reactivate when events are triggered, and thus, cannot register events.</span></span>

### <a name="common-apis"></a><span data-ttu-id="fdb3a-126">Общие API</span><span class="sxs-lookup"><span data-stu-id="fdb3a-126">Common APIs</span></span>

<span data-ttu-id="fdb3a-127">Office Скрипты не могут использовать [общие API.](/javascript/api/office)</span><span class="sxs-lookup"><span data-stu-id="fdb3a-127">Office Scripts cannot use [Common APIs](/javascript/api/office).</span></span> <span data-ttu-id="fdb3a-128">Если требуется проверка подлинности, диалоговое окно или другие функции, поддерживаемые только общими API, скорее всего, потребуется создать надстройки Office, а не Office скрипта.</span><span class="sxs-lookup"><span data-stu-id="fdb3a-128">If you need authentication, dialog windows, or other features that are only supported by Common APIs, you'll likely need to create an Office Add-in instead of an Office Script.</span></span>

## <a name="see-also"></a><span data-ttu-id="fdb3a-129">См. также</span><span class="sxs-lookup"><span data-stu-id="fdb3a-129">See also</span></span>

- [<span data-ttu-id="fdb3a-130">Сценарии Office в Excel в Интернете</span><span class="sxs-lookup"><span data-stu-id="fdb3a-130">Office Scripts in Excel on the web</span></span>](../overview/excel.md)
- [<span data-ttu-id="fdb3a-131">Различия между Office скриптами и макросами VBA</span><span class="sxs-lookup"><span data-stu-id="fdb3a-131">Differences between Office Scripts and VBA macros</span></span>](vba-differences.md)
- [<span data-ttu-id="fdb3a-132">Устранение неполадок в сценариях Office</span><span class="sxs-lookup"><span data-stu-id="fdb3a-132">Troubleshooting Office Scripts</span></span>](../testing/troubleshooting.md)
- [<span data-ttu-id="fdb3a-133">Создание надстройки области задач Excel</span><span class="sxs-lookup"><span data-stu-id="fdb3a-133">Build an Excel task pane add-in</span></span>](/office/dev/add-ins/quickstarts/excel-quickstart-jquery)
