---
title: Office 脚本与 Office 加载项之间的差异
description: Office 脚本与 Office 外接程序之间的行为和 API 差异。
ms.date: 06/01/2020
localization_priority: Normal
ms.openlocfilehash: fc2029780190672c633e00e26f44273e4311c754
ms.sourcegitcommit: aec3c971c6640429f89b6bb99d2c95ea06725599
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 06/25/2020
ms.locfileid: "44878659"
---
# <a name="differences-between-office-scripts-and-office-add-ins"></a><span data-ttu-id="c186a-103">Office 脚本与 Office 加载项之间的差异</span><span class="sxs-lookup"><span data-stu-id="c186a-103">Differences between Office Scripts and Office Add-ins</span></span>

<span data-ttu-id="c186a-104">Office 外接程序和 Office 脚本具有很多共同之处。</span><span class="sxs-lookup"><span data-stu-id="c186a-104">Office Add-ins and Office Scripts have a lot in common.</span></span> <span data-ttu-id="c186a-105">它们都提供对 Excel 工作簿的自动控制（JavaScript API）。</span><span class="sxs-lookup"><span data-stu-id="c186a-105">They both offer automated control of an Excel workbook a JavaScript API.</span></span> <span data-ttu-id="c186a-106">但是，Office 脚本 Api 是 Office JavaScript API 的专用的同步版本。</span><span class="sxs-lookup"><span data-stu-id="c186a-106">However, the Office Scripts APIs are a specialized, synchronous version of the Office JavaScript API.</span></span>

![显示不同 Office 扩展性解决方案的焦点区域的四象限图。](../images/office-programmability-diagram.png)

<span data-ttu-id="c186a-109">Office 脚本运行到完成时需要按手动按钮按下或以 "[自动](https://flow.microsoft.com/)运行" 的步骤，而 office 加载项在其任务窗格处于打开状态时保持不变。</span><span class="sxs-lookup"><span data-stu-id="c186a-109">Office Scripts run to completion with a manual button press or as a step in [Power Automate](https://flow.microsoft.com/), whereas Office Add-ins persist while their task panes are open.</span></span> <span data-ttu-id="c186a-110">这意味着外接程序可以在会话期间维护状态，而 Office 脚本不会在两个运行之间保持内部状态。</span><span class="sxs-lookup"><span data-stu-id="c186a-110">This means the add-ins can maintain state during a session, whereas Office Scripts do not maintain an internal state between runs.</span></span> <span data-ttu-id="c186a-111">如果发现您的 Excel 扩展需要超过脚本平台的功能，请访问[Office 外接程序文档](/office/dev/add-ins)以了解有关 Office 外接程序的详细信息。</span><span class="sxs-lookup"><span data-stu-id="c186a-111">If you find that your Excel extension needs to exceed the scripting platform's capabilities, visit the [Office Add-ins documentation](/office/dev/add-ins) to learn more about Office Add-ins.</span></span>

<span data-ttu-id="c186a-112">本文的其余部分将介绍 Office 外接程序和 Office 脚本之间的主要差异。</span><span class="sxs-lookup"><span data-stu-id="c186a-112">The rest of this article describes on the main differences between Office Add-ins and Office Scripts.</span></span>

## <a name="platform-support"></a><span data-ttu-id="c186a-113">平台支持</span><span class="sxs-lookup"><span data-stu-id="c186a-113">Platform Support</span></span>

<span data-ttu-id="c186a-114">Office 外接程序是跨平台的。</span><span class="sxs-lookup"><span data-stu-id="c186a-114">Office Add-ins are cross-platform.</span></span> <span data-ttu-id="c186a-115">它们在 Windows 桌面、Mac、iOS 和 web 平台上工作，并在每个平台上提供相同的体验。</span><span class="sxs-lookup"><span data-stu-id="c186a-115">They work across Windows desktop, Mac, iOS, and web platforms and provide the same experience on each.</span></span> <span data-ttu-id="c186a-116">每个 API 的文档中注明了此错误的任何例外。</span><span class="sxs-lookup"><span data-stu-id="c186a-116">Any exception to this is noted in the documentation of the individual API.</span></span>

<span data-ttu-id="c186a-117">Office 脚本目前仅对 web 上的 Excel 受支持。</span><span class="sxs-lookup"><span data-stu-id="c186a-117">Office Scripts are currently only supported by for Excel on the web.</span></span> <span data-ttu-id="c186a-118">所有录制、编辑和运行都是在 web 平台上完成的。</span><span class="sxs-lookup"><span data-stu-id="c186a-118">All recording, editing, and running is done on the web platform.</span></span>

## <a name="apis"></a><span data-ttu-id="c186a-119">API</span><span class="sxs-lookup"><span data-stu-id="c186a-119">APIs</span></span>

<span data-ttu-id="c186a-120">Office 外接程序没有 Office JavaScript Api 的同步版本。标准 Office 脚本 api 对平台是唯一的，并进行了大量优化和变更，以避免使用 `load` / `sync` 范例。</span><span class="sxs-lookup"><span data-stu-id="c186a-120">There is no synchronous version of the Office JavaScript APIs for Office Add-ins. The standard Office Scripts APIs are unique to the platform and have numerous optimizations and alterations to avoid the usage of the `load`/`sync` paradigm.</span></span>

<span data-ttu-id="c186a-121">某些[Excel JavaScript api](/javascript/api/excel?view=excel-js-preview)与[Office 脚本异步 api](../develop/excel-async-model.md)兼容。</span><span class="sxs-lookup"><span data-stu-id="c186a-121">Some of the [Excel JavaScript APIs](/javascript/api/excel?view=excel-js-preview) are compatible with the [Office Scripts Async APIs](../develop/excel-async-model.md).</span></span> <span data-ttu-id="c186a-122">某些示例和外接代码块可以 `Excel.run` 通过最少的转换移植到块。</span><span class="sxs-lookup"><span data-stu-id="c186a-122">Some samples and add-in code blocks could be ported to `Excel.run` blocks with minimal translation.</span></span> <span data-ttu-id="c186a-123">虽然这两个平台共享功能，但有一些缺口。</span><span class="sxs-lookup"><span data-stu-id="c186a-123">While the two platforms share functionality, there are gaps.</span></span> <span data-ttu-id="c186a-124">Office 外接程序设置了两个主要 API，但 Office 脚本不是事件和常见 Api。</span><span class="sxs-lookup"><span data-stu-id="c186a-124">The two major API sets that Office Add-ins have but Office Scripts do not are events and the Common APIs.</span></span>

### <a name="events"></a><span data-ttu-id="c186a-125">活动</span><span class="sxs-lookup"><span data-stu-id="c186a-125">Events</span></span>

<span data-ttu-id="c186a-126">Office 脚本不支持[事件](/office/dev/add-ins/excel/excel-add-ins-events)。</span><span class="sxs-lookup"><span data-stu-id="c186a-126">Office Scripts do not support [events](/office/dev/add-ins/excel/excel-add-ins-events).</span></span> <span data-ttu-id="c186a-127">每个脚本在一个方法中运行代码 `main` ，然后结束。</span><span class="sxs-lookup"><span data-stu-id="c186a-127">Every script runs the code in a single `main` method, then ends.</span></span> <span data-ttu-id="c186a-128">触发事件时不会重新激活，因此无法注册事件。</span><span class="sxs-lookup"><span data-stu-id="c186a-128">It does not reactivate when events are triggered, and thus, cannot register events.</span></span>

### <a name="common-apis"></a><span data-ttu-id="c186a-129">通用 API</span><span class="sxs-lookup"><span data-stu-id="c186a-129">Common APIs</span></span>

<span data-ttu-id="c186a-130">Office 脚本无法使用[通用 api](/javascript/api/office)。</span><span class="sxs-lookup"><span data-stu-id="c186a-130">Office Scripts cannot use [Common APIs](/javascript/api/office).</span></span> <span data-ttu-id="c186a-131">如果需要身份验证、对话窗口或其他仅受常见 Api 支持的功能，则您可能需要创建 Office 加载项，而不是 Office 脚本。</span><span class="sxs-lookup"><span data-stu-id="c186a-131">If you need authentication, dialog windows, or other features that are only supported by Common APIs, you'll likely need to create an Office Add-in instead of an Office Script.</span></span>

## <a name="see-also"></a><span data-ttu-id="c186a-132">另请参阅</span><span class="sxs-lookup"><span data-stu-id="c186a-132">See also</span></span>

- [<span data-ttu-id="c186a-133">Excel web 版中的 Office 脚本</span><span class="sxs-lookup"><span data-stu-id="c186a-133">Office Scripts in Excel on the web</span></span>](../overview/excel.md)
- [<span data-ttu-id="c186a-134">Office 脚本和 VBA 宏之间的区别</span><span class="sxs-lookup"><span data-stu-id="c186a-134">Differences between Office Scripts and VBA macros</span></span>](vba-differences.md)
- [<span data-ttu-id="c186a-135">Office 脚本疑难解答</span><span class="sxs-lookup"><span data-stu-id="c186a-135">Troubleshooting Office Scripts</span></span>](../testing/troubleshooting.md)
- [<span data-ttu-id="c186a-136">生成 Excel 任务窗格加载项</span><span class="sxs-lookup"><span data-stu-id="c186a-136">Build an Excel task pane add-in</span></span>](/office/dev/add-ins/quickstarts/excel-quickstart-jquery)
