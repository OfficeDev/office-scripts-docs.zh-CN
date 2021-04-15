---
title: Office 脚本与 Office 加载项之间的差异
description: Office 脚本和 Office 外接程序之间的行为和 API 差异。
ms.date: 06/01/2020
localization_priority: Normal
ms.openlocfilehash: 96af98ca9f247406c5cc916f38892c318d33c560
ms.sourcegitcommit: 45ffe3dbd2c834b78592ad35928cf8096f5e80bc
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 04/14/2021
ms.locfileid: "51755096"
---
# <a name="differences-between-office-scripts-and-office-add-ins"></a><span data-ttu-id="8e3a9-103">Office 脚本与 Office 加载项之间的差异</span><span class="sxs-lookup"><span data-stu-id="8e3a9-103">Differences between Office Scripts and Office Add-ins</span></span>

<span data-ttu-id="8e3a9-104">Office 外接程序和 Office 脚本有很多共同之处。</span><span class="sxs-lookup"><span data-stu-id="8e3a9-104">Office Add-ins and Office Scripts have a lot in common.</span></span> <span data-ttu-id="8e3a9-105">它们均提供对 Excel 工作簿的 JavaScript API 的自动化控制。</span><span class="sxs-lookup"><span data-stu-id="8e3a9-105">They both offer automated control of an Excel workbook a JavaScript API.</span></span> <span data-ttu-id="8e3a9-106">但是，Office 脚本 API 是 Office JavaScript API 的专用同步版本。</span><span class="sxs-lookup"><span data-stu-id="8e3a9-106">However, the Office Scripts APIs are a specialized, synchronous version of the Office JavaScript API.</span></span>

:::image type="content" source="../images/office-programmability-diagram.png" alt-text="一个四象限图表，显示不同 Office 扩展性解决方案的重点区域。Office 脚本和 Office Web 外接程序均侧重于 Web 和协作，但 Office 脚本适合最终用户 (而 Office Web 外接程序面向专业开发人员) 。":::

<span data-ttu-id="8e3a9-108">Office 脚本通过手动按下按钮或作为 [Power Automate](https://flow.microsoft.com/)中的一个步骤运行以完成，而 Office 外接程序在任务窗格打开时仍然存在。</span><span class="sxs-lookup"><span data-stu-id="8e3a9-108">Office Scripts run to completion with a manual button press or as a step in [Power Automate](https://flow.microsoft.com/), whereas Office Add-ins persist while their task panes are open.</span></span> <span data-ttu-id="8e3a9-109">这意味着加载项可以在会话期间保持状态，而 Office 脚本不会在两次运行之间保持内部状态。</span><span class="sxs-lookup"><span data-stu-id="8e3a9-109">This means the add-ins can maintain state during a session, whereas Office Scripts do not maintain an internal state between runs.</span></span> <span data-ttu-id="8e3a9-110">如果发现 Excel 扩展需要超过脚本平台的功能，请访问 [Office](/office/dev/add-ins) 加载项文档，详细了解 Office 加载项。</span><span class="sxs-lookup"><span data-stu-id="8e3a9-110">If you find that your Excel extension needs to exceed the scripting platform's capabilities, visit the [Office Add-ins documentation](/office/dev/add-ins) to learn more about Office Add-ins.</span></span>

<span data-ttu-id="8e3a9-111">本文的其余部分将介绍 Office 外接程序和 Office 脚本之间的主要差异。</span><span class="sxs-lookup"><span data-stu-id="8e3a9-111">The rest of this article describes on the main differences between Office Add-ins and Office Scripts.</span></span>

## <a name="platform-support"></a><span data-ttu-id="8e3a9-112">平台支持</span><span class="sxs-lookup"><span data-stu-id="8e3a9-112">Platform Support</span></span>

<span data-ttu-id="8e3a9-113">Office 外接程序是跨平台的。</span><span class="sxs-lookup"><span data-stu-id="8e3a9-113">Office Add-ins are cross-platform.</span></span> <span data-ttu-id="8e3a9-114">它们跨 Windows 桌面、Mac、iOS 和 Web 平台运行，并且在每个平台上提供相同的体验。</span><span class="sxs-lookup"><span data-stu-id="8e3a9-114">They work across Windows desktop, Mac, iOS, and web platforms and provide the same experience on each.</span></span> <span data-ttu-id="8e3a9-115">有关此情况的任何例外情况都记录在单个 API 的文档中。</span><span class="sxs-lookup"><span data-stu-id="8e3a9-115">Any exception to this is noted in the documentation of the individual API.</span></span>

<span data-ttu-id="8e3a9-116">Office 脚本当前仅受 Excel 网页版本支持。</span><span class="sxs-lookup"><span data-stu-id="8e3a9-116">Office Scripts are currently only supported by for Excel on the web.</span></span> <span data-ttu-id="8e3a9-117">所有录制、编辑和运行均在 Web 平台上完成。</span><span class="sxs-lookup"><span data-stu-id="8e3a9-117">All recording, editing, and running is done on the web platform.</span></span>

## <a name="apis"></a><span data-ttu-id="8e3a9-118">API</span><span class="sxs-lookup"><span data-stu-id="8e3a9-118">APIs</span></span>

<span data-ttu-id="8e3a9-119">没有适用于 Office 外接程序的 Office JavaScript API 的同步版本。标准 Office 脚本 API 对于平台是唯一的，并且具有大量优化和更改，以避免使用 `load` / `sync` 范例。</span><span class="sxs-lookup"><span data-stu-id="8e3a9-119">There is no synchronous version of the Office JavaScript APIs for Office Add-ins. The standard Office Scripts APIs are unique to the platform and have numerous optimizations and alterations to avoid the usage of the `load`/`sync` paradigm.</span></span>

<span data-ttu-id="8e3a9-120">某些 [Excel JavaScript API](/javascript/api/excel?view=excel-js-preview&preserve-view=true) 与 Office [脚本异步 API 兼容](../develop/excel-async-model.md)。</span><span class="sxs-lookup"><span data-stu-id="8e3a9-120">Some of the [Excel JavaScript APIs](/javascript/api/excel?view=excel-js-preview&preserve-view=true) are compatible with the [Office Scripts Async APIs](../develop/excel-async-model.md).</span></span> <span data-ttu-id="8e3a9-121">一些示例和外接程序代码块可以移植到 `Excel.run` 转换最少的块。</span><span class="sxs-lookup"><span data-stu-id="8e3a9-121">Some samples and add-in code blocks could be ported to `Excel.run` blocks with minimal translation.</span></span> <span data-ttu-id="8e3a9-122">虽然这两个平台共享功能，但存在一些差异。</span><span class="sxs-lookup"><span data-stu-id="8e3a9-122">While the two platforms share functionality, there are gaps.</span></span> <span data-ttu-id="8e3a9-123">Office 外接程序具有的两个主要 API 集，但 Office 脚本不是事件和通用 API。</span><span class="sxs-lookup"><span data-stu-id="8e3a9-123">The two major API sets that Office Add-ins have but Office Scripts do not are events and the Common APIs.</span></span>

### <a name="events"></a><span data-ttu-id="8e3a9-124">活动</span><span class="sxs-lookup"><span data-stu-id="8e3a9-124">Events</span></span>

<span data-ttu-id="8e3a9-125">Office 脚本不支持 [事件](/office/dev/add-ins/excel/excel-add-ins-events)。</span><span class="sxs-lookup"><span data-stu-id="8e3a9-125">Office Scripts do not support [events](/office/dev/add-ins/excel/excel-add-ins-events).</span></span> <span data-ttu-id="8e3a9-126">每个脚本在一个方法中运行 `main` 代码，然后结束。</span><span class="sxs-lookup"><span data-stu-id="8e3a9-126">Every script runs the code in a single `main` method, then ends.</span></span> <span data-ttu-id="8e3a9-127">它不会在触发事件时重新激活，因此无法注册事件。</span><span class="sxs-lookup"><span data-stu-id="8e3a9-127">It does not reactivate when events are triggered, and thus, cannot register events.</span></span>

### <a name="common-apis"></a><span data-ttu-id="8e3a9-128">通用 API</span><span class="sxs-lookup"><span data-stu-id="8e3a9-128">Common APIs</span></span>

<span data-ttu-id="8e3a9-129">Office 脚本不能使用[通用 API。](/javascript/api/office)</span><span class="sxs-lookup"><span data-stu-id="8e3a9-129">Office Scripts cannot use [Common APIs](/javascript/api/office).</span></span> <span data-ttu-id="8e3a9-130">如果您需要身份验证、对话框窗口或其他仅受通用 API 支持的功能，您可能需要创建 Office 外接程序而不是 Office 脚本。</span><span class="sxs-lookup"><span data-stu-id="8e3a9-130">If you need authentication, dialog windows, or other features that are only supported by Common APIs, you'll likely need to create an Office Add-in instead of an Office Script.</span></span>

## <a name="see-also"></a><span data-ttu-id="8e3a9-131">另请参阅</span><span class="sxs-lookup"><span data-stu-id="8e3a9-131">See also</span></span>

- [<span data-ttu-id="8e3a9-132">Excel web 版中的 Office 脚本</span><span class="sxs-lookup"><span data-stu-id="8e3a9-132">Office Scripts in Excel on the web</span></span>](../overview/excel.md)
- [<span data-ttu-id="8e3a9-133">Office 脚本和 VBA 宏之间的差异</span><span class="sxs-lookup"><span data-stu-id="8e3a9-133">Differences between Office Scripts and VBA macros</span></span>](vba-differences.md)
- [<span data-ttu-id="8e3a9-134">Office 脚本疑难解答</span><span class="sxs-lookup"><span data-stu-id="8e3a9-134">Troubleshooting Office Scripts</span></span>](../testing/troubleshooting.md)
- [<span data-ttu-id="8e3a9-135">生成 Excel 任务窗格加载项</span><span class="sxs-lookup"><span data-stu-id="8e3a9-135">Build an Excel task pane add-in</span></span>](/office/dev/add-ins/quickstarts/excel-quickstart-jquery)
