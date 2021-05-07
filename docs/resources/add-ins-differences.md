---
title: Office 脚本与 Office 加载项之间的差异
description: 脚本和加载项Office API 的行为Office API 差异。
ms.date: 06/01/2020
localization_priority: Normal
ms.openlocfilehash: 45993d08d85cfceb299216dddbe2e7da9fd2e404
ms.sourcegitcommit: f7a7aebfb687f2a35dbed07ed62ff352a114525a
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 05/06/2021
ms.locfileid: "52232632"
---
# <a name="differences-between-office-scripts-and-office-add-ins"></a><span data-ttu-id="050ec-103">Office 脚本与 Office 加载项之间的差异</span><span class="sxs-lookup"><span data-stu-id="050ec-103">Differences between Office Scripts and Office Add-ins</span></span>

<span data-ttu-id="050ec-104">Office加载项和Office脚本有很多共同之处。</span><span class="sxs-lookup"><span data-stu-id="050ec-104">Office Add-ins and Office Scripts have a lot in common.</span></span> <span data-ttu-id="050ec-105">它们都提供对 JavaScript API Excel工作簿的自动化控制。</span><span class="sxs-lookup"><span data-stu-id="050ec-105">They both offer automated control of an Excel workbook a JavaScript API.</span></span> <span data-ttu-id="050ec-106">但是，Office脚本 API 是 JavaScript API 的专用Office同步版本。</span><span class="sxs-lookup"><span data-stu-id="050ec-106">However, the Office Scripts APIs are a specialized, synchronous version of the Office JavaScript API.</span></span>

:::image type="content" source="../images/office-programmability-diagram.png" alt-text="四象限图，显示不同扩展性解决方案Office区域。Office 脚本和 Office Web 外接程序均侧重于 Web 和协作，但 Office 脚本适合最终用户 (而 Office Web 外接程序面向专业开发人员) ":::

<span data-ttu-id="050ec-108">Office脚本通过手动按下按钮或作为 Power Automate 中的步骤运行以[](https://flow.microsoft.com/)完成，Office任务窗格打开时，外接程序将保持运行状态。</span><span class="sxs-lookup"><span data-stu-id="050ec-108">Office Scripts run to completion with a manual button press or as a step in [Power Automate](https://flow.microsoft.com/), whereas Office Add-ins persist while their task panes are open.</span></span> <span data-ttu-id="050ec-109">这意味着加载项可以在会话期间保持状态，而Office脚本不会在两次运行之间保持内部状态。</span><span class="sxs-lookup"><span data-stu-id="050ec-109">This means the add-ins can maintain state during a session, whereas Office Scripts do not maintain an internal state between runs.</span></span> <span data-ttu-id="050ec-110">如果您发现您的 Excel 扩展需要超过脚本平台的功能，请访问[Office 外接程序](/office/dev/add-ins)文档以了解有关 Office 外接程序的信息。</span><span class="sxs-lookup"><span data-stu-id="050ec-110">If you find that your Excel extension needs to exceed the scripting platform's capabilities, visit the [Office Add-ins documentation](/office/dev/add-ins) to learn more about Office Add-ins.</span></span>

<span data-ttu-id="050ec-111">本文的其余部分介绍加载项和脚本Office之间的主要Office区别。</span><span class="sxs-lookup"><span data-stu-id="050ec-111">The rest of this article describes on the main differences between Office Add-ins and Office Scripts.</span></span>

## <a name="platform-support"></a><span data-ttu-id="050ec-112">平台支持</span><span class="sxs-lookup"><span data-stu-id="050ec-112">Platform Support</span></span>

<span data-ttu-id="050ec-113">Office外接程序是跨平台的。</span><span class="sxs-lookup"><span data-stu-id="050ec-113">Office Add-ins are cross-platform.</span></span> <span data-ttu-id="050ec-114">它们跨桌面Windows、Mac、iOS 和 Web 平台运行，并在每个平台上提供相同的体验。</span><span class="sxs-lookup"><span data-stu-id="050ec-114">They work across Windows desktop, Mac, iOS, and web platforms and provide the same experience on each.</span></span> <span data-ttu-id="050ec-115">有关此情况的任何例外情况都记录在单个 API 的文档中。</span><span class="sxs-lookup"><span data-stu-id="050ec-115">Any exception to this is noted in the documentation of the individual API.</span></span>

<span data-ttu-id="050ec-116">Office脚本当前仅受 Excel web 版。</span><span class="sxs-lookup"><span data-stu-id="050ec-116">Office Scripts are currently only supported by for Excel on the web.</span></span> <span data-ttu-id="050ec-117">所有录制、编辑和运行均在 Web 平台上完成。</span><span class="sxs-lookup"><span data-stu-id="050ec-117">All recording, editing, and running is done on the web platform.</span></span>

## <a name="apis"></a><span data-ttu-id="050ec-118">API</span><span class="sxs-lookup"><span data-stu-id="050ec-118">APIs</span></span>

<span data-ttu-id="050ec-119">没有适用于外接程序的 Office JavaScript API Office版本。标准Office脚本 API 对于平台是唯一的，并且具有许多优化和更改以避免使用 `load` / `sync` 范例。</span><span class="sxs-lookup"><span data-stu-id="050ec-119">There is no synchronous version of the Office JavaScript APIs for Office Add-ins. The standard Office Scripts APIs are unique to the platform and have numerous optimizations and alterations to avoid the usage of the `load`/`sync` paradigm.</span></span>

<span data-ttu-id="050ec-120">一些[Excel JavaScript API](/javascript/api/excel?view=excel-js-preview&preserve-view=true)与 Office[脚本异步 API 兼容](../develop/excel-async-model.md)。</span><span class="sxs-lookup"><span data-stu-id="050ec-120">Some of the [Excel JavaScript APIs](/javascript/api/excel?view=excel-js-preview&preserve-view=true) are compatible with the [Office Scripts Async APIs](../develop/excel-async-model.md).</span></span> <span data-ttu-id="050ec-121">一些示例和外接程序代码块可以移植到 `Excel.run` 转换最少的块。</span><span class="sxs-lookup"><span data-stu-id="050ec-121">Some samples and add-in code blocks could be ported to `Excel.run` blocks with minimal translation.</span></span> <span data-ttu-id="050ec-122">虽然这两个平台共享功能，但存在一些差异。</span><span class="sxs-lookup"><span data-stu-id="050ec-122">While the two platforms share functionality, there are gaps.</span></span> <span data-ttu-id="050ec-123">加载项具有但Office脚本的两个主要 API 集Office事件和通用 API。</span><span class="sxs-lookup"><span data-stu-id="050ec-123">The two major API sets that Office Add-ins have but Office Scripts do not are events and the Common APIs.</span></span>

### <a name="events"></a><span data-ttu-id="050ec-124">事件</span><span class="sxs-lookup"><span data-stu-id="050ec-124">Events</span></span>

<span data-ttu-id="050ec-125">Office脚本不支持[事件](/office/dev/add-ins/excel/excel-add-ins-events)。</span><span class="sxs-lookup"><span data-stu-id="050ec-125">Office Scripts do not support [events](/office/dev/add-ins/excel/excel-add-ins-events).</span></span> <span data-ttu-id="050ec-126">每个脚本在一个方法中运行 `main` 代码，然后结束。</span><span class="sxs-lookup"><span data-stu-id="050ec-126">Every script runs the code in a single `main` method, then ends.</span></span> <span data-ttu-id="050ec-127">它不会在触发事件时重新激活，因此无法注册事件。</span><span class="sxs-lookup"><span data-stu-id="050ec-127">It does not reactivate when events are triggered, and thus, cannot register events.</span></span>

### <a name="common-apis"></a><span data-ttu-id="050ec-128">通用 API</span><span class="sxs-lookup"><span data-stu-id="050ec-128">Common APIs</span></span>

<span data-ttu-id="050ec-129">Office脚本不能使用[通用 API。](/javascript/api/office)</span><span class="sxs-lookup"><span data-stu-id="050ec-129">Office Scripts cannot use [Common APIs](/javascript/api/office).</span></span> <span data-ttu-id="050ec-130">如果你需要身份验证、对话框窗口或其他仅受通用 API 支持的功能，你可能需要创建一个 Office 外接程序，而不是一个 Office 脚本。</span><span class="sxs-lookup"><span data-stu-id="050ec-130">If you need authentication, dialog windows, or other features that are only supported by Common APIs, you'll likely need to create an Office Add-in instead of an Office Script.</span></span>

## <a name="see-also"></a><span data-ttu-id="050ec-131">另请参阅</span><span class="sxs-lookup"><span data-stu-id="050ec-131">See also</span></span>

- [<span data-ttu-id="050ec-132">Excel web 版中的 Office 脚本</span><span class="sxs-lookup"><span data-stu-id="050ec-132">Office Scripts in Excel on the web</span></span>](../overview/excel.md)
- [<span data-ttu-id="050ec-133">脚本Office VBA 宏之间的差异</span><span class="sxs-lookup"><span data-stu-id="050ec-133">Differences between Office Scripts and VBA macros</span></span>](vba-differences.md)
- [<span data-ttu-id="050ec-134">Office 脚本疑难解答</span><span class="sxs-lookup"><span data-stu-id="050ec-134">Troubleshooting Office Scripts</span></span>](../testing/troubleshooting.md)
- [<span data-ttu-id="050ec-135">生成 Excel 任务窗格加载项</span><span class="sxs-lookup"><span data-stu-id="050ec-135">Build an Excel task pane add-in</span></span>](/office/dev/add-ins/quickstarts/excel-quickstart-jquery)
