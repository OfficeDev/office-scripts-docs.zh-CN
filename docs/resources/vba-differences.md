---
title: Office 脚本和 VBA 宏之间的区别
description: Office 脚本和 Excel VBA 宏之间的行为和 API 差异。
ms.date: 06/30/2020
localization_priority: Normal
ms.openlocfilehash: 8a8929f0c6a73a8e9041bb4b55cce1edd539e166
ms.sourcegitcommit: bf9f33c37c6f7805d6b408aa648bb9785a7cd133
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 07/06/2020
ms.locfileid: "45043389"
---
# <a name="differences-between-office-scripts-and-vba-macros"></a><span data-ttu-id="e4986-103">Office 脚本和 VBA 宏之间的区别</span><span class="sxs-lookup"><span data-stu-id="e4986-103">Differences between Office Scripts and VBA macros</span></span>

<span data-ttu-id="e4986-104">Office 脚本和 VBA 宏具有很多共同之处。</span><span class="sxs-lookup"><span data-stu-id="e4986-104">Office Scripts and VBA macros have a lot in common.</span></span> <span data-ttu-id="e4986-105">它们都允许用户通过易于使用的操作录制器自动执行解决方案，并允许编辑这些录制。</span><span class="sxs-lookup"><span data-stu-id="e4986-105">They both allow users to automate solutions through an easy-to-use action recorder and allow edits of those recordings.</span></span> <span data-ttu-id="e4986-106">这两个框架都旨在让可能不会考虑自己的程序员在 Excel 中创建小型程序的人员。</span><span class="sxs-lookup"><span data-stu-id="e4986-106">Both frameworks are designed to empower people who may not consider themselves programmers to create small programs in Excel.</span></span>
<span data-ttu-id="e4986-107">基本区别在于，VBA 宏是为桌面解决方案开发的，而 Office 脚本是通过跨平台支持和安全性设计的，以指导原则为依据。</span><span class="sxs-lookup"><span data-stu-id="e4986-107">The fundamental difference is that VBA macros are developed for desktop solutions and Office Scripts are designed with cross-platform support and security as the guiding principles.</span></span> <span data-ttu-id="e4986-108">目前，仅在 web 上的 Excel 中支持 Office 脚本。</span><span class="sxs-lookup"><span data-stu-id="e4986-108">Currently, Office Scripts are only supported in Excel on the web.</span></span>

![显示不同 Office 扩展性解决方案的重点领域的四象限图。](../images/office-programmability-diagram.png)

<span data-ttu-id="e4986-111">本文介绍 VBA 宏（以及常规的 VBA）和 Office 脚本之间的主要差异。</span><span class="sxs-lookup"><span data-stu-id="e4986-111">This article describes the main differences between VBA macros (as well as VBA in general) and Office Scripts.</span></span> <span data-ttu-id="e4986-112">由于 Office 脚本仅适用于 Excel，所以这里只讨论唯一的主机。</span><span class="sxs-lookup"><span data-stu-id="e4986-112">Since Office Scripts are only available for Excel, that is the only host being discussed here.</span></span>

## <a name="platform-and-ecosystem"></a><span data-ttu-id="e4986-113">平台和生态系统</span><span class="sxs-lookup"><span data-stu-id="e4986-113">Platform and ecosystem</span></span>

<span data-ttu-id="e4986-114">VBA 设计用于桌面，而 Office 脚本是为 web 设计的。</span><span class="sxs-lookup"><span data-stu-id="e4986-114">VBA is designed for the desktop and Office Scripts are designed for the web.</span></span> <span data-ttu-id="e4986-115">VBA 可以与用户桌面进行交互，以与类似的技术（如 COM 和 OLE）进行连接。</span><span class="sxs-lookup"><span data-stu-id="e4986-115">VBA can interact with a user's desktop to connect with similar technologies, such as COM and OLE.</span></span> <span data-ttu-id="e4986-116">但是，VBA 无法方便地调用到 internet。</span><span class="sxs-lookup"><span data-stu-id="e4986-116">However, VBA has no convenient way to call out to the internet.</span></span>

<span data-ttu-id="e4986-117">Office 脚本使用通用运行时或 JavaScript。</span><span class="sxs-lookup"><span data-stu-id="e4986-117">Office Scripts use a universal runtime or JavaScript.</span></span> <span data-ttu-id="e4986-118">这将提供一致的行为和可访问性，而无需考虑用于运行脚本的计算机。</span><span class="sxs-lookup"><span data-stu-id="e4986-118">This gives consistent behavior and accessibility, regardless of the machine being used to run the script.</span></span> <span data-ttu-id="e4986-119">它们还可以调用其他 web 服务。</span><span class="sxs-lookup"><span data-stu-id="e4986-119">They can also make calls to other web services.</span></span>

## <a name="security"></a><span data-ttu-id="e4986-120">安全性</span><span class="sxs-lookup"><span data-stu-id="e4986-120">Security</span></span>

<span data-ttu-id="e4986-121">VBA 宏与 Excel 具有相同的安全净空。</span><span class="sxs-lookup"><span data-stu-id="e4986-121">VBA macros have the same security clearance as Excel.</span></span> <span data-ttu-id="e4986-122">这样，他们就可以拥有对桌面的完全访问权限。</span><span class="sxs-lookup"><span data-stu-id="e4986-122">This gives them full access to your desktop.</span></span> <span data-ttu-id="e4986-123">Office 脚本仅具有对工作簿的访问权限，而不是承载工作簿的计算机。</span><span class="sxs-lookup"><span data-stu-id="e4986-123">Office Scripts only have access to the workbook, not the machine hosting the workbook.</span></span> <span data-ttu-id="e4986-124">此外，不能与脚本共享任何 JavaScript 身份验证令牌，因此脚本永远不能通过外部服务进行身份验证。</span><span class="sxs-lookup"><span data-stu-id="e4986-124">Additionally, no JavaScript authentication tokens can be shared with scripts, so scripts can never authenticate with an external service.</span></span>

<span data-ttu-id="e4986-125">管理员具有三个 VBA 宏选项：允许租户上的所有宏、不允许租户上的宏，或仅允许带有签名证书的宏。</span><span class="sxs-lookup"><span data-stu-id="e4986-125">Admins have three options for VBA macros: allow all macros on the tenant, allow no macros on the tenant, or allow only macros with signed certificates.</span></span> <span data-ttu-id="e4986-126">这种缺乏的粒度使得难以隔离单个损坏的主角。</span><span class="sxs-lookup"><span data-stu-id="e4986-126">This lack of granularity makes it hard to isolate a single bad actor.</span></span> <span data-ttu-id="e4986-127">目前，Office 脚本是针对租户的 "打开" 或 "关闭"。</span><span class="sxs-lookup"><span data-stu-id="e4986-127">Currently, Office Scripts are either on or off for a tenant.</span></span> <span data-ttu-id="e4986-128">不过，我们正在努力为管理员提供对各个脚本和脚本编写者的更多控制。</span><span class="sxs-lookup"><span data-stu-id="e4986-128">However, we are working to give admins more control over individual scripts and script creators.</span></span>

## <a name="coverage"></a><span data-ttu-id="e4986-129">报道</span><span class="sxs-lookup"><span data-stu-id="e4986-129">Coverage</span></span>

<span data-ttu-id="e4986-130">目前，VBA 提供了更全面的 Excel 功能，尤其是在桌面客户端上提供的功能。</span><span class="sxs-lookup"><span data-stu-id="e4986-130">Currently, VBA offers a more complete coverage of Excel features, particularly those available on the desktop client.</span></span> <span data-ttu-id="e4986-131">Office 脚本涵盖了 web 上的 Excel 的几乎所有方案。</span><span class="sxs-lookup"><span data-stu-id="e4986-131">Office Scripts cover nearly all of the scenarios for Excel on the web.</span></span> <span data-ttu-id="e4986-132">此外，在 web 上 debut 新功能时，Office 脚本将同时为操作记录器和 JavaScript Api 支持这些功能。</span><span class="sxs-lookup"><span data-stu-id="e4986-132">Additionally, as new features debut on the web, Office Scripts will support them for both the Action Recorder and JavaScript APIs.</span></span>

## <a name="power-automate"></a><span data-ttu-id="e4986-133">Power Automate</span><span class="sxs-lookup"><span data-stu-id="e4986-133">Power Automate</span></span>

<span data-ttu-id="e4986-134">可以通过 "Power 自动化" 运行 Office 脚本。</span><span class="sxs-lookup"><span data-stu-id="e4986-134">Office Scripts can be run through Power Automate.</span></span> <span data-ttu-id="e4986-135">您的工作簿可以通过计划或事件驱动的流进行更新，让您无需打开 Excel 即可自动执行工作流。</span><span class="sxs-lookup"><span data-stu-id="e4986-135">Your workbook can be updated through scheduled or event-driven flows, letting you automate workflows without even opening Excel.</span></span> <span data-ttu-id="e4986-136">这意味着只要您的工作簿存储在 OneDrive 中（并可供电源自动访问），流就可以运行您的脚本，而不管您和您的组织使用的是 Excel 的桌面、Mac 还是 web 客户端。</span><span class="sxs-lookup"><span data-stu-id="e4986-136">This means that as long as your workbook is stored in OneDrive (and accessible to Power Automate), a flow can run your scripts regardless of whether you and your organization use Excel's desktop, Mac, or web client.</span></span>

<span data-ttu-id="e4986-137">VBA 没有电源自动连接器。</span><span class="sxs-lookup"><span data-stu-id="e4986-137">VBA doesn't have a Power Automate connector.</span></span> <span data-ttu-id="e4986-138">所有受支持的 VBA 方案都涉及用户参与宏的执行。</span><span class="sxs-lookup"><span data-stu-id="e4986-138">All supported VBA scenarios involved a user attending to the macro's execution.</span></span>

## <a name="see-also"></a><span data-ttu-id="e4986-139">另请参阅</span><span class="sxs-lookup"><span data-stu-id="e4986-139">See also</span></span>

- [<span data-ttu-id="e4986-140">Excel web 版中的 Office 脚本</span><span class="sxs-lookup"><span data-stu-id="e4986-140">Office Scripts in Excel on the web</span></span>](../overview/excel.md)
- [<span data-ttu-id="e4986-141">Office 脚本与 Office 加载项之间的差异</span><span class="sxs-lookup"><span data-stu-id="e4986-141">Differences between Office Scripts and Office Add-ins</span></span>](add-ins-differences.md)
- [<span data-ttu-id="e4986-142">Office 脚本疑难解答</span><span class="sxs-lookup"><span data-stu-id="e4986-142">Troubleshooting Office Scripts</span></span>](../testing/troubleshooting.md)
- [<span data-ttu-id="e4986-143">Excel VBA 参考</span><span class="sxs-lookup"><span data-stu-id="e4986-143">Excel VBA reference</span></span>](/office/vba/api/overview/excel)
