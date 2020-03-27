---
title: Office 脚本和 VBA 宏之间的区别
description: Office 脚本和 Excel VBA 宏之间的行为和 API 差异。
ms.date: 03/23/2020
localization_priority: Normal
ms.openlocfilehash: 3a0f2c9a2ed7181a10e41d1f45b3af695877a680
ms.sourcegitcommit: d556aaefac80e55f53ac56b7f6ecbc657ebd426f
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/26/2020
ms.locfileid: "42978700"
---
# <a name="differences-between-office-scripts-and-vba-macros"></a><span data-ttu-id="1439b-103">Office 脚本和 VBA 宏之间的区别</span><span class="sxs-lookup"><span data-stu-id="1439b-103">Differences between Office Scripts and VBA macros</span></span>

<span data-ttu-id="1439b-104">Office 脚本和 VBA 宏具有很多共同之处。</span><span class="sxs-lookup"><span data-stu-id="1439b-104">Office Scripts and VBA macros have a lot in common.</span></span> <span data-ttu-id="1439b-105">它们都允许用户通过易于使用的操作录制器自动执行解决方案，并允许编辑这些录制。</span><span class="sxs-lookup"><span data-stu-id="1439b-105">They both allow users to automate solutions through an easy-to-use action recorder and allow edits of those recordings.</span></span> <span data-ttu-id="1439b-106">这两个框架都旨在让可能不会考虑自己的程序员在 Excel 中创建小型程序的人员。</span><span class="sxs-lookup"><span data-stu-id="1439b-106">Both frameworks are designed to empower people who may not consider themselves programmers to create small programs in Excel.</span></span>
<span data-ttu-id="1439b-107">基本区别在于，VBA 宏是为桌面解决方案开发的，而 Office 脚本是通过跨平台支持和安全性设计的，以指导原则为依据。</span><span class="sxs-lookup"><span data-stu-id="1439b-107">The fundamental difference is that VBA macros are developed for desktop solutions and Office Scripts are designed with cross-platform support and security as the guiding principles.</span></span> <span data-ttu-id="1439b-108">目前，仅在 web 上的 Excel 中支持 Office 脚本。</span><span class="sxs-lookup"><span data-stu-id="1439b-108">Currently, Office Scripts are only supported in Excel on the web.</span></span>

![显示不同 Office 扩展性解决方案的重点领域的四象限图。](../images/office-programmability-diagram.png)

<span data-ttu-id="1439b-111">本文介绍 VBA 宏（以及常规的 VBA）和 Office 脚本之间的主要差异。</span><span class="sxs-lookup"><span data-stu-id="1439b-111">This article describes the main differences between VBA macros (as well as VBA in general) and Office Scripts.</span></span> <span data-ttu-id="1439b-112">由于 Office 脚本仅适用于 Excel，所以这里只讨论唯一的主机。</span><span class="sxs-lookup"><span data-stu-id="1439b-112">Since Office Scripts are only available for Excel, that is the only host being discussed here.</span></span>

## <a name="platform-and-ecosystem"></a><span data-ttu-id="1439b-113">平台和生态系统</span><span class="sxs-lookup"><span data-stu-id="1439b-113">Platform and ecosystem</span></span>

<span data-ttu-id="1439b-114">VBA 设计用于桌面，而 Office 脚本是为 web 设计的。</span><span class="sxs-lookup"><span data-stu-id="1439b-114">VBA is designed for the desktop and Office Scripts are designed for the web.</span></span> <span data-ttu-id="1439b-115">VBA 可以与用户桌面进行交互。</span><span class="sxs-lookup"><span data-stu-id="1439b-115">VBA can interact with a user's desktop.</span></span> <span data-ttu-id="1439b-116">这使其能够与类似的技术（如 COM 和 OLE）集成。</span><span class="sxs-lookup"><span data-stu-id="1439b-116">This lets it integrate with similar technologies, such as COM and OLE.</span></span> <span data-ttu-id="1439b-117">但是，VBA 无法方便地调用到 internet。</span><span class="sxs-lookup"><span data-stu-id="1439b-117">However, VBA has no convenient way to call out to the internet.</span></span>

<span data-ttu-id="1439b-118">Office 脚本使用通用运行时或 JavaScript。</span><span class="sxs-lookup"><span data-stu-id="1439b-118">Office Scripts use a universal runtime or JavaScript.</span></span> <span data-ttu-id="1439b-119">这将提供一致的行为和可访问性，而无需考虑用于运行脚本的计算机。</span><span class="sxs-lookup"><span data-stu-id="1439b-119">This gives consistent behavior and accessibility, regardless of the machine being used to run the script.</span></span> <span data-ttu-id="1439b-120">它们还可以调用其他 web 服务。</span><span class="sxs-lookup"><span data-stu-id="1439b-120">They can also make calls to other web services.</span></span>

## <a name="security"></a><span data-ttu-id="1439b-121">安全性</span><span class="sxs-lookup"><span data-stu-id="1439b-121">Security</span></span>

<span data-ttu-id="1439b-122">VBA 宏与 Excel 具有相同的安全净空。</span><span class="sxs-lookup"><span data-stu-id="1439b-122">VBA macros have the same security clearance as Excel.</span></span> <span data-ttu-id="1439b-123">这样，他们就可以拥有对桌面的完全访问权限。</span><span class="sxs-lookup"><span data-stu-id="1439b-123">This gives them full access to your desktop.</span></span> <span data-ttu-id="1439b-124">Office 脚本仅具有对工作簿的访问权限，而不是承载工作簿的计算机。</span><span class="sxs-lookup"><span data-stu-id="1439b-124">Office Scripts only have access to the workbook, not the machine hosting the workbook.</span></span> <span data-ttu-id="1439b-125">此外，不能与脚本共享任何 JavaScript 身份验证令牌，因此脚本永远不能通过外部服务进行身份验证。</span><span class="sxs-lookup"><span data-stu-id="1439b-125">Additionally, no JavaScript authentication tokens can be shared with scripts, so scripts can never authenticate with an external service.</span></span>

<span data-ttu-id="1439b-126">管理员具有三个 VBA 宏选项：允许租户上的所有宏、不允许租户上的宏，或仅允许带有签名证书的宏。</span><span class="sxs-lookup"><span data-stu-id="1439b-126">Admins have three options for VBA macros: allow all macros on the tenant, allow no macros on the tenant, or allow only macros with signed certificates.</span></span> <span data-ttu-id="1439b-127">这种缺乏的粒度使得难以隔离单个损坏的主角。</span><span class="sxs-lookup"><span data-stu-id="1439b-127">This lack of granularity makes it hard to isolate a single bad actor.</span></span> <span data-ttu-id="1439b-128">目前，Office 脚本是针对租户的 "打开" 或 "关闭"。</span><span class="sxs-lookup"><span data-stu-id="1439b-128">Currently, Office Scripts are either on or off for a tenant.</span></span> <span data-ttu-id="1439b-129">不过，我们正在努力为管理员提供对各个脚本和脚本编写者的更多控制。</span><span class="sxs-lookup"><span data-stu-id="1439b-129">However, we are working to give admins more control over individual scripts and script creators.</span></span>

## <a name="coverage"></a><span data-ttu-id="1439b-130">报道</span><span class="sxs-lookup"><span data-stu-id="1439b-130">Coverage</span></span>

<span data-ttu-id="1439b-131">目前，VBA 提供了更全面的 Excel 功能，尤其是在桌面客户端上提供的功能。</span><span class="sxs-lookup"><span data-stu-id="1439b-131">Currently, VBA offers a more complete coverage of Excel features, particularly those available on the desktop client.</span></span> <span data-ttu-id="1439b-132">Office 脚本涵盖了 web 上的 Excel 的几乎所有方案。</span><span class="sxs-lookup"><span data-stu-id="1439b-132">Office Scripts cover nearly all of the scenarios for Excel on the web.</span></span> <span data-ttu-id="1439b-133">此外，在 web 上 debut 新功能时，Office 脚本将同时为操作记录器和 JavaScript Api 支持这些功能。</span><span class="sxs-lookup"><span data-stu-id="1439b-133">Additionally, as new features debut on the web, Office Scripts will support them for both the Action Recorder and JavaScript APIs.</span></span>

## <a name="power-automate"></a><span data-ttu-id="1439b-134">电源自动化</span><span class="sxs-lookup"><span data-stu-id="1439b-134">Power Automate</span></span>

<span data-ttu-id="1439b-135">可以通过 "Power 自动化" 运行 Office 脚本。</span><span class="sxs-lookup"><span data-stu-id="1439b-135">Office Scripts can be run through Power Automate.</span></span> <span data-ttu-id="1439b-136">您的工作簿可以通过计划或事件驱动的流进行更新，让您无需打开 Excel 即可自动执行工作流。</span><span class="sxs-lookup"><span data-stu-id="1439b-136">Your workbook can be updated through scheduled or event-driven flows, letting you automate workflows without even opening Excel.</span></span> <span data-ttu-id="1439b-137">这意味着只要您的工作簿存储在 OneDrive 中（并可供电源自动访问），流就可以运行您的脚本，而不管您和您的组织使用的是 Excel 的桌面、Mac 还是 web 客户端。</span><span class="sxs-lookup"><span data-stu-id="1439b-137">This means that as long as your workbook is stored in OneDrive (and accessible to Power Automate), a flow can run your scripts regardless of whether you and your organization use Excel's desktop, Mac, or web client.</span></span>

<span data-ttu-id="1439b-138">VBA 没有与电源自动化的集成。</span><span class="sxs-lookup"><span data-stu-id="1439b-138">VBA has no integration with Power Automate.</span></span> <span data-ttu-id="1439b-139">所有受支持的 VBA 方案都涉及用户参与宏的执行。</span><span class="sxs-lookup"><span data-stu-id="1439b-139">All supported VBA scenarios involved a user attending to the macro's execution.</span></span>

## <a name="see-also"></a><span data-ttu-id="1439b-140">另请参阅</span><span class="sxs-lookup"><span data-stu-id="1439b-140">See also</span></span>

- [<span data-ttu-id="1439b-141">Excel 网页版中的 Office 脚本</span><span class="sxs-lookup"><span data-stu-id="1439b-141">Office Scripts in Excel on the web</span></span>](../overview/excel.md)
- [<span data-ttu-id="1439b-142">Office 脚本与 Office 外接程序之间的差异</span><span class="sxs-lookup"><span data-stu-id="1439b-142">Differences between Office Scripts and Office Add-ins</span></span>](add-ins-differences.md)
- [<span data-ttu-id="1439b-143">Office 脚本疑难解答</span><span class="sxs-lookup"><span data-stu-id="1439b-143">Troubleshooting Office Scripts</span></span>](../testing/troubleshooting.md)
- [<span data-ttu-id="1439b-144">Excel VBA 参考</span><span class="sxs-lookup"><span data-stu-id="1439b-144">Excel VBA reference</span></span>](/office/vba/api/overview/excel)
