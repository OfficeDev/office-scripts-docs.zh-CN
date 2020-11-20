---
title: Office 脚本和 VBA 宏之间的区别
description: Office 脚本和 Excel VBA 宏之间的行为和 API 差异。
ms.date: 11/13/2020
localization_priority: Normal
ms.openlocfilehash: 7b9186d03489a43836c6e9da7bd28e0abc135f63
ms.sourcegitcommit: 82d3c0ef1e187bcdeceb2b5fc3411186674fe150
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 11/18/2020
ms.locfileid: "49342883"
---
# <a name="differences-between-office-scripts-and-vba-macros"></a><span data-ttu-id="4aba4-103">Office 脚本和 VBA 宏之间的区别</span><span class="sxs-lookup"><span data-stu-id="4aba4-103">Differences between Office Scripts and VBA macros</span></span>

<span data-ttu-id="4aba4-104">Office 脚本和 VBA 宏具有很多共同之处。</span><span class="sxs-lookup"><span data-stu-id="4aba4-104">Office Scripts and VBA macros have a lot in common.</span></span> <span data-ttu-id="4aba4-105">它们都允许用户通过易于使用的操作录制器自动执行解决方案，并允许编辑这些录制。</span><span class="sxs-lookup"><span data-stu-id="4aba4-105">They both allow users to automate solutions through an easy-to-use action recorder and allow edits of those recordings.</span></span> <span data-ttu-id="4aba4-106">这两个框架都旨在让可能不会考虑自己的程序员在 Excel 中创建小型程序的人员。</span><span class="sxs-lookup"><span data-stu-id="4aba4-106">Both frameworks are designed to empower people who may not consider themselves programmers to create small programs in Excel.</span></span>
<span data-ttu-id="4aba4-107">基本区别在于，VBA 宏是为桌面解决方案开发的，而 Office 脚本是通过跨平台支持和安全性设计的，以指导原则为依据。</span><span class="sxs-lookup"><span data-stu-id="4aba4-107">The fundamental difference is that VBA macros are developed for desktop solutions and Office Scripts are designed with cross-platform support and security as the guiding principles.</span></span> <span data-ttu-id="4aba4-108">目前，仅在 web 上的 Excel 中支持 Office 脚本。</span><span class="sxs-lookup"><span data-stu-id="4aba4-108">Currently, Office Scripts are only supported in Excel on the web.</span></span>

![显示不同 Office 扩展性解决方案的重点领域的四象限图。](../images/office-programmability-diagram.png)

<span data-ttu-id="4aba4-111">本文介绍 VBA 宏 (的主要区别以及常规) 和 Office 脚本中的 VBA。</span><span class="sxs-lookup"><span data-stu-id="4aba4-111">This article describes the main differences between VBA macros (as well as VBA in general) and Office Scripts.</span></span> <span data-ttu-id="4aba4-112">由于 Office 脚本仅适用于 Excel，所以这里只讨论唯一的主机。</span><span class="sxs-lookup"><span data-stu-id="4aba4-112">Since Office Scripts are only available for Excel, that is the only host being discussed here.</span></span>

## <a name="platform-and-ecosystem"></a><span data-ttu-id="4aba4-113">平台和生态系统</span><span class="sxs-lookup"><span data-stu-id="4aba4-113">Platform and ecosystem</span></span>

<span data-ttu-id="4aba4-114">VBA 设计用于桌面，而 Office 脚本是为 web 设计的。</span><span class="sxs-lookup"><span data-stu-id="4aba4-114">VBA is designed for the desktop and Office Scripts are designed for the web.</span></span> <span data-ttu-id="4aba4-115">VBA 可以与用户桌面进行交互，以与类似的技术（如 COM 和 OLE）进行连接。</span><span class="sxs-lookup"><span data-stu-id="4aba4-115">VBA can interact with a user's desktop to connect with similar technologies, such as COM and OLE.</span></span> <span data-ttu-id="4aba4-116">但是，VBA 无法方便地调用到 internet。</span><span class="sxs-lookup"><span data-stu-id="4aba4-116">However, VBA has no convenient way to call out to the internet.</span></span>

<span data-ttu-id="4aba4-117">Office 脚本使用适用于 JavaScript 的通用运行时。</span><span class="sxs-lookup"><span data-stu-id="4aba4-117">Office Scripts use a universal runtime for JavaScript.</span></span> <span data-ttu-id="4aba4-118">这将提供一致的行为和可访问性，而无需考虑用于运行脚本的计算机。</span><span class="sxs-lookup"><span data-stu-id="4aba4-118">This gives consistent behavior and accessibility, regardless of the machine being used to run the script.</span></span> <span data-ttu-id="4aba4-119">它们还可以调用其他 web 服务。</span><span class="sxs-lookup"><span data-stu-id="4aba4-119">They can also make calls to other web services.</span></span>

## <a name="security"></a><span data-ttu-id="4aba4-120">安全性</span><span class="sxs-lookup"><span data-stu-id="4aba4-120">Security</span></span>

<span data-ttu-id="4aba4-121">VBA 宏与 Excel 具有相同的安全净空。</span><span class="sxs-lookup"><span data-stu-id="4aba4-121">VBA macros have the same security clearance as Excel.</span></span> <span data-ttu-id="4aba4-122">这样，他们就可以拥有对桌面的完全访问权限。</span><span class="sxs-lookup"><span data-stu-id="4aba4-122">This gives them full access to your desktop.</span></span> <span data-ttu-id="4aba4-123">Office 脚本仅具有对工作簿的访问权限，而不是承载工作簿的计算机。</span><span class="sxs-lookup"><span data-stu-id="4aba4-123">Office Scripts only have access to the workbook, not the machine hosting the workbook.</span></span> <span data-ttu-id="4aba4-124">此外，不能与脚本共享任何 JavaScript 身份验证令牌。</span><span class="sxs-lookup"><span data-stu-id="4aba4-124">Additionally, no JavaScript authentication tokens can be shared with scripts.</span></span> <span data-ttu-id="4aba4-125">这意味着，该脚本既不具有已登录用户的令牌，也不具有用于登录外部服务的任何 API 功能，因此它们无法使用现有令牌代表用户进行外部呼叫。</span><span class="sxs-lookup"><span data-stu-id="4aba4-125">This means the script has neither the tokens of the signed-in user nor are there any API capabilities for signing in to an external service, so they are unable to use existing tokens to make external calls on behalf of the user.</span></span>

<span data-ttu-id="4aba4-126">管理员具有三个 VBA 宏选项：允许租户上的所有宏、不允许租户上的宏，或仅允许带有签名证书的宏。</span><span class="sxs-lookup"><span data-stu-id="4aba4-126">Admins have three options for VBA macros: allow all macros on the tenant, allow no macros on the tenant, or allow only macros with signed certificates.</span></span> <span data-ttu-id="4aba4-127">这种缺乏的粒度使得难以隔离单个损坏的主角。</span><span class="sxs-lookup"><span data-stu-id="4aba4-127">This lack of granularity makes it hard to isolate a single bad actor.</span></span> <span data-ttu-id="4aba4-128">目前，Office 脚本是针对租户的 "打开" 或 "关闭"。</span><span class="sxs-lookup"><span data-stu-id="4aba4-128">Currently, Office Scripts are either on or off for a tenant.</span></span> <span data-ttu-id="4aba4-129">不过，我们正在努力为管理员提供对各个脚本和脚本编写者的更多控制。</span><span class="sxs-lookup"><span data-stu-id="4aba4-129">However, we are working to give admins more control over individual scripts and script creators.</span></span>

## <a name="coverage"></a><span data-ttu-id="4aba4-130">报道</span><span class="sxs-lookup"><span data-stu-id="4aba4-130">Coverage</span></span>

<span data-ttu-id="4aba4-131">目前，VBA 提供了更全面的 Excel 功能，尤其是在桌面客户端上提供的功能。</span><span class="sxs-lookup"><span data-stu-id="4aba4-131">Currently, VBA offers a more complete coverage of Excel features, particularly those available on the desktop client.</span></span> <span data-ttu-id="4aba4-132">Office 脚本涵盖了 web 上的 Excel 的几乎所有方案。</span><span class="sxs-lookup"><span data-stu-id="4aba4-132">Office Scripts cover nearly all of the scenarios for Excel on the web.</span></span> <span data-ttu-id="4aba4-133">此外，在 web 上 debut 新功能时，Office 脚本将同时为操作记录器和 JavaScript Api 支持这些功能。</span><span class="sxs-lookup"><span data-stu-id="4aba4-133">Additionally, as new features debut on the web, Office Scripts will support them for both the Action Recorder and JavaScript APIs.</span></span>

## <a name="power-automate"></a><span data-ttu-id="4aba4-134">Power Automate</span><span class="sxs-lookup"><span data-stu-id="4aba4-134">Power Automate</span></span>

<span data-ttu-id="4aba4-135">可以通过 "Power 自动化" 运行 Office 脚本。</span><span class="sxs-lookup"><span data-stu-id="4aba4-135">Office Scripts can be run through Power Automate.</span></span> <span data-ttu-id="4aba4-136">您的工作簿可以通过计划或事件驱动的流进行更新，让您无需打开 Excel 即可自动执行工作流。</span><span class="sxs-lookup"><span data-stu-id="4aba4-136">Your workbook can be updated through scheduled or event-driven flows, letting you automate workflows without even opening Excel.</span></span> <span data-ttu-id="4aba4-137">这意味着，只要您的工作簿存储在 OneDrive (中并可访问以实现自动) ，流就可以运行您的脚本，而不管您和您的组织使用的是 Excel 的桌面、Mac 还是 web 客户端。</span><span class="sxs-lookup"><span data-stu-id="4aba4-137">This means that as long as your workbook is stored in OneDrive (and accessible to Power Automate), a flow can run your scripts regardless of whether you and your organization use Excel's desktop, Mac, or web client.</span></span>

<span data-ttu-id="4aba4-138">VBA 没有电源自动连接器。</span><span class="sxs-lookup"><span data-stu-id="4aba4-138">VBA doesn't have a Power Automate connector.</span></span> <span data-ttu-id="4aba4-139">所有受支持的 VBA 方案都涉及用户参与宏的执行。</span><span class="sxs-lookup"><span data-stu-id="4aba4-139">All supported VBA scenarios involved a user attending to the macro's execution.</span></span>

## <a name="see-also"></a><span data-ttu-id="4aba4-140">另请参阅</span><span class="sxs-lookup"><span data-stu-id="4aba4-140">See also</span></span>

- [<span data-ttu-id="4aba4-141">Excel web 版中的 Office 脚本</span><span class="sxs-lookup"><span data-stu-id="4aba4-141">Office Scripts in Excel on the web</span></span>](../overview/excel.md)
- [<span data-ttu-id="4aba4-142">Office 脚本与 Office 加载项之间的差异</span><span class="sxs-lookup"><span data-stu-id="4aba4-142">Differences between Office Scripts and Office Add-ins</span></span>](add-ins-differences.md)
- [<span data-ttu-id="4aba4-143">Office 脚本疑难解答</span><span class="sxs-lookup"><span data-stu-id="4aba4-143">Troubleshooting Office Scripts</span></span>](../testing/troubleshooting.md)
- [<span data-ttu-id="4aba4-144">Excel VBA 参考</span><span class="sxs-lookup"><span data-stu-id="4aba4-144">Excel VBA reference</span></span>](/office/vba/api/overview/excel)
