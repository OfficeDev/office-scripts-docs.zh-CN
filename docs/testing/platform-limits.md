---
title: Office 脚本的平台限制和要求
description: 在 web 上与 Excel 一起使用时，Office 脚本的资源限制和浏览器支持
ms.date: 07/23/2020
localization_priority: Normal
ms.openlocfilehash: 6e297cba0b9f984f2d541cc3c441a666f9ebfcef
ms.sourcegitcommit: ff7fde04ce5a66d8df06ed505951c8111e2e9833
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 08/11/2020
ms.locfileid: "46618155"
---
# <a name="platform-limits-and-requirements-with-office-scripts"></a><span data-ttu-id="caf45-103">Office 脚本的平台限制和要求</span><span class="sxs-lookup"><span data-stu-id="caf45-103">Platform limits and requirements with Office Scripts</span></span>

<span data-ttu-id="caf45-104">开发 Office 脚本时，应注意一些平台限制。</span><span class="sxs-lookup"><span data-stu-id="caf45-104">There are some platform limitations of which you should be aware when developing Office Scripts.</span></span> <span data-ttu-id="caf45-105">本文详细介绍了 web 上的适用于 Excel 的 Office 脚本的浏览器支持和数据限制。</span><span class="sxs-lookup"><span data-stu-id="caf45-105">This article details the browser support and data limits for Office Scripts for Excel on the web.</span></span>

## <a name="browser-support"></a><span data-ttu-id="caf45-106">浏览器支持</span><span class="sxs-lookup"><span data-stu-id="caf45-106">Browser support</span></span>

<span data-ttu-id="caf45-107">Office 脚本在任何[支持 Web Office 的](https://support.microsoft.com/office/ad1303e0-a318-47aa-b409-d3a5eb44e452)浏览器中工作。</span><span class="sxs-lookup"><span data-stu-id="caf45-107">Office Scripts work in any browser that [supports Office for the web](https://support.microsoft.com/office/ad1303e0-a318-47aa-b409-d3a5eb44e452).</span></span> <span data-ttu-id="caf45-108">但是，Internet Explorer 11 (IE 11) 中不支持某些 JavaScript 功能。</span><span class="sxs-lookup"><span data-stu-id="caf45-108">However, some JavaScript features aren't supported in Internet Explorer 11 (IE 11).</span></span> <span data-ttu-id="caf45-109">[ES6 或更高版本](https://www.w3schools.com/Js/js_es6.asp)中引入的任何功能将不适用于 IE 11。</span><span class="sxs-lookup"><span data-stu-id="caf45-109">Any features introduced in [ES6 or later](https://www.w3schools.com/Js/js_es6.asp) won't work with IE 11.</span></span> <span data-ttu-id="caf45-110">如果组织中的人员仍在使用该浏览器，请务必在共享这些脚本时在该环境中对其进行测试。</span><span class="sxs-lookup"><span data-stu-id="caf45-110">If people in your organization still use that browser, be sure to test your scripts in that environment when sharing them.</span></span>

### <a name="third-party-cookies"></a><span data-ttu-id="caf45-111">第三方 cookie</span><span class="sxs-lookup"><span data-stu-id="caf45-111">Third-party cookies</span></span>

<span data-ttu-id="caf45-112">你的浏览器需要启用了第三方 cookie，才能在 Excel 网页上显示 "**自动**" 选项卡。</span><span class="sxs-lookup"><span data-stu-id="caf45-112">Your browser needs third-party cookies enabled to show the **Automate** tab in Excel on the web.</span></span> <span data-ttu-id="caf45-113">如果不显示该选项卡，请检查您的浏览器设置。</span><span class="sxs-lookup"><span data-stu-id="caf45-113">Check your browser settings if the tab isn't being displayed.</span></span> <span data-ttu-id="caf45-114">如果使用的是专用浏览器会话，则每次可能需要重新启用此设置。</span><span class="sxs-lookup"><span data-stu-id="caf45-114">If you're using a private browser session, you may need to re-enable this setting each time.</span></span>

> [!NOTE]
> <span data-ttu-id="caf45-115">某些浏览器将此设置称为 "所有 cookie"，而不是 "第三方 cookie"。</span><span class="sxs-lookup"><span data-stu-id="caf45-115">Some browsers refer to this setting as "all cookies", instead of "third-party cookies".</span></span>

## <a name="data-limits"></a><span data-ttu-id="caf45-116">数据限制</span><span class="sxs-lookup"><span data-stu-id="caf45-116">Data limits</span></span>

<span data-ttu-id="caf45-117">对可以一次传输多少个 Excel 数据以及可以执行多少个单独的电源自动化事务的操作有限制。</span><span class="sxs-lookup"><span data-stu-id="caf45-117">There are limits on how much Excel data can be transferred at once and how many individual Power Automate transactions can be conducted.</span></span>

### <a name="excel"></a><span data-ttu-id="caf45-118">Excel</span><span class="sxs-lookup"><span data-stu-id="caf45-118">Excel</span></span>

<span data-ttu-id="caf45-119">在通过脚本调用工作簿时，网站的 Excel 具有以下限制：</span><span class="sxs-lookup"><span data-stu-id="caf45-119">Excel for the web has the following limitations when making calls to the workbook through a script:</span></span>

- <span data-ttu-id="caf45-120">请求和响应限制为**5mb**。</span><span class="sxs-lookup"><span data-stu-id="caf45-120">Requests and responses are limited to **5MB**.</span></span>
- <span data-ttu-id="caf45-121">范围限制为5000000个**单元格**。</span><span class="sxs-lookup"><span data-stu-id="caf45-121">A range is limited to **five million cells**.</span></span>

<span data-ttu-id="caf45-122">如果在处理大型数据集时遇到错误，请尝试使用多个较小的范围，而不是更大的范围。</span><span class="sxs-lookup"><span data-stu-id="caf45-122">If you're encountering errors when dealing with large datasets, try using multiple smaller ranges instead of larger ranges.</span></span> <span data-ttu-id="caf45-123">您还可以将[getSpecialCells](/javascript/api/office-scripts/excelscript/excelscript.range#getspecialcells-celltype--cellvaluetype-)作为目标单元格（而不是大型区域）的 api。</span><span class="sxs-lookup"><span data-stu-id="caf45-123">You can also APIs like [Range.getSpecialCells](/javascript/api/office-scripts/excelscript/excelscript.range#getspecialcells-celltype--cellvaluetype-) to target specific cells instead of large ranges.</span></span>

### <a name="power-automate"></a><span data-ttu-id="caf45-124">Power Automate</span><span class="sxs-lookup"><span data-stu-id="caf45-124">Power Automate</span></span>

<span data-ttu-id="caf45-125">在使用带电自动化的 Office 脚本时，**每日限制为200个呼叫**。</span><span class="sxs-lookup"><span data-stu-id="caf45-125">When using Office Scripts with Power Automate, you're limited to **200 calls per day**.</span></span> <span data-ttu-id="caf45-126">此限制在 UTC 时间重置为 12:00 AM。</span><span class="sxs-lookup"><span data-stu-id="caf45-126">This limit resets at 12:00 AM UTC.</span></span>

<span data-ttu-id="caf45-127">电源自动化平台还有使用限制，可在[电源自动化的文章限制和配置](/power-automate/limits-and-config)中找到。</span><span class="sxs-lookup"><span data-stu-id="caf45-127">The Power Automate platform also has usage limitations, which can be found in the article [Limits and configuration in Power Automate](/power-automate/limits-and-config).</span></span>

## <a name="see-also"></a><span data-ttu-id="caf45-128">另请参阅</span><span class="sxs-lookup"><span data-stu-id="caf45-128">See also</span></span>

- [<span data-ttu-id="caf45-129">Office 脚本疑难解答</span><span class="sxs-lookup"><span data-stu-id="caf45-129">Troubleshooting Office Scripts</span></span>](troubleshooting.md)
- [<span data-ttu-id="caf45-130">消除 Office 脚本的影响</span><span class="sxs-lookup"><span data-stu-id="caf45-130">Undo the effects of an Office Script</span></span>](undo.md)
- [<span data-ttu-id="caf45-131">提高 Office 脚本的性能</span><span class="sxs-lookup"><span data-stu-id="caf45-131">Improve the performance of your Office Scripts</span></span>](../develop/web-client-performance.md)
- [<span data-ttu-id="caf45-132">Web 上的 Excel 中 Office 脚本的脚本基础</span><span class="sxs-lookup"><span data-stu-id="caf45-132">Scripting Fundamentals for Office Scripts in Excel on the web</span></span>](../develop/scripting-fundamentals.md)
