---
title: Office 脚本疑难解答
description: Office 脚本的调试提示和技术，以及帮助资源。
ms.date: 07/23/2020
localization_priority: Normal
ms.openlocfilehash: 00727b497d49a2d1d3f9c61e259b8d8d75028a59
ms.sourcegitcommit: ff7fde04ce5a66d8df06ed505951c8111e2e9833
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 08/11/2020
ms.locfileid: "46616680"
---
# <a name="troubleshooting-office-scripts"></a><span data-ttu-id="17c79-103">Office 脚本疑难解答</span><span class="sxs-lookup"><span data-stu-id="17c79-103">Troubleshooting Office Scripts</span></span>

<span data-ttu-id="17c79-104">开发 Office 脚本时，可能会产生错误。</span><span class="sxs-lookup"><span data-stu-id="17c79-104">As you develop Office Scripts, you may make mistakes.</span></span> <span data-ttu-id="17c79-105">没关系。</span><span class="sxs-lookup"><span data-stu-id="17c79-105">It's okay.</span></span> <span data-ttu-id="17c79-106">我们有一些工具，可帮助查找问题并使你的脚本完美运行。</span><span class="sxs-lookup"><span data-stu-id="17c79-106">We have tools that help find the problems and get your scripts working perfectly.</span></span>

## <a name="console-logs"></a><span data-ttu-id="17c79-107">控制台日志</span><span class="sxs-lookup"><span data-stu-id="17c79-107">Console logs</span></span>

<span data-ttu-id="17c79-108">有时，在进行故障排除时，您需要将消息打印到屏幕。</span><span class="sxs-lookup"><span data-stu-id="17c79-108">Sometimes while troubleshooting, you'll want to print messages to the screen.</span></span> <span data-ttu-id="17c79-109">这些值可显示变量的当前值或触发的代码路径。</span><span class="sxs-lookup"><span data-stu-id="17c79-109">These can show you the current value of variables or which code paths are being triggered.</span></span> <span data-ttu-id="17c79-110">为此，请将文本记录到控制台。</span><span class="sxs-lookup"><span data-stu-id="17c79-110">To do this, log text to the console.</span></span>

```TypeScript
console.log("Logging myRange's address.");
console.log(myRange.getAddress());
```

<span data-ttu-id="17c79-111">传递给的字符串 `console.log` 将显示在代码编辑器的日志记录控制台中。</span><span class="sxs-lookup"><span data-stu-id="17c79-111">Strings passed to`console.log` will be displayed in the Code Editor's logging console.</span></span> <span data-ttu-id="17c79-112">若要打开控制台，请按**省略号**按钮，然后选择 "**日志 ...** "。</span><span class="sxs-lookup"><span data-stu-id="17c79-112">To turn on the console, press the **Ellipses** button and select **Logs...**</span></span>

<span data-ttu-id="17c79-113">日志不会影响工作簿。</span><span class="sxs-lookup"><span data-stu-id="17c79-113">Logs do not affect the workbook.</span></span>

## <a name="error-messages"></a><span data-ttu-id="17c79-114">错误消息</span><span class="sxs-lookup"><span data-stu-id="17c79-114">Error messages</span></span>

<span data-ttu-id="17c79-115">如果 Excel 脚本在运行时遇到问题，则会产生错误。</span><span class="sxs-lookup"><span data-stu-id="17c79-115">When your Excel Script encounters a problem running, it produces an error.</span></span> <span data-ttu-id="17c79-116">您将看到提示窗口，询问您是否要**查看日志**。</span><span class="sxs-lookup"><span data-stu-id="17c79-116">You'll see a prompt pop-up asking if you want to **View Logs**.</span></span> <span data-ttu-id="17c79-117">按该按钮打开控制台并显示任何错误。</span><span class="sxs-lookup"><span data-stu-id="17c79-117">Press that button to open the console and display any errors.</span></span>

## <a name="automate-tab-not-appearing"></a><span data-ttu-id="17c79-118">"自动" 选项卡未显示</span><span class="sxs-lookup"><span data-stu-id="17c79-118">Automate tab not appearing</span></span>

<span data-ttu-id="17c79-119">以下步骤将帮助解决与 Excel for the web 中未显示的 "**自动**" 选项卡相关的任何问题。</span><span class="sxs-lookup"><span data-stu-id="17c79-119">The following steps should help troubleshoot any problems related to the **Automate** tab not appearing in Excel for the web.</span></span>

1. <span data-ttu-id="17c79-120">[请确保你的 Microsoft 365 许可证包括 Office 脚本](../overview/excel.md#requirements)。</span><span class="sxs-lookup"><span data-stu-id="17c79-120">[Make sure your Microsoft 365 license includes Office Scripts](../overview/excel.md#requirements).</span></span>
1. <span data-ttu-id="17c79-121">[让管理员启用该功能](https://support.office.com/article/office-scripts-settings-in-m365-19d3c51a-6ca2-40ab-978d-60fa49554dcf)。</span><span class="sxs-lookup"><span data-stu-id="17c79-121">[Have your admin enable the feature](https://support.office.com/article/office-scripts-settings-in-m365-19d3c51a-6ca2-40ab-978d-60fa49554dcf).</span></span>
1. <span data-ttu-id="17c79-122">[检查您的浏览器是否受支持](platform-limits.md#browser-support)。</span><span class="sxs-lookup"><span data-stu-id="17c79-122">[Check that your browser is supported](platform-limits.md#browser-support).</span></span>
1. <span data-ttu-id="17c79-123">[确保启用了第三方 cookie](platform-limits.md#third-party-cookies)。</span><span class="sxs-lookup"><span data-stu-id="17c79-123">[Ensure third-party cookies are enabled](platform-limits.md#third-party-cookies).</span></span>

## <a name="help-resources"></a><span data-ttu-id="17c79-124">帮助资源</span><span class="sxs-lookup"><span data-stu-id="17c79-124">Help resources</span></span>

<span data-ttu-id="17c79-125">[堆栈溢出](https://stackoverflow.com/questions/tagged/office-scripts)是一种愿意帮助处理编码问题的开发人员社区。</span><span class="sxs-lookup"><span data-stu-id="17c79-125">[Stack Overflow](https://stackoverflow.com/questions/tagged/office-scripts) is a community of developers willing to help with coding problems.</span></span> <span data-ttu-id="17c79-126">通常情况下，你可以通过快速堆栈溢出搜索找到问题的解决方案。</span><span class="sxs-lookup"><span data-stu-id="17c79-126">Often, you'll be able to find the solution to your problem through a quick Stack Overflow search.</span></span> <span data-ttu-id="17c79-127">如果不是，请询问问题并使用 "office-scripts" 标记对其进行标记。</span><span class="sxs-lookup"><span data-stu-id="17c79-127">If not, ask your question and tag it with the "office-scripts" tag.</span></span> <span data-ttu-id="17c79-128">请务必指出您正在创建 Office*脚本*，而不是 office*外接程序*。</span><span class="sxs-lookup"><span data-stu-id="17c79-128">Be sure to mention you're creating an Office *Script*, not an Office *Add-in*.</span></span>

<span data-ttu-id="17c79-129">如果您遇到 Office JavaScript API 问题，请在[OfficeDev/Office js](https://github.com/OfficeDev/office-js) GitHub 存储库中创建问题。</span><span class="sxs-lookup"><span data-stu-id="17c79-129">If you encounter a problem with the Office JavaScript API, create an issue in the [OfficeDev/office-js](https://github.com/OfficeDev/office-js) GitHub repository.</span></span> <span data-ttu-id="17c79-130">产品团队的成员将响应问题并提供进一步的帮助。</span><span class="sxs-lookup"><span data-stu-id="17c79-130">Members of the product team will respond to issues and provide further assistance.</span></span> <span data-ttu-id="17c79-131">在**OfficeDev/js**存储库中创建问题表示您在 OFFICE JavaScript API 库中发现产品团队应解决的缺陷。</span><span class="sxs-lookup"><span data-stu-id="17c79-131">Creating an issue in the **OfficeDev/office-js** repository indicates you have found a flaw in the Office JavaScript API library that the product team should address.</span></span>

<span data-ttu-id="17c79-132">如果操作记录器或编辑器存在问题，请通过 Excel 中的 "帮助" **> 反馈**按钮发送反馈。</span><span class="sxs-lookup"><span data-stu-id="17c79-132">If there is a problem with the Action Recorder or Editor, send feedback through the **Help > Feedback** button in Excel.</span></span>

## <a name="see-also"></a><span data-ttu-id="17c79-133">另请参阅</span><span class="sxs-lookup"><span data-stu-id="17c79-133">See also</span></span>

- [<span data-ttu-id="17c79-134">Excel web 版中的 Office 脚本</span><span class="sxs-lookup"><span data-stu-id="17c79-134">Office Scripts in Excel on the web</span></span>](../overview/excel.md)
- [<span data-ttu-id="17c79-135">Web 上的 Excel 中 Office 脚本的脚本基础</span><span class="sxs-lookup"><span data-stu-id="17c79-135">Scripting Fundamentals for Office Scripts in Excel on the web</span></span>](../develop/scripting-fundamentals.md)
- [<span data-ttu-id="17c79-136">Office 脚本的平台限制</span><span class="sxs-lookup"><span data-stu-id="17c79-136">Platform Limits with Office Scripts</span></span>](platform-limits.md)
- [<span data-ttu-id="17c79-137">提高 Office 脚本的性能</span><span class="sxs-lookup"><span data-stu-id="17c79-137">Improve the performance of your Office Scripts</span></span>](../develop/web-client-performance.md)
- [<span data-ttu-id="17c79-138">消除 Office 脚本的影响</span><span class="sxs-lookup"><span data-stu-id="17c79-138">Undo the effects of an Office Script</span></span>](undo.md)
