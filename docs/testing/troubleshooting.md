---
title: Office 脚本疑难解答
description: Office 脚本的调试提示和技术，以及帮助资源。
ms.date: 05/14/2020
localization_priority: Normal
ms.openlocfilehash: 6448980eec45214a589444229db0fd781b9fea13
ms.sourcegitcommit: aec3c971c6640429f89b6bb99d2c95ea06725599
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 06/25/2020
ms.locfileid: "44878617"
---
# <a name="troubleshooting-office-scripts"></a><span data-ttu-id="4d376-103">Office 脚本疑难解答</span><span class="sxs-lookup"><span data-stu-id="4d376-103">Troubleshooting Office Scripts</span></span>

<span data-ttu-id="4d376-104">开发 Office 脚本时，可能会产生错误。</span><span class="sxs-lookup"><span data-stu-id="4d376-104">As you develop Office Scripts, you may make mistakes.</span></span> <span data-ttu-id="4d376-105">没关系。</span><span class="sxs-lookup"><span data-stu-id="4d376-105">It's okay.</span></span> <span data-ttu-id="4d376-106">我们有一些工具，可帮助查找问题并使你的脚本完美运行。</span><span class="sxs-lookup"><span data-stu-id="4d376-106">We have tools that help find the problems and get your scripts working perfectly.</span></span>

## <a name="console-logs"></a><span data-ttu-id="4d376-107">控制台日志</span><span class="sxs-lookup"><span data-stu-id="4d376-107">Console logs</span></span>

<span data-ttu-id="4d376-108">有时，在进行故障排除时，您需要将消息打印到屏幕。</span><span class="sxs-lookup"><span data-stu-id="4d376-108">Sometimes while troubleshooting, you'll want to print messages to the screen.</span></span> <span data-ttu-id="4d376-109">这些值可显示变量的当前值或触发的代码路径。</span><span class="sxs-lookup"><span data-stu-id="4d376-109">These can show you the current value of variables or which code paths are being triggered.</span></span> <span data-ttu-id="4d376-110">为此，请将文本记录到控制台。</span><span class="sxs-lookup"><span data-stu-id="4d376-110">To do this, log text to the console.</span></span>

```TypeScript
console.log("Logging myRange's address.");
console.log(myRange.getAddress());
```

<span data-ttu-id="4d376-111">传递给的字符串 `console.log` 将显示在代码编辑器的日志记录控制台中。</span><span class="sxs-lookup"><span data-stu-id="4d376-111">Strings passed to`console.log` will be displayed in the Code Editor's logging console.</span></span> <span data-ttu-id="4d376-112">若要打开控制台，请按**省略号**按钮，然后选择 "**日志 ...** "。</span><span class="sxs-lookup"><span data-stu-id="4d376-112">To turn on the console, press the **Ellipses** button and select **Logs...**</span></span>

<span data-ttu-id="4d376-113">日志不会影响工作簿。</span><span class="sxs-lookup"><span data-stu-id="4d376-113">Logs do not affect the workbook.</span></span>

## <a name="error-messages"></a><span data-ttu-id="4d376-114">错误消息</span><span class="sxs-lookup"><span data-stu-id="4d376-114">Error messages</span></span>

<span data-ttu-id="4d376-115">如果 Excel 脚本在运行时遇到问题，则会产生错误。</span><span class="sxs-lookup"><span data-stu-id="4d376-115">When your Excel Script encounters a problem running, it produces an error.</span></span> <span data-ttu-id="4d376-116">您将看到提示窗口，询问您是否要**查看日志**。</span><span class="sxs-lookup"><span data-stu-id="4d376-116">You'll see a prompt pop-up asking if you want to **View Logs**.</span></span> <span data-ttu-id="4d376-117">按该按钮打开控制台并显示任何错误。</span><span class="sxs-lookup"><span data-stu-id="4d376-117">Press that button to open the console and display any errors.</span></span>

## <a name="help-resources"></a><span data-ttu-id="4d376-118">帮助资源</span><span class="sxs-lookup"><span data-stu-id="4d376-118">Help resources</span></span>

<span data-ttu-id="4d376-119">[堆栈溢出](https://stackoverflow.com/questions/tagged/office-scripts)是一种愿意帮助处理编码问题的开发人员社区。</span><span class="sxs-lookup"><span data-stu-id="4d376-119">[Stack Overflow](https://stackoverflow.com/questions/tagged/office-scripts) is a community of developers willing to help with coding problems.</span></span> <span data-ttu-id="4d376-120">通常情况下，你可以通过快速堆栈溢出搜索找到问题的解决方案。</span><span class="sxs-lookup"><span data-stu-id="4d376-120">Often, you'll be able to find the solution to your problem through a quick Stack Overflow search.</span></span> <span data-ttu-id="4d376-121">如果不是，请询问问题并使用 "office-scripts" 标记对其进行标记。</span><span class="sxs-lookup"><span data-stu-id="4d376-121">If not, ask your question and tag it with the "office-scripts" tag.</span></span> <span data-ttu-id="4d376-122">请务必指出您正在创建 Office*脚本*，而不是 office*外接程序*。</span><span class="sxs-lookup"><span data-stu-id="4d376-122">Be sure to mention you're creating an Office *Script*, not an Office *Add-in*.</span></span>

<span data-ttu-id="4d376-123">如果您遇到 Office JavaScript API 问题，请在[OfficeDev/Office js](https://github.com/OfficeDev/office-js) GitHub 存储库中创建问题。</span><span class="sxs-lookup"><span data-stu-id="4d376-123">If you encounter a problem with the Office JavaScript API, create an issue in the [OfficeDev/office-js](https://github.com/OfficeDev/office-js) GitHub repository.</span></span> <span data-ttu-id="4d376-124">产品团队的成员将响应问题并提供进一步的帮助。</span><span class="sxs-lookup"><span data-stu-id="4d376-124">Members of the product team will respond to issues and provide further assistance.</span></span> <span data-ttu-id="4d376-125">在**OfficeDev/js**存储库中创建问题表示您在 OFFICE JavaScript API 库中发现产品团队应解决的缺陷。</span><span class="sxs-lookup"><span data-stu-id="4d376-125">Creating an issue in the **OfficeDev/office-js** repository indicates you have found a flaw in the Office JavaScript API library that the product team should address.</span></span>

<span data-ttu-id="4d376-126">如果操作记录器或编辑器存在问题，请通过 Excel 中的 "帮助" **> 反馈**按钮发送反馈。</span><span class="sxs-lookup"><span data-stu-id="4d376-126">If there is a problem with the Action Recorder or Editor, send feedback through the **Help > Feedback** button in Excel.</span></span>

## <a name="see-also"></a><span data-ttu-id="4d376-127">另请参阅</span><span class="sxs-lookup"><span data-stu-id="4d376-127">See also</span></span>

- [<span data-ttu-id="4d376-128">Excel web 版中的 Office 脚本</span><span class="sxs-lookup"><span data-stu-id="4d376-128">Office Scripts in Excel on the web</span></span>](../overview/excel.md)
- [<span data-ttu-id="4d376-129">Web 上的 Excel 中 Office 脚本的脚本基础</span><span class="sxs-lookup"><span data-stu-id="4d376-129">Scripting Fundamentals for Office Scripts in Excel on the web</span></span>](../develop/scripting-fundamentals.md)
- [<span data-ttu-id="4d376-130">消除 Office 脚本的影响</span><span class="sxs-lookup"><span data-stu-id="4d376-130">Undo the effects of an Office Script</span></span>](undo.md)
- [<span data-ttu-id="4d376-131">提高 Office 脚本的性能</span><span class="sxs-lookup"><span data-stu-id="4d376-131">Improve the performance of your Office Scripts</span></span>](../develop/web-client-performance.md)
