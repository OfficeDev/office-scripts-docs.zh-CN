---
title: Office 脚本疑难解答
description: Office 脚本的调试提示和技术，以及帮助资源。
ms.date: 12/13/2019
localization_priority: Normal
ms.openlocfilehash: 959faff875f342dc1b1ab158ad9ded24732b0894
ms.sourcegitcommit: b075eed5a6f275274fbbf6d62633219eac416f26
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/10/2020
ms.locfileid: "42700114"
---
# <a name="troubleshooting-office-scripts"></a><span data-ttu-id="ca198-103">Office 脚本疑难解答</span><span class="sxs-lookup"><span data-stu-id="ca198-103">Troubleshooting Office Scripts</span></span>

<span data-ttu-id="ca198-104">开发 Office 脚本时，可能会产生错误。</span><span class="sxs-lookup"><span data-stu-id="ca198-104">As you develop Office Scripts, you may make mistakes.</span></span> <span data-ttu-id="ca198-105">没关系。</span><span class="sxs-lookup"><span data-stu-id="ca198-105">It's okay.</span></span> <span data-ttu-id="ca198-106">我们有一些工具，可帮助查找问题并使你的脚本完美运行。</span><span class="sxs-lookup"><span data-stu-id="ca198-106">We have tools that help find the problems and get your scripts working perfectly.</span></span>

## <a name="console-logs"></a><span data-ttu-id="ca198-107">控制台日志</span><span class="sxs-lookup"><span data-stu-id="ca198-107">Console logs</span></span>

<span data-ttu-id="ca198-108">有时，在进行故障排除时，您需要将消息打印到屏幕。</span><span class="sxs-lookup"><span data-stu-id="ca198-108">Sometimes while troubleshooting, you'll want to print messages to the screen.</span></span> <span data-ttu-id="ca198-109">这些值可显示变量的当前值或触发的代码路径。</span><span class="sxs-lookup"><span data-stu-id="ca198-109">These can show you the current value of variables or which code paths are being triggered.</span></span> <span data-ttu-id="ca198-110">为此，请将文本记录到控制台。</span><span class="sxs-lookup"><span data-stu-id="ca198-110">To do this, log text to the console.</span></span>

```TypeScript
console.log("Logging my range's address.");
myRange.load("address");
await context.sync();
console.log(myRange.address);
```

> [!IMPORTANT]
> <span data-ttu-id="ca198-111">在记录对象`load`属性之前， `sync`请不要忘记工作表数据和工作簿。</span><span class="sxs-lookup"><span data-stu-id="ca198-111">Don't forget to `load` worksheet data and `sync` with the workbook before logging object properties.</span></span>

<span data-ttu-id="ca198-112">传递给`console.log`的字符串将显示在代码编辑器的日志记录控制台中。</span><span class="sxs-lookup"><span data-stu-id="ca198-112">Strings passed to`console.log` will be displayed in the Code Editor's logging console.</span></span> <span data-ttu-id="ca198-113">若要打开控制台，请按**省略号**按钮，然后选择 "**日志 ...** "。</span><span class="sxs-lookup"><span data-stu-id="ca198-113">To turn on the console, press the **Ellipses** button and select **Logs...**</span></span>

<span data-ttu-id="ca198-114">日志不会影响工作簿。</span><span class="sxs-lookup"><span data-stu-id="ca198-114">Logs do not affect the workbook.</span></span>

## <a name="error-messages"></a><span data-ttu-id="ca198-115">错误消息</span><span class="sxs-lookup"><span data-stu-id="ca198-115">Error messages</span></span>

<span data-ttu-id="ca198-116">如果 Excel 脚本在运行时遇到问题，则会产生错误。</span><span class="sxs-lookup"><span data-stu-id="ca198-116">When your Excel Script encounters a problem running, it produces an error.</span></span> <span data-ttu-id="ca198-117">您将看到提示窗口，询问您是否要**查看日志**。</span><span class="sxs-lookup"><span data-stu-id="ca198-117">You'll see a prompt pop-up asking if you want to **View Logs**.</span></span> <span data-ttu-id="ca198-118">按该按钮打开控制台并显示任何错误。</span><span class="sxs-lookup"><span data-stu-id="ca198-118">Press that button to open the console and display any errors.</span></span>

## <a name="help-resources"></a><span data-ttu-id="ca198-119">帮助资源</span><span class="sxs-lookup"><span data-stu-id="ca198-119">Help resources</span></span>

<span data-ttu-id="ca198-120">[堆栈溢出](https://stackoverflow.com/questions/tagged/office-scripts)是一种愿意帮助处理编码问题的开发人员社区。</span><span class="sxs-lookup"><span data-stu-id="ca198-120">[Stack Overflow](https://stackoverflow.com/questions/tagged/office-scripts) is a community of developers willing to help with coding problems.</span></span> <span data-ttu-id="ca198-121">通常情况下，你可以通过快速堆栈溢出搜索找到问题的解决方案。</span><span class="sxs-lookup"><span data-stu-id="ca198-121">Often, you'll be able to find the solution to your problem through a quick Stack Overflow search.</span></span> <span data-ttu-id="ca198-122">如果不是，请询问问题并使用 "office-scripts" 标记对其进行标记。</span><span class="sxs-lookup"><span data-stu-id="ca198-122">If not, ask your question and tag it with the "office-scripts" tag.</span></span> <span data-ttu-id="ca198-123">请务必指出您正在创建 Office*脚本*，而不是 office*外接程序*。</span><span class="sxs-lookup"><span data-stu-id="ca198-123">Be sure to mention you're creating an Office *Script*, not an Office *Add-in*.</span></span>

<span data-ttu-id="ca198-124">如果您遇到 Office JavaScript API 问题，请在[OfficeDev/Office js](https://github.com/OfficeDev/office-js) GitHub 存储库中创建问题。</span><span class="sxs-lookup"><span data-stu-id="ca198-124">If you encounter a problem with the Office JavaScript API, create an issue in the [OfficeDev/office-js](https://github.com/OfficeDev/office-js) GitHub repository.</span></span> <span data-ttu-id="ca198-125">产品团队的成员将响应问题并提供进一步的帮助。</span><span class="sxs-lookup"><span data-stu-id="ca198-125">Members of the product team will respond to issues and provide further assistance.</span></span> <span data-ttu-id="ca198-126">在**OfficeDev/js**存储库中创建问题表示您在 OFFICE JavaScript API 库中发现产品团队应解决的缺陷。</span><span class="sxs-lookup"><span data-stu-id="ca198-126">Creating an issue in the **OfficeDev/office-js** repository indicates you have found a flaw in the Office JavaScript API library that the product team should address.</span></span>

<span data-ttu-id="ca198-127">如果操作记录器或编辑器存在问题，请通过 Excel 中的 "帮助" **> 反馈**按钮发送反馈。</span><span class="sxs-lookup"><span data-stu-id="ca198-127">If there is a problem with the Action Recorder or Editor, send feedback through the **Help > Feedback** button in Excel.</span></span>

## <a name="see-also"></a><span data-ttu-id="ca198-128">另请参阅</span><span class="sxs-lookup"><span data-stu-id="ca198-128">See also</span></span>

- [<span data-ttu-id="ca198-129">Web 上的 Excel 中的 Office 脚本</span><span class="sxs-lookup"><span data-stu-id="ca198-129">Office Scripts in Excel on the web</span></span>](../overview/excel.md)
- [<span data-ttu-id="ca198-130">Web 上的 Excel 中 Office 脚本的脚本基础</span><span class="sxs-lookup"><span data-stu-id="ca198-130">Scripting Fundamentals for Office Scripts in Excel on the web</span></span>](../develop/scripting-fundamentals.md)
- [<span data-ttu-id="ca198-131">撤消 Office 脚本的效果</span><span class="sxs-lookup"><span data-stu-id="ca198-131">Undo the effects of an Office Script</span></span>](undo.md)
