---
title: 脚本Office疑难解答
description: 调试脚本的Office以及帮助资源。
ms.date: 05/17/2021
localization_priority: Normal
ms.openlocfilehash: 251ad72588422a86c52c81666164c2c4bd79bdb5
ms.sourcegitcommit: 4693c8f79428ec74695328275703af0ba1bfea8f
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 06/23/2021
ms.locfileid: "53074646"
---
# <a name="troubleshoot-office-scripts"></a><span data-ttu-id="3591b-103">脚本Office疑难解答</span><span class="sxs-lookup"><span data-stu-id="3591b-103">Troubleshoot Office Scripts</span></span>

<span data-ttu-id="3591b-104">开发脚本Office时，可能会出错。</span><span class="sxs-lookup"><span data-stu-id="3591b-104">As you develop Office Scripts, you may make mistakes.</span></span> <span data-ttu-id="3591b-105">没关系。</span><span class="sxs-lookup"><span data-stu-id="3591b-105">It's okay.</span></span> <span data-ttu-id="3591b-106">你拥有可帮助查找问题和使脚本正常工作的工具。</span><span class="sxs-lookup"><span data-stu-id="3591b-106">You have the tools to help find the problems and get your scripts working perfectly.</span></span>

## <a name="types-of-errors"></a><span data-ttu-id="3591b-107">错误类型</span><span class="sxs-lookup"><span data-stu-id="3591b-107">Types of errors</span></span>

<span data-ttu-id="3591b-108">Office脚本错误分为两类之一：</span><span class="sxs-lookup"><span data-stu-id="3591b-108">Office Scripts errors fall into one of two categories:</span></span>

* <span data-ttu-id="3591b-109">编译时错误或警告</span><span class="sxs-lookup"><span data-stu-id="3591b-109">Compile-time errors or warnings</span></span>
* <span data-ttu-id="3591b-110">运行时错误</span><span class="sxs-lookup"><span data-stu-id="3591b-110">Runtime errors</span></span>

### <a name="compile-time-errors"></a><span data-ttu-id="3591b-111">编译时错误</span><span class="sxs-lookup"><span data-stu-id="3591b-111">Compile-time errors</span></span>

<span data-ttu-id="3591b-112">编译时错误和警告最初显示在代码编辑器中。</span><span class="sxs-lookup"><span data-stu-id="3591b-112">Compile-time errors and warnings are initially shown in the Code Editor.</span></span> <span data-ttu-id="3591b-113">这些由编辑器中的红色波浪下划线显示。</span><span class="sxs-lookup"><span data-stu-id="3591b-113">These are shown by the wavy red underlines in the editor.</span></span> <span data-ttu-id="3591b-114">它们还会显示在"代码 **编辑器"** 任务窗格底部的"问题"选项卡下。</span><span class="sxs-lookup"><span data-stu-id="3591b-114">They are also displayed under the **Problems** tab at the bottom of the Code Editor task pane.</span></span> <span data-ttu-id="3591b-115">选择该错误将提供有关问题的更多详细信息，并给出解决方案建议。</span><span class="sxs-lookup"><span data-stu-id="3591b-115">Selecting the error will give more details about the problem and suggest solutions.</span></span> <span data-ttu-id="3591b-116">在运行脚本之前，应解决编译时错误。</span><span class="sxs-lookup"><span data-stu-id="3591b-116">Compile-time errors should be addressed before running the script.</span></span>

:::image type="content" source="../images/explicit-any-editor-message.png" alt-text="代码编辑器悬停文本中显示的编译器错误。":::

<span data-ttu-id="3591b-118">你还可能会看到橙色警告下划线和灰色信息性消息。</span><span class="sxs-lookup"><span data-stu-id="3591b-118">You may also see orange warning underlines and grey informational messages.</span></span> <span data-ttu-id="3591b-119">这些指示性能建议或脚本可能有意外影响的其他可能性。</span><span class="sxs-lookup"><span data-stu-id="3591b-119">These indicate performance suggestions or other possibilities where the script may have unintentional effects.</span></span> <span data-ttu-id="3591b-120">在消除这些警告之前，应仔细检查这些警告。</span><span class="sxs-lookup"><span data-stu-id="3591b-120">Such warnings should be examined closely before dismissing them.</span></span>

### <a name="runtime-errors"></a><span data-ttu-id="3591b-121">运行时错误</span><span class="sxs-lookup"><span data-stu-id="3591b-121">Runtime errors</span></span>

<span data-ttu-id="3591b-122">运行时错误是由于脚本中的逻辑问题而发生的。</span><span class="sxs-lookup"><span data-stu-id="3591b-122">Runtime errors happen because of logic issues in the script.</span></span> <span data-ttu-id="3591b-123">这可能是因为脚本中使用的对象不在工作簿中，表的格式与预期不同，或者脚本的要求与当前工作簿之间稍有差异。</span><span class="sxs-lookup"><span data-stu-id="3591b-123">This could be because an object used in the script isn't in the workbook, a table is formatted differently than anticipated, or some other slight discrepancy between the script's requirements and the current workbook.</span></span> <span data-ttu-id="3591b-124">当不存在名为"TestSheet"的工作表时，以下脚本将生成错误。</span><span class="sxs-lookup"><span data-stu-id="3591b-124">The following script generates an error when a worksheet named "TestSheet" is not present.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  let mySheet = workbook.getWorksheet('TestSheet');

  // This will throw an error if there is no "TestSheet".
  mySheet.getRange("A1");
}
```

### <a name="console-messages"></a><span data-ttu-id="3591b-125">控制台消息</span><span class="sxs-lookup"><span data-stu-id="3591b-125">Console messages</span></span>

<span data-ttu-id="3591b-126">编译时错误和运行时错误在脚本运行时在控制台中显示错误消息。</span><span class="sxs-lookup"><span data-stu-id="3591b-126">Both compile-time and runtime errors display error messages in the console when a script runs.</span></span> <span data-ttu-id="3591b-127">它们提供遇到问题的行号。</span><span class="sxs-lookup"><span data-stu-id="3591b-127">They give a line number where the problem was encountered.</span></span> <span data-ttu-id="3591b-128">请记住，任何问题的根本原因可能是与控制台中指示的代码行不同的代码行。</span><span class="sxs-lookup"><span data-stu-id="3591b-128">Keep in mind that the root cause of any issue may be a different line of code than what is indicated in the console.</span></span>

<span data-ttu-id="3591b-129">下图显示了显式编译器错误的[控制台 `any` ](../develop/typescript-restrictions.md)输出。</span><span class="sxs-lookup"><span data-stu-id="3591b-129">The following image shows the console output for the [explicit `any`](../develop/typescript-restrictions.md) compiler error.</span></span> <span data-ttu-id="3591b-130">请注意 `[5, 16]` 错误字符串开头的文本。</span><span class="sxs-lookup"><span data-stu-id="3591b-130">Note the text `[5, 16]` at the beginning of the error string.</span></span> <span data-ttu-id="3591b-131">这表示错误位于第 5 行，从第 16 个字符开始。</span><span class="sxs-lookup"><span data-stu-id="3591b-131">This indicates the error is on line 5, starting at character 16.</span></span>
:::image type="content" source="../images/explicit-any-error-message.png" alt-text="代码编辑器控制台显示一条明确的&quot;任何&quot;错误消息。":::

<span data-ttu-id="3591b-133">下图显示了运行时错误的控制台输出。</span><span class="sxs-lookup"><span data-stu-id="3591b-133">The follow image shows the console output for a runtime error.</span></span> <span data-ttu-id="3591b-134">在此，脚本尝试添加具有现有工作表名称的工作表。</span><span class="sxs-lookup"><span data-stu-id="3591b-134">Here, the script tries to add a worksheet with a the name of an existing worksheet.</span></span> <span data-ttu-id="3591b-135">同样，请注意错误前面的"第 2 行"，以显示要调查的行。</span><span class="sxs-lookup"><span data-stu-id="3591b-135">Again, note the "Line 2" preceding the error to show which line to investigate.</span></span>
:::image type="content" source="../images/runtime-error-console.png" alt-text="代码编辑器控制台显示&quot;addWorksheet&quot;调用中的错误。":::

## <a name="console-logs"></a><span data-ttu-id="3591b-137">控制台日志</span><span class="sxs-lookup"><span data-stu-id="3591b-137">Console logs</span></span>

<span data-ttu-id="3591b-138">使用 语句将消息打印到 `console.log` 屏幕。</span><span class="sxs-lookup"><span data-stu-id="3591b-138">Print messages to the screen with the `console.log` statement.</span></span> <span data-ttu-id="3591b-139">这些日志可以显示变量的当前值或触发的代码路径。</span><span class="sxs-lookup"><span data-stu-id="3591b-139">These logs can show you the current value of variables or which code paths are being triggered.</span></span> <span data-ttu-id="3591b-140">为此，请 `console.log` 调用任意对象作为参数。</span><span class="sxs-lookup"><span data-stu-id="3591b-140">To do this, call `console.log` with any object as a parameter.</span></span> <span data-ttu-id="3591b-141">通常， `string` 是在控制台中读取的最简单类型。</span><span class="sxs-lookup"><span data-stu-id="3591b-141">Usually, a `string` is the easiest type to read in the console.</span></span>

```TypeScript
console.log("Logging myRange's address.");
console.log(myRange.getAddress());
```

<span data-ttu-id="3591b-142">传递给 的字符串显示在任务窗格底部的代码编辑器的日志记录 `console.log` 控制台中。</span><span class="sxs-lookup"><span data-stu-id="3591b-142">Strings passed to `console.log` are displayed in the Code Editor's logging console, at the bottom of the task pane.</span></span> <span data-ttu-id="3591b-143">日志位于"输出" **选项卡上** ，但写入日志时选项卡会自动获得焦点。</span><span class="sxs-lookup"><span data-stu-id="3591b-143">Logs are found on the **Output** tab, though the tab automatically gains focus when a log is written.</span></span>

<span data-ttu-id="3591b-144">日志不会影响工作簿。</span><span class="sxs-lookup"><span data-stu-id="3591b-144">Logs do not affect the workbook.</span></span>

## <a name="automate-tab-not-appearing-or-office-scripts-unavailable"></a><span data-ttu-id="3591b-145">"自动化"选项卡不显示或Office脚本不可用</span><span class="sxs-lookup"><span data-stu-id="3591b-145">Automate tab not appearing or Office Scripts unavailable</span></span>

<span data-ttu-id="3591b-146">以下步骤应有助于解决与"自动"选项卡未显示在"自动"选项卡Excel web 版。</span><span class="sxs-lookup"><span data-stu-id="3591b-146">The following steps should help troubleshoot any problems related to the **Automate** tab not appearing in Excel on the web.</span></span>

1. <span data-ttu-id="3591b-147">[请确保你的Microsoft 365包括Office脚本](../overview/excel.md#requirements)。</span><span class="sxs-lookup"><span data-stu-id="3591b-147">[Make sure your Microsoft 365 license includes Office Scripts](../overview/excel.md#requirements).</span></span>
1. <span data-ttu-id="3591b-148">[检查浏览器是否受支持](platform-limits.md#browser-support)。</span><span class="sxs-lookup"><span data-stu-id="3591b-148">[Check that your browser is supported](platform-limits.md#browser-support).</span></span>
1. <span data-ttu-id="3591b-149">[确保已启用第三方 Cookie。](platform-limits.md#third-party-cookies)</span><span class="sxs-lookup"><span data-stu-id="3591b-149">[Ensure third-party cookies are enabled](platform-limits.md#third-party-cookies).</span></span>
1. <span data-ttu-id="3591b-150">[确保管理员未禁用脚本Office脚本Microsoft 365 管理中心。](/microsoft-365/admin/manage/manage-office-scripts-settings)</span><span class="sxs-lookup"><span data-stu-id="3591b-150">[Ensure that your admin has not disabled Office Scripts in the Microsoft 365 admin center](/microsoft-365/admin/manage/manage-office-scripts-settings).</span></span>

[!INCLUDE [Teams support note](../includes/teams-support-note.md)]

## <a name="troubleshoot-scripts-in-power-automate"></a><span data-ttu-id="3591b-151">疑难解答脚本Power Automate</span><span class="sxs-lookup"><span data-stu-id="3591b-151">Troubleshoot scripts in Power Automate</span></span>

<span data-ttu-id="3591b-152">有关通过脚本运行脚本Power Automate的信息，请参阅 Troubleshoot [Office Scripts running in Power Automate](power-automate-troubleshooting.md)。</span><span class="sxs-lookup"><span data-stu-id="3591b-152">For information specific to running scripts through Power Automate, see [Troubleshoot Office Scripts running in Power Automate](power-automate-troubleshooting.md).</span></span>

## <a name="help-resources"></a><span data-ttu-id="3591b-153">帮助资源</span><span class="sxs-lookup"><span data-stu-id="3591b-153">Help resources</span></span>

<span data-ttu-id="3591b-154">[Stack Overflow](https://stackoverflow.com/questions/tagged/office-scripts) 是开发人员愿意帮助解决编码问题的社区。</span><span class="sxs-lookup"><span data-stu-id="3591b-154">[Stack Overflow](https://stackoverflow.com/questions/tagged/office-scripts) is a community of developers willing to help with coding problems.</span></span> <span data-ttu-id="3591b-155">通常，你能够通过快速 Stack Overflow 搜索找到问题的解决方案。</span><span class="sxs-lookup"><span data-stu-id="3591b-155">Often, you'll be able to find the solution to your problem through a quick Stack Overflow search.</span></span> <span data-ttu-id="3591b-156">如果没有，请提出你的问题，并标记"office-scripts"标记。</span><span class="sxs-lookup"><span data-stu-id="3591b-156">If not, ask your question and tag it with the "office-scripts" tag.</span></span> <span data-ttu-id="3591b-157">请务必提及你正在创建一个Office *脚本*，而不是Office *加载项。*</span><span class="sxs-lookup"><span data-stu-id="3591b-157">Be sure to mention you're creating an Office *Script*, not an Office *Add-in*.</span></span>

<span data-ttu-id="3591b-158">若要提交对 Office 脚本的功能请求，将你的想法张贴到"用户[](https://excel.uservoice.com/forums/274580-excel-for-the-web?category_id=143439)语音"页面，或者如果功能请求已存在，请为它添加投票。</span><span class="sxs-lookup"><span data-stu-id="3591b-158">To submit a feature request for Office Scripts, post your idea to our [User Voice page](https://excel.uservoice.com/forums/274580-excel-for-the-web?category_id=143439), or if the feature request already exists there, add your vote for it.</span></span> <span data-ttu-id="3591b-159">请确保将请求提交到Excel 网页版、脚本和外接程序"类别中的"文件"下。</span><span class="sxs-lookup"><span data-stu-id="3591b-159">Be sure to file the request under Excel for the web in the "Macros, Scripts and Add-ins" category.</span></span>

<span data-ttu-id="3591b-160">如果操作录制器或编辑器出现问题，请告诉我们。</span><span class="sxs-lookup"><span data-stu-id="3591b-160">If there is a problem with the Action Recorder or Editor, please let us know.</span></span> <span data-ttu-id="3591b-161">在"代码编辑器"任务窗格的 **"..."** 菜单中，选择" **发送反馈"** 按钮以共享任何问题。</span><span class="sxs-lookup"><span data-stu-id="3591b-161">In the Code Editor task pane's **...** menu, select the **Send feedback** button to share any issues.</span></span>

:::image type="content" source="../images/code-editor-feedback.png" alt-text="具有&quot;发送反馈&quot;按钮的代码编辑器溢出菜单。":::

## <a name="see-also"></a><span data-ttu-id="3591b-163">另请参阅</span><span class="sxs-lookup"><span data-stu-id="3591b-163">See also</span></span>

- [<span data-ttu-id="3591b-164">Office 脚本中的最佳实践</span><span class="sxs-lookup"><span data-stu-id="3591b-164">Best practices in Office Scripts</span></span>](../develop/best-practices.md)
- [<span data-ttu-id="3591b-165">Office 脚本的平台限制</span><span class="sxs-lookup"><span data-stu-id="3591b-165">Platform limits with Office Scripts</span></span>](platform-limits.md)
- [<span data-ttu-id="3591b-166">提高脚本Office性能</span><span class="sxs-lookup"><span data-stu-id="3591b-166">Improve the performance of your Office Scripts</span></span>](../develop/web-client-performance.md)
- [<span data-ttu-id="3591b-167">PowerAutomate Office中运行的脚本疑难解答</span><span class="sxs-lookup"><span data-stu-id="3591b-167">Troubleshoot Office Scripts running in PowerAutomate</span></span>](power-automate-troubleshooting.md)
- [<span data-ttu-id="3591b-168">消除 Office 脚本的影响</span><span class="sxs-lookup"><span data-stu-id="3591b-168">Undo the effects of Office Scripts</span></span>](undo.md)
