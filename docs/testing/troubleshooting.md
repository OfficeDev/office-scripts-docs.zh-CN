---
title: 脚本Office疑难解答
description: 调试脚本的Office以及帮助资源。
ms.date: 11/11/2021
ms.localizationpriority: medium
ms.openlocfilehash: 2a4514aa55550311223cf6fa1179541a37e37f56
ms.sourcegitcommit: 7023b9e23499806901a5ecf8ebc460b76887cca6
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/31/2022
ms.locfileid: "64586043"
---
# <a name="troubleshoot-office-scripts"></a>脚本Office疑难解答

开发脚本Office时，可能会出错。 没关系。 你拥有可帮助查找问题和使脚本正常工作的工具。

> [!NOTE]
> 有关特定于使用 Office 脚本的Power Automate疑难解答Office[脚本疑](power-automate-troubleshooting.md)难解答Power Automate。

## <a name="types-of-errors"></a>错误类型

Office脚本错误分为两类之一：

* 编译时错误或警告
* 运行时错误

### <a name="compile-time-errors"></a>编译时错误

编译时错误和警告最初显示在代码编辑器中。 这些由编辑器中的红色波浪下划线显示。 它们还会显示在"代码 **编辑器"** 任务窗格底部的"问题"选项卡下。 选择该错误将提供有关问题的更多详细信息，并给出解决方案建议。 在运行脚本之前，应解决编译时错误。

:::image type="content" source="../images/explicit-any-editor-message.png" alt-text="代码编辑器悬停文本中显示的编译器错误。":::

你还可能会看到橙色警告下划线和灰色信息性消息。 这些指示性能建议或脚本可能有意外影响的其他可能性。 在消除这些警告之前，应仔细检查这些警告。

### <a name="runtime-errors"></a>运行时错误

运行时错误是由于脚本中的逻辑问题而发生的。 这可能是因为脚本中使用的对象不在工作簿中，表的格式与预期不同，或者脚本的要求与当前工作簿之间稍有差异。 当不存在名为"TestSheet"的工作表时，以下脚本将生成错误。

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  let mySheet = workbook.getWorksheet('TestSheet');

  // This will throw an error if there is no "TestSheet".
  mySheet.getRange("A1");
}
```

### <a name="console-messages"></a>控制台消息

编译时错误和运行时错误在脚本运行时在控制台中显示错误消息。 它们提供遇到问题的行号。 请记住，任何问题的根本原因可能是与控制台中指示的代码行不同的代码行。

下图显示了显式编译器错误的 [控制台 `any`](../develop/typescript-restrictions.md) 输出。 请注意错误 `[5, 16]` 字符串开头的文本。 这表示错误位于第 5 行，从第 16 个字符开始。
:::image type="content" source="../images/explicit-any-error-message.png" alt-text="代码编辑器控制台显示一条明确的&quot;任何&quot;错误消息。":::

下图显示了运行时错误的控制台输出。 在此，脚本尝试添加具有现有工作表名称的工作表。 同样，请注意错误前面的"第 2 行"，以显示要调查的行。
:::image type="content" source="../images/runtime-error-console.png" alt-text="代码编辑器控制台显示&quot;addWorksheet&quot;调用中的错误。":::

## <a name="console-logs"></a>控制台日志

使用 语句将消息打印到 `console.log` 屏幕。 这些日志可以显示变量的当前值或触发的代码路径。 为此，请调用 `console.log` 任意对象作为参数。 通常， `string` 是在控制台中读取的最简单类型。

```TypeScript
console.log("Logging myRange's address.");
console.log(myRange.getAddress());
```

传递给 的 `console.log` 字符串显示在任务窗格底部的代码编辑器的日志记录控制台中。 日志位于"输出 **"选项卡上** ，但写入日志时选项卡会自动获得焦点。

日志不会影响工作簿。

## <a name="automate-tab-not-appearing-or-office-scripts-unavailable"></a>"自动执行"选项卡不显示或Office脚本不可用

以下步骤应有助于解决与"自动"选项卡未显示在"自动"选项卡Excel web 版。

1. [请确保你的Microsoft 365包含Office脚本](../overview/excel.md#requirements)。
1. [检查浏览器是否受支持](platform-limits.md#browser-support)。
1. [确保已启用第三方 Cookie](platform-limits.md#third-party-cookies)。
1. [确保管理员未禁用Office脚本Microsoft 365 管理中心](/microsoft-365/admin/manage/manage-office-scripts-settings)。

[!INCLUDE [Teams support note](../includes/teams-support-note.md)]

## <a name="help-resources"></a>帮助资源

[Stack Overflow](https://stackoverflow.com/questions/tagged/office-scripts) 是开发人员愿意帮助解决编码问题的社区。 通常，你能够通过快速 Stack Overflow 搜索找到问题的解决方案。 如果没有，请提出你的问题，并标记"office-scripts"标记。 请务必提及你正在创建Office *脚本，* 而不是Office *加载项*。

## <a name="see-also"></a>另请参阅

- [Office 脚本中的最佳实践](../develop/best-practices.md)
- [Office 脚本的平台限制](platform-limits.md)
- [提高脚本Office性能](../develop/web-client-performance.md)
- [PowerAutomate Office脚本疑难解答](power-automate-troubleshooting.md)
- [消除 Office 脚本的影响](undo.md)
