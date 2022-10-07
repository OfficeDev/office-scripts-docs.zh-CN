---
title: Office 脚本疑难解答
description: Office 脚本的调试提示和技术，以及帮助资源。
ms.date: 10/05/2022
ms.localizationpriority: medium
ms.openlocfilehash: 4fe4a9b17d51d078403d1a46abed774d38eeaa80
ms.sourcegitcommit: 64d506257bee282fb01aedbf4d090781b06e4900
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 10/07/2022
ms.locfileid: "68495465"
---
# <a name="troubleshoot-office-scripts"></a>Office 脚本疑难解答

开发 Office 脚本时，可能会出错。 没关系。 你可以使用工具帮助查找问题并使脚本完美工作。

> [!NOTE]
> 有关特定于 Power Automate 的 Office 脚本的故障排除建议，请参阅 [Power Automate 中运行的 Office 脚本故障排除](power-automate-troubleshooting.md)。

## <a name="types-of-errors"></a>错误类型

Office 脚本错误分为两个类别之一：

* 编译时错误或警告
* 运行时错误

### <a name="compile-time-errors"></a>编译时错误

编译时错误和警告最初显示在代码编辑器中。 编辑器中的波浪红色下划线显示了这些下划线。 它们还会显示在“代码编辑器”任务窗格底部的“ **问题** ”选项卡下。 选择错误将提供有关问题的更多详细信息，并建议解决方案。 在运行脚本之前，应解决编译时错误。

:::image type="content" source="../images/explicit-any-editor-message.png" alt-text="代码编辑器的悬停文本中显示的编译器错误。":::

还可能会看到橙色警告下划线和灰色信息性消息。 这些指示性能建议或其他可能的脚本可能具有无意的效果。 在消除这些警告之前，应仔细检查这些警告。

### <a name="runtime-errors"></a>运行时错误

由于脚本中的逻辑问题，会发生运行时错误。 这可能是因为脚本中使用的对象不在工作簿中，表的格式与预期不同，或者脚本的要求与当前工作簿之间存在一些其他细微差异。 当不存在名为“TestSheet”的工作表时，以下脚本会生成错误。

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  let mySheet = workbook.getWorksheet('TestSheet');

  // This will throw an error if there is no "TestSheet".
  mySheet.getRange("A1");
}
```

### <a name="console-messages"></a>控制台消息

在运行脚本时，编译时和运行时错误都会在控制台中显示错误消息。 它们提供遇到问题的行号。 请记住，任何问题的根本原因可能与控制台中指示的代码行不同。

下图显示了 [显式 `any`](../develop/typescript-restrictions.md) 编译器错误的控制台输出。 记下错误字符串开头的文本 `[5, 16]` 。 这表示错误位于第 5 行，从字符 16 开始。
:::image type="content" source="../images/explicit-any-error-message.png" alt-text="显示显式“any”错误消息的代码编辑器控制台。":::

下图显示了运行时错误的控制台输出。 此处，脚本尝试添加具有现有工作表名称的工作表。 同样，请注意错误前面的“行 2”，以显示要调查的行。
:::image type="content" source="../images/runtime-error-console.png" alt-text="显示“addWorksheet”调用错误的代码编辑器控制台。":::

## <a name="console-logs"></a>控制台日志

使用 `console.log` 语句将消息打印到屏幕。 这些日志可以显示变量的当前值或触发的代码路径。 若要执行此操作，请使用任何对象作为参数调用 `console.log` 。 通常，a `string` 是最容易在控制台中读取的类型。

```TypeScript
console.log("Logging myRange's address.");
console.log(myRange.getAddress());
```

传递给 `console.log` 的字符串显示在代码编辑器的日志记录控制台中，位于任务窗格的底部。 日志位于 **“输出** ”选项卡上，但写入日志时，该选项卡会自动获得焦点。

日志不会影响工作簿。

## <a name="automate-tab-not-appearing-or-office-scripts-unavailable"></a>未显示自动执行选项卡或 Office 脚本不可用

以下步骤应有助于排查与“**自动”** 选项卡未出现在Excel web 版中相关的任何问题。

1. [确保 Microsoft 365 许可证包含 Office 脚本](../overview/excel.md#requirements)。
1. [检查是否支持浏览器](platform-limits.md#browser-support)。
1. [确保已启用第三方 Cookie](platform-limits.md#third-party-cookies)。
1. [确保管理员未在Microsoft 365 管理中心中禁用 Office 脚本](/microsoft-365/admin/manage/manage-office-scripts-settings)。
1. 确保未以外部用户或来宾用户身份登录到租户。

[!INCLUDE [Teams support note](../includes/teams-support-note.md)]

> [!NOTE]
> 有一个已知问题阻止存储在 SharePoint 中的脚本始终出现在最近使用的列表中。 当管理员关闭 Exchange Web 服务 (EWS) 时，会发生这种情况。 基于 SharePoint 的脚本仍可访问，可通过文件对话框使用。

## <a name="help-resources"></a>帮助资源

[Stack Overflow](https://stackoverflow.com/questions/tagged/office-scripts) 是一个开发人员社区，他们愿意帮助解决编码问题。 通常，可以通过快速的 Stack Overflow 搜索找到问题的解决方案。 如果没有，请提出问题，并使用“office-scripts”标记对其进行标记。 请务必提及你正在创建 Office *脚本*，而不是 Office *加载项*。

## <a name="see-also"></a>另请参阅

- [Office 脚本中的最佳实践](../develop/best-practices.md)
- [Office 脚本的平台限制](platform-limits.md)
- [提高 Office 脚本的性能](../develop/web-client-performance.md)
- [排查 PowerAutomate 中运行的 Office 脚本问题](power-automate-troubleshooting.md)
- [消除 Office 脚本的影响](undo.md)
