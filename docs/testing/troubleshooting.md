---
title: Office 脚本疑难解答
description: Office 脚本的调试提示和技术，以及帮助资源。
ms.date: 10/08/2020
localization_priority: Normal
ms.openlocfilehash: 9b3f4be778f3cdb4711d1e41d4d68f87ebca8152
ms.sourcegitcommit: 42fa3b629c93930b4e73e9c4c01d0c8bdf6d7487
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 10/09/2020
ms.locfileid: "48411542"
---
# <a name="troubleshooting-office-scripts"></a>Office 脚本疑难解答

开发 Office 脚本时，可能会产生错误。 没关系。 我们有一些工具，可帮助查找问题并使你的脚本完美运行。

## <a name="console-logs"></a>控制台日志

有时，在进行故障排除时，您需要将消息打印到屏幕。 这些值可显示变量的当前值或触发的代码路径。 为此，请将文本记录到控制台。

```TypeScript
console.log("Logging myRange's address.");
console.log(myRange.getAddress());
```

传递给的字符串 `console.log` 将显示在代码编辑器的日志记录控制台中。 若要打开控制台，请按**省略号**按钮，然后选择 "**日志 ...** "。

日志不会影响工作簿。

## <a name="error-messages"></a>错误消息

如果 Excel 脚本在运行时遇到问题，则会产生错误。 您将看到提示窗口，询问您是否要 **查看日志**。 按该按钮打开控制台并显示任何错误。

## <a name="automate-tab-not-appearing-or-office-scripts-unavailable"></a>"自动" 选项卡未显示或 Office 脚本不可用

以下步骤将帮助解决与 web 上的 Excel 中未出现的 " **自动** " 选项卡相关的任何问题。

1. [请确保你的 Microsoft 365 许可证包括 Office 脚本](../overview/excel.md#requirements)。
1. [让管理员启用该功能](/microsoft-365/admin/manage/manage-office-scripts-settings)。
1. [检查您的浏览器是否受支持](platform-limits.md#browser-support)。
1. [确保启用了第三方 cookie](platform-limits.md#third-party-cookies)。

[!INCLUDE [Teams support note](../includes/teams-support-note.md)]

## <a name="help-resources"></a>帮助资源

[堆栈溢出](https://stackoverflow.com/questions/tagged/office-scripts) 是一种愿意帮助处理编码问题的开发人员社区。 通常情况下，你可以通过快速堆栈溢出搜索找到问题的解决方案。 如果不是，请询问问题并使用 "office-scripts" 标记对其进行标记。 请务必指出您正在创建 Office *脚本*，而不是 office *外接程序*。

如果您遇到 Office JavaScript API 问题，请在 [OfficeDev/Office js](https://github.com/OfficeDev/office-js) GitHub 存储库中创建问题。 产品团队的成员将响应问题并提供进一步的帮助。 在 **OfficeDev/js** 存储库中创建问题表示您在 OFFICE JavaScript API 库中发现产品团队应解决的缺陷。

如果操作记录器或编辑器存在问题，请通过 Excel 中的 "帮助" **> 反馈** 按钮发送反馈。

## <a name="see-also"></a>另请参阅

- [Excel 网页版中的 Office 脚本](../overview/excel.md)
- [Web 上的 Excel 中 Office 脚本的脚本基础](../develop/scripting-fundamentals.md)
- [Office 脚本的平台限制](platform-limits.md)
- [提高 Office 脚本的性能](../develop/web-client-performance.md)
- [消除 Office 脚本的影响](undo.md)
