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
# <a name="troubleshooting-office-scripts"></a>Office 脚本疑难解答

开发 Office 脚本时，可能会产生错误。 没关系。 我们有一些工具，可帮助查找问题并使你的脚本完美运行。

## <a name="console-logs"></a>控制台日志

有时，在进行故障排除时，您需要将消息打印到屏幕。 这些值可显示变量的当前值或触发的代码路径。 为此，请将文本记录到控制台。

```TypeScript
console.log("Logging my range's address.");
myRange.load("address");
await context.sync();
console.log(myRange.address);
```

> [!IMPORTANT]
> 在记录对象`load`属性之前， `sync`请不要忘记工作表数据和工作簿。

传递给`console.log`的字符串将显示在代码编辑器的日志记录控制台中。 若要打开控制台，请按**省略号**按钮，然后选择 "**日志 ...** "。

日志不会影响工作簿。

## <a name="error-messages"></a>错误消息

如果 Excel 脚本在运行时遇到问题，则会产生错误。 您将看到提示窗口，询问您是否要**查看日志**。 按该按钮打开控制台并显示任何错误。

## <a name="help-resources"></a>帮助资源

[堆栈溢出](https://stackoverflow.com/questions/tagged/office-scripts)是一种愿意帮助处理编码问题的开发人员社区。 通常情况下，你可以通过快速堆栈溢出搜索找到问题的解决方案。 如果不是，请询问问题并使用 "office-scripts" 标记对其进行标记。 请务必指出您正在创建 Office*脚本*，而不是 office*外接程序*。

如果您遇到 Office JavaScript API 问题，请在[OfficeDev/Office js](https://github.com/OfficeDev/office-js) GitHub 存储库中创建问题。 产品团队的成员将响应问题并提供进一步的帮助。 在**OfficeDev/js**存储库中创建问题表示您在 OFFICE JavaScript API 库中发现产品团队应解决的缺陷。

如果操作记录器或编辑器存在问题，请通过 Excel 中的 "帮助" **> 反馈**按钮发送反馈。

## <a name="see-also"></a>另请参阅

- [Web 上的 Excel 中的 Office 脚本](../overview/excel.md)
- [Web 上的 Excel 中 Office 脚本的脚本基础](../develop/scripting-fundamentals.md)
- [撤消 Office 脚本的效果](undo.md)
