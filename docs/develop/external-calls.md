---
title: Office 脚本中的外部 API 调用支持
description: 在 Office 脚本中进行外部 API 调用的支持和指南。
ms.date: 06/25/2020
localization_priority: Normal
ms.openlocfilehash: ec8281551cbe7c500eee40ec86067e5efbfcfc31
ms.sourcegitcommit: aec3c971c6640429f89b6bb99d2c95ea06725599
ms.translationtype: Auto
ms.contentlocale: zh-CN
ms.lasthandoff: 06/25/2020
ms.locfileid: "44878723"
---
# <a name="external-api-call-support-in-office-scripts"></a>Office 脚本中的外部 API 调用支持

Office 脚本平台不支持对[外部 api](https://developer.mozilla.org/docs/Web/API)的调用。 但是，在适当的情况下，可以运行这些呼叫。 外部呼叫只能通过 Excel 客户端进行，而不是[在正常情况下](#external-calls-from-power-automate)的电源自动运行。

在平台的预览阶段使用外部 Api 时，脚本作者不应预期一致的行为。 这是因为 JavaScript 运行时如何管理与工作簿的交互。 脚本可能在 API 调用完成之前结束（或 `Promise` 完全解决）。 因此，不要依赖于对关键脚本方案的外部 Api。

> [!CAUTION]
> 外部调用可能会导致敏感数据暴露给不需要的终结点。 你的管理员可以针对此类呼叫建立防火墙保护。

## <a name="definition-files-for-external-apis"></a>外部 Api 的定义文件

外部 Api 的定义文件不包含在 Office 脚本中。 使用此类 Api 会为缺少的定义生成编译时错误。 Api 仍运行（尽管仅在通过 Excel 客户端运行时），如以下脚本所示：

```typescript
async function main(workbook: ExcelScript.Workbook): Promise <void> {
  /* The following line of code generates the error:
   * "Cannot find name 'fetch'".
   * It will still run and return the JSON from the testing service.
   */
  let fetchResult = await fetch('https://jsonplaceholder.typicode.com/todos/1');
  let json = await fetchResult.json();

  // Displays the content from https://jsonplaceholder.typicode.com/todos/1
  console.log(JSON.stringify(json));
}
```

## <a name="external-calls-from-power-automate"></a>来自电源自动执行的外部呼叫

使用 Power 自动化运行脚本时，任何外部 API 调用都将失败。 这是通过 Excel 客户端和 Power 自动化运行脚本之间的行为差异。 在将脚本生成到流中之前，请务必检查脚本中的这些引用。

> [!WARNING]
> 如果外部呼叫[Excel Online 连接器](/connectors/excelonlinebusiness)在 Power 自动化中出现故障，则可以帮助您降低现有数据丢失防护策略。 但是，在您的组织外和组织的防火墙外部，这些脚本将通过 "电源自动完成" 运行。 若要进一步保护此外部环境中的恶意用户，管理员可以控制 Office 脚本的使用。 您的管理员可以通过[Office 脚本管理员控件](https://support.microsoft.com/office/19d3c51a-6ca2-40ab-978d-60fa49554dcf)在 web 上禁用 excel Online 连接器，或在 web 上关闭 Excel 的 office 脚本。

## <a name="see-also"></a>另请参阅

- [在 Office 脚本中使用内置的 JavaScript 对象](javascript-objects.md)