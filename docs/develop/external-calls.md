---
title: Office 脚本中的外部 API 呼叫支持
description: 在 Office 脚本中执行外部 API 调用的支持和指导。
ms.date: 01/05/2021
localization_priority: Normal
ms.openlocfilehash: 74b8750f609370370759ca4a4a1daa998363ac2e
ms.sourcegitcommit: 5d24e77df70aa2c1c982275d53213c2a9323ff86
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 04/02/2021
ms.locfileid: "51570309"
---
# <a name="external-api-call-support-in-office-scripts"></a>Office 脚本中的外部 API 呼叫支持

在平台的预览阶段使用外部 API[](https://developer.mozilla.org/docs/Web/API)时，脚本作者不应期望行为一致。 因此，不要依赖外部 API 实现关键脚本方案。

对外部 API 的调用只能通过 Excel 应用程序进行，而不是在正常情况下通过 Power Automate [进行](#external-calls-from-power-automate)。

> [!CAUTION]
> 外部调用可能会导致敏感数据向不需要的终结点公开。 管理员可以针对此类呼叫建立防火墙保护。

## <a name="working-with-fetch"></a>使用 `fetch`

提取 [API](https://developer.mozilla.org/docs/Web/API/Fetch_API) 从外部服务检索信息。 它是 `async` 一个 API，因此你需要调整 `main` 脚本的签名。 创建 `main` 函数 `async` ，并返回 `Promise<void>` 。 还应确保进行 `await` `fetch` 呼叫和 `json` 检索。 这将确保在脚本结束之前完成这些操作。

以下脚本使用 `fetch` 从给定 URL 中的测试服务器检索 JSON 数据。

```TypeScript
async function main(workbook: ExcelScript.Workbook): Promise <void> {
  /* 
   * Retrieve JSON data from a test server.
   */
  let fetchResult = await fetch('https://jsonplaceholder.typicode.com/todos/1');
  let json = await fetchResult.json();

  // Displays the content from https://jsonplaceholder.typicode.com/todos/1
  console.log(JSON.stringify(json));
}
```

[Office 脚本示例方案：来自 NOAA 的图形](../resources/scenarios/noaa-data-fetch.md)水级数据演示了用于从"国家/地区"和"国家/地区管理中心"的"目录和当前"数据库中检索记录的提取命令。

## <a name="external-calls-from-power-automate"></a>来自 Power Automate 的外部呼叫

使用 Power Automate 运行脚本时，任何外部 API 调用都失败。 这是通过 Excel 客户端和 Power Automate 运行脚本的行为差异。 在将脚本构建到流中之前，请务必检查脚本中是否包含此类引用。

> [!WARNING]
> 通过 Power Automate [Excel Online](/connectors/excelonlinebusiness) 连接器进行的外部调用失败，以帮助制定现有数据丢失防护策略。 但是，通过 Power Automate 运行的脚本在组织外部和组织的防火墙之外执行。 要在此外部环境中对恶意用户进行其他保护，管理员可以控制 Office 脚本的使用。 管理员可以在 Power Automate 中禁用 Excel Online 连接器，或者通过 Office 脚本管理员控件关闭 Excel 网页版 [Office 脚本](/microsoft-365/admin/manage/manage-office-scripts-settings)。

## <a name="see-also"></a>另请参阅

- [在 Office 脚本中使用内置的 JavaScript 对象](javascript-objects.md)
- [Office 脚本示例方案：GRAPH NOAA 中的水级数据](../resources/scenarios/noaa-data-fetch.md)
