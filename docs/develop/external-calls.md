---
title: Office 脚本中的外部 API 呼叫支持
description: 在脚本中执行外部 API 调用Office指南。
ms.date: 05/21/2021
ms.localizationpriority: medium
ms.openlocfilehash: e7be505f13529e1d3bcff22ce9fa18cc36148f7b
ms.sourcegitcommit: 79ce4fad6d284b1aa71f5ad6d2938d9ad6a09fee
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/12/2022
ms.locfileid: "63459604"
---
# <a name="external-api-call-support-in-office-scripts"></a>Office 脚本中的外部 API 呼叫支持

脚本支持对外部服务的调用。 使用这些服务向工作簿提供数据和其他信息。

> [!CAUTION]
> 外部调用可能会导致敏感数据向不需要的终结点公开。 管理员可以针对此类呼叫建立防火墙保护。

> [!IMPORTANT]
> 对外部 API 的调用只能通过 Excel 应用程序进行，而在正常情况下Power Automate[调用](#external-calls-from-power-automate)。

## <a name="configure-your-script-for-external-calls"></a>为外部调用配置脚本

外部调用 [是异步](https://developer.mozilla.org/docs/Learn/JavaScript/Asynchronous/Async_await) 的，需要将脚本标记为 `async`。 将前缀 `async` 添加到 函数 `main` ，并返回 `Promise`，如下所示：

```typescript
async function main(workbook: ExcelScript.Workbook) : Promise <void>
```

> [!NOTE]
> 返回其他信息的脚本可以返回该 `Promise` 类型的 。 例如，如果您的脚本需要返回对象 `Employee` ，则返回签名为 `: Promise <Employee>`

您需要了解外部服务的接口，以调用该服务。 如果你使用的是 或 `fetch` [REST API](https://wikipedia.org/wiki/Representational_state_transfer)，则需要确定返回数据的 JSON 结构。 对于脚本的输入和输出，请考虑使 与 `interface` 所需的 JSON 结构相匹配。 这为脚本提供了更多的类型安全性。 有关此内容的示例，请参阅 [Using fetch from Office Scripts](../resources/samples/external-fetch-calls.md)。

### <a name="limitations-with-external-calls-from-office-scripts"></a>来自脚本的外部调用Office限制

* 无法登录或使用 OAuth2 类型的身份验证流。 所有密钥和凭据必须硬编码 (源文件进行硬编码) 。
* 没有用于存储 API 凭据和密钥的基础结构。 这必须由用户管理。
* 不支持文档 `localStorage`Cookie、 `sessionStorage` 和 对象。
* 外部调用可能会导致向不需要的终结点公开敏感数据，或导致外部数据进入内部工作簿。 管理员可以针对此类呼叫建立防火墙保护。 在依赖外部调用之前，请务必检查本地策略。
* 请务必在依赖关系之前检查数据吞吐量。 例如，下拉整个外部数据集可能不是最佳选择，而应该使用分页获取区块中的数据。

## <a name="retrieve-information-with-fetch"></a>使用 检索信息 `fetch`

提取 [API](https://developer.mozilla.org/docs/Web/API/Fetch_API) 从外部服务检索信息。 它是一 `async` 个 API，因此你需要调整 `main` 脚本的签名。 将 函数 `main` 制作为 `async`。 还应确保进行呼叫`await``fetch`和`json`检索。 这将确保在脚本结束之前完成这些操作。

检索到的任何 JSON 数据 `fetch` 都必须与脚本中定义的接口匹配。 返回的值必须分配给特定类型，因为Office[脚本不支持`any`该类型](typescript-restrictions.md#no-any-type-in-office-scripts)。 应参考服务文档，以查看返回的属性的名称和类型。 然后，将匹配的接口添加到脚本。

以下脚本使用 `fetch` 从给定 URL 中的测试服务器检索 JSON 数据。 `JSONData`请注意用于将数据存储为匹配类型的接口。

```TypeScript
async function main(workbook: ExcelScript.Workbook){
  // Retrieve sample JSON data from a test server.
  let fetchResult = await fetch('https://jsonplaceholder.typicode.com/todos/1');

  // Convert the returned data to the expected JSON structure.
  let json : JSONData = await fetchResult.json();

  // Display the content in a readable format.
  console.log(JSON.stringify(json));
}

/**
 * An interface that matches the returned JSON structure.
 * The property names match exactly.
 */
interface JSONData {
  userId: number;
  id: number;
  title: string;
  completed: boolean;
}
```

### <a name="other-fetch-samples"></a>其他 `fetch` 示例

* Use [external fetch calls in Office Scripts](../resources/samples/external-fetch-calls.md)示例演示如何获取有关用户的 GitHub 存储库的基本信息。
* Office脚本示例方案：[Graph NOAA](../resources/scenarios/noaa-data-fetch.md) 中的水级数据演示了用于从国家远地和省管理局的"目录和当前"数据库中检索记录的提取命令。

## <a name="external-calls-from-power-automate"></a>外部呼叫Power Automate

在使用脚本运行时，任何外部 API 调用Power Automate。 这是通过应用程序运行脚本和Excel运行脚本的行为Power Automate。 在将脚本构建到流中之前，请务必检查脚本中是否包含此类引用。

您必须将 [HTTP](/connectors/webcontents/) 与 Azure AD操作一起从外部服务拉取数据或推送到外部服务。

> [!WARNING]
> 通过 Power Automate [Excel Online](/connectors/excelonlinebusiness) 连接器进行的外部呼叫失败，以帮助制定现有数据丢失防护策略。 但是，通过 Power Automate运行的脚本在组织外部和组织的防火墙之外执行。 要在此外部环境中对恶意用户进行其他保护，管理员可以控制对脚本Office的使用。 管理员可以在 Power Automate 中禁用 Excel Online 连接器，Office脚本管理员Excel web 版禁用 Office [Scripts](/microsoft-365/admin/manage/manage-office-scripts-settings)。

## <a name="see-also"></a>另请参阅

* [在 Office 脚本中使用内置的 JavaScript 对象](javascript-objects.md)
* [在 Office 脚本中使用外部提取呼叫](../resources/samples/external-fetch-calls.md)
* [Office脚本示例方案：Graph NOAA 中的水级数据](../resources/scenarios/noaa-data-fetch.md)
