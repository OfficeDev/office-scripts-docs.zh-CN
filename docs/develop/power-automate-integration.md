---
title: 使用Office运行脚本Power Automate
description: 如何获取Office脚本Excel web 版工作流Power Automate脚本。
ms.date: 11/01/2021
ms.localizationpriority: medium
ms.openlocfilehash: 1a335944230011bc8f5967004b7394f3f5958321
ms.sourcegitcommit: 634ad2061e683ae1032c1e0b55b00ac577adc34f
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 11/03/2021
ms.locfileid: "60725591"
---
# <a name="run-office-scripts-with-power-automate"></a>使用Office运行脚本Power Automate

[Power Automate，](https://flow.microsoft.com)可以将Office脚本添加到更大的自动化工作流。 可以使用Power Automate操作，如将电子邮件内容添加到工作表表中，或在项目管理工具中基于工作簿注释创建操作。

## <a name="get-started"></a>开始行动

如果你刚开始Power Automate，我们建议你访问入门[Power Automate。](/power-automate/getting-started) 在那里，你可以了解有关所有可用的自动化可能性的信息。 此处的文档重点介绍脚本Office脚本Power Automate，以及这如何有助于改进Excel体验。

若要开始Power Automate脚本Office脚本，请按照教程[开始将脚本与脚本Power Automate。](../tutorials/excel-power-automate-manual.md) 这将教您如何创建调用简单脚本的流。 完成本教程和在自动运行的[Power Automate](../tutorials/excel-power-automate-trigger.md)流教程中将数据传递到脚本后，请返回此处，详细了解如何将 Office 脚本连接到 Power Automate 流。

## <a name="excel-online-business-connector"></a>ExcelOnline (Business) 连接器

[连接器](/connectors/connectors)是应用程序与Power Automate之间的桥梁。 使用[Excel Online (Business) 连接器](/connectors/excelonlinebusiness)，流可以访问Excel工作簿。 通过"运行脚本"操作，您可以调用Office工作簿访问的任何脚本。 还可以为脚本提供输入参数，以便流提供数据，或让脚本返回流中稍后步骤的信息。

> [!IMPORTANT]
> "运行脚本"操作为使用 Excel 连接器的人提供对工作簿及其数据的重要访问权限。 此外，执行外部 API 调用的脚本存在安全风险，如外部从 Power Automate[中Power Automate。](external-calls.md) 如果你的管理员关注高度敏感数据的曝光，他们可以通过 Excel 脚本管理员控制来关闭 Excel Online 连接器或Office[脚本Office访问](/microsoft-365/admin/manage/manage-office-scripts-settings)。

## <a name="data-transfer-in-flows-for-scripts"></a>脚本流中的数据传输

Power Automate允许你在流的步骤之间传递数据片段。 可以将脚本配置为接受所需的任何类型的信息，并返回流中所需的工作簿中的内容。 通过向函数添加参数来指定脚本的 (`main` 以及 `workbook: ExcelScript.Workbook`) 。 脚本的输出通过向 添加返回类型进行声明 `main` 。

> [!NOTE]
> 当您在流中创建"Run Script"块时，将填充接受的参数和返回的类型。 如果更改脚本的参数或返回类型，则需要恢复流的"运行脚本"块。 这可确保正确分析数据。

以下各节介绍用于脚本的输入和输出的详细信息Power Automate。 如果你想要实践学习本主题的方法，请尝试在自动运行的 Power Automate 流教程中将数据传递到脚本，或浏览自动[任务](../resources/scenarios/task-reminders.md)[提醒](../tutorials/excel-power-automate-trigger.md)示例方案。

### <a name="main-parameters-pass-data-to-a-script"></a>`main` 参数：将数据传递给脚本

所有脚本输入都指定为 函数的其他 `main` 参数。 例如，如果您希望脚本接受表示作为输入的名称的 ， `string` 则您需要将 `main` 签名更改为 `function main(workbook: ExcelScript.Workbook, name: string)` 。

在脚本中配置流Power Automate，可以将脚本输入指定为静态值、表达式或动态内容。 [](/power-automate/use-expressions-in-conditions) 有关单个服务连接器的详细信息，请参阅[Power Automate Connector 文档](/connectors/)。

向脚本函数添加输入参数 `main` 时，请考虑以下允许和限制。

1. 第一个参数必须为 类型 `ExcelScript.Workbook` 。 其参数名称无关紧要。

1. 每个参数都必须具有类型 (，如 `string` 或 `number`) 。

1. 支持基本 `string` 类型 `number` 、 、 、 、 `boolean` `unknown` 和 `object` `undefined` 。

1. 支持 `[]` () 类型中的数组。
    > [!IMPORTANT]
    > 请注意，该对象 `Array<T>` 不是受支持的参数类型。

1. 嵌套数组作为参数受支持， (作为返回类型) 。

1. 如果联合类型是属于单个类型文本（如文本）的 (，则允许 `"Left" | "Right"`) 。 支持未定义类型的联合也受支持 (如 `string | undefined`) 。

1. 如果对象类型包含类型 、支持的数组或其他受支持对象的属性 `string` `number` ，则 `boolean` 允许这些对象类型。 以下示例演示作为参数类型支持的嵌套对象：

    ```TypeScript
    // Office Scripts can return an Employee object because Position only contains strings and numbers.
    interface Employee {
        name: string;
        job: Position;
    }

    interface Position {
        id: number;
        title: string;
    }
    ```

1. 对象必须在脚本中定义其接口或类定义。 也可以匿名内联定义对象，如以下示例所示：

    ```TypeScript
    function main(workbook: ExcelScript.Workbook): {name: string, email: string}
    ```

1. 允许使用可选参数，并且可以使用可选修饰符参数进行 (`?` 例如 `function main(workbook: ExcelScript.Workbook, Name?: string)` ，) 。

1. 允许默认参数值 (例如 `async function main(workbook: ExcelScript.Workbook, Name: string = 'Jane Doe')` 。

### <a name="return-data-from-a-script"></a>从脚本返回数据

脚本可以从工作簿中返回要用作活动流中的动态Power Automate内容。 与输入参数一样，Power Automate对返回类型施加一些限制。

1. 支持 `string` 基本类型 、 `number` 、 、 `boolean` 和 `void` `undefined` 。

1. 用作返回类型的联合类型遵循与用作脚本参数时相同的限制。

1. 如果 `[]` () 、或 ，则允许使用 `string` `number` 数组类型 `boolean` 。 如果类型是受支持的联合或受支持的文字类型，则也允许它们。
    > [!IMPORTANT]
    > 请注意，该对象 `Array<T>` 不是受支持的返回类型。

1. 用作返回类型的对象类型遵循与用作脚本参数时相同的限制。

1. 支持隐式键入，尽管它必须遵循与定义类型相同的规则。

## <a name="example"></a>示例

以下屏幕截图显示了Power Automate分配问题时触发GitHub流。 [](https://github.com/) 该流运行一个脚本，该脚本将问题添加到工作簿Excel表中。 如果该表中存在五个或多个问题，则流将发送电子邮件提醒。

:::image type="content" source="../images/power-automate-parameter-return-sample.png" alt-text="显示Power Automate流的流编辑器。":::

脚本函数将问题 ID 和问题标题指定为输入参数，脚本返回问题 `main` 表中的行数。

```TypeScript
function main(
  workbook: ExcelScript.Workbook,
  issueId: string,
  issueTitle: string): number {
  // Get the "GitHub" worksheet.
  let worksheet = workbook.getWorksheet("GitHub");

  // Get the first table in this worksheet, which contains the table of GitHub issues.
  let issueTable = worksheet.getTables()[0];

  // Add the issue ID and issue title as a row.
  issueTable.addRow(-1, [issueId, issueTitle]);

  // Return the number of rows in the table, which represents how many issues are assigned to this user.
  return issueTable.getRangeBetweenHeaderAndTotal().getRowCount();
}
```

## <a name="see-also"></a>另请参阅

- [在Office中运行Excel web 版脚本Power Automate](../tutorials/excel-power-automate-manual.md)
- [将数据传递到自动运行的 Power Automate 流中的脚本](../tutorials/excel-power-automate-trigger.md)
- [从脚本返回数据到自动运行 Power Automated 流](../tutorials/excel-power-automate-returns.md)
- [有关使用脚本Power Automate疑Office信息](../testing/power-automate-troubleshooting.md)
- [Power Automate 入门](/power-automate/getting-started)
- [ExcelOnline (Business) 连接器参考文档](/connectors/excelonlinebusiness/)
