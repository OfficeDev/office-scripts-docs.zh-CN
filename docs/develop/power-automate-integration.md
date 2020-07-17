---
title: 使用 Power 自动运行 Office 脚本
description: 如何在使用 Power 自动工作流的网站上获取适用于 Excel 的 Office 脚本。
ms.date: 07/10/2020
localization_priority: Normal
ms.openlocfilehash: bd8fea08b7a9303ad2ceace787de6457a33fb979
ms.sourcegitcommit: ebd1079c7e2695ac0e7e4c616f2439975e196875
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 07/17/2020
ms.locfileid: "45160444"
---
# <a name="run-office-scripts-with-power-automate"></a>使用 Power 自动运行 Office 脚本

通过使用[电源自动化](https://flow.microsoft.com)，可以将 Office 脚本添加到更大的自动化工作流中。 您可以使用 Power 自动执行操作，例如，将电子邮件的内容添加到工作表的表中，或在基于工作簿注释的项目管理工具中创建操作。

## <a name="getting-started"></a>入门

如果你刚开始使用 "电源自动化"，我们建议[使用 Power 自动化获取访问入门](/power-automate/getting-started)。 在这里，你可以了解有关你可使用的所有自动化可能性的详细信息。 此处的文档重点介绍 Office 脚本与电源自动化的工作方式，以及如何帮助改进 Excel 体验。

若要开始结合使用电源自动化功能和 Office 脚本，请遵循教程[开始使用启用电源自动化的脚本](../tutorials/excel-power-automate-manual.md)。 这将教您如何创建调用简单脚本的流。 在完成了教程和将[数据传递到自动运行电源自动化流教程中的脚本](../tutorials/excel-power-automate-trigger.md)之后，请返回此处以了解有关连接 Office 脚本以实现自动处理功能流的详细信息。

## <a name="excel-online-business-connector"></a>Excel Online （业务）连接器

[连接器](/connectors/connectors)是电源自动化和应用程序之间的桥梁。 [Excel Online （业务）连接器](/connectors/excelonlinebusiness)提供对 excel 工作簿的流访问。 "运行脚本" 操作允许您调用任何可通过所选工作簿访问的 Office 脚本。 您不仅可以通过流运行脚本，还可以通过脚本在工作簿之间传递数据。

> [!IMPORTANT]
> "运行脚本" 操作为使用 Excel connector 的用户提供对工作簿及其数据的有效访问权限。 此外，还存在一些使用脚本进行外部 API 调用的安全风险，如[Power 自动化中的外部调用](external-calls.md)中所述。 如果您的管理员担心暴露高度敏感的数据，则可以关闭 Excel Online 连接器或限制对 Office 脚本的访问，方法是通过[Office 脚本管理员控件](https://support.microsoft.com/office/19d3c51a-6ca2-40ab-978d-60fa49554dcf)。

## <a name="data-transfer-in-flows-for-scripts"></a>脚本流中的数据传输

利用电源自动化，可以在流的各个步骤之间传递数据片段。 可以将脚本配置为接受所需的任何类型的信息，并从您的工作簿中返回您想要的任何内容。 您的脚本的输入通过向函数添加参数 `main` （除了）来指定 `workbook: ExcelScript.Workbook` 。 脚本中的输出通过将返回类型添加到来声明 `main` 。

> [!NOTE]
> 在流中创建 "运行脚本" 块时，将填充接受的参数和返回的类型。 如果更改了脚本的参数或返回类型，您将需要恢复流的 "运行脚本" 块。 这样可确保正确分析数据。

以下各节介绍了用于 Power 自动化的脚本输入和输出的详细信息。 如果你想要学习本主题的实践方法，请尝试[在自动运行电源自动化流教程中将数据传递到脚本](../tutorials/excel-power-automate-trigger.md)，或浏览[自动任务提醒](../resources/scenarios/task-reminders.md)示例方案。

### <a name="main-parameters-passing-data-to-a-script"></a>`main`参数：将数据传递给脚本

所有脚本输入都被指定为函数的附加参数 `main` 。 例如，如果您希望脚本接受一个 `string` 表示输入名称的，则会将 `main` 签名更改为 `function main(workbook: ExcelScript.Workbook, name: string)` 。

当您在电源自动化中配置流时，您可以将脚本输入指定为静态值、[表达式](/power-automate/use-expressions-in-conditions)或动态内容。 有关单个服务连接器的详细信息，请参阅[Power 自动连接器文档](/connectors/)中的。

向脚本函数中添加输入参数时 `main` ，请考虑以下余量和限制。

1. 第一个参数的类型必须为 `ExcelScript.Workbook` 。 其参数名称无关紧要。

2. 每个参数都必须具有一个类型。

3. 支持基本类型 `string` 、、、、、 `number` `boolean` `any` `unknown` `object` 和 `undefined` 。

4. 支持前面列出的基本类型的数组。

5. 嵌套的数组支持作为参数（而不是返回类型）。

6. 如果联合类型是属于单个类型（ `string` 、或）的文本的联合，则允许联合类型 `number` `boolean` 。 此外，还支持具有未定义的受支持类型的联合。

7. 如果对象类型包含类型 `string` 、 `number` 、、支持的 `boolean` 数组或其他受支持的对象的属性，则允许这些对象类型。 下面的示例演示受支持为参数类型的嵌套对象：

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

8. 对象必须在脚本中定义其接口或类定义。 也可以以匿名方式直接定义对象，如下面的示例所示：

    ```TypeScript
    function main(workbook: ExcelScript.Workbook): {name: string, email: string}
    ```

9. 可选参数是允许的，并且可以使用 optional 修饰符 `?` （例如，）来表示 `function main(workbook: ExcelScript.Workbook, Name?: string)` 。

10. 允许使用默认参数值（例如 `async function main(workbook: ExcelScript.Workbook, Name: string = 'Jane Doe')` 。

### <a name="returning-data-from-a-script"></a>从脚本中返回数据

脚本可以返回工作簿中的数据，以用作电源自动化流中的动态内容。 与输入参数一样，Power 自动化将一些限制放在返回类型上。

1. 支持基本类型 `string` 、 `number` 、 `boolean` `void` 和 `undefined` 。

2. 用作返回类型的联合类型遵循与用作脚本参数时相同的限制。

3. 如果数组类型为类型 `string` 、或，则允许使用数组类型 `number` `boolean` 。 如果类型是受支持的联合或受支持的文本类型，也可以使用它们。

4. 用作返回类型的对象类型遵循与用作脚本参数时相同的限制。

5. 虽然支持隐式键入，但它必须遵循与定义的类型相同的规则。

## <a name="avoid-using-relative-references"></a>避免使用相对引用

Power 自动在所选的 Excel 工作簿中代表你运行脚本。 在这种情况下，工作簿可能会关闭。 在运行时，任何依赖用户的当前状态（如）的 API `Workbook.getActiveWorksheet` 都将在通过电源自动运行时失败。 在设计脚本时，请务必对工作表和区域使用绝对引用。

如果从 Power 自动流中的脚本调用，以下方法将引发错误并失败。

| Class | 方法 |
|--|--|
| [Chart](/javascript/api/office-scripts/excelscript/excelscript.chart) | `activate` |
| [区域](/javascript/api/office-scripts/excelscript/excelscript.range) | `select` |
| [Workbook](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getActiveCell` |
| [Workbook](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getActiveChart` |
| [Workbook](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getActiveSlicer` |
| [Workbook](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getActiveWorksheet` |
| [Workbook](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getSelectedRange` |
| [Workbook](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getSelectedRanges` |
| [Worksheet](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `activate` |

## <a name="example"></a>示例

下面的屏幕截图显示了只要向您分配[GitHub](https://github.com/)问题时触发的电源自动化流。 流运行一个将问题添加到 Excel 工作簿中的表的脚本。 如果该表中有五个或更多问题，流将发送电子邮件提醒。

![示例流，如 Power 自动化流编辑器中所示。](../images/power-automate-parameter-return-sample.png)

`main`脚本的功能将问题 ID 和问题标题指定为输入参数，脚本将返回 "问题" 表中的行数。

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

- [在使用 Power 自动化的 web 上运行 Excel 中的 Office 脚本](../tutorials/excel-power-automate-manual.md)
- [在自动运行的电源自动化流中将数据传递给脚本](../tutorials/excel-power-automate-trigger.md)
- [Excel 网页版中 Office 脚本的脚本基础知识](scripting-fundamentals.md)
- [Power Automate 入门](/power-automate/getting-started)
- [Excel Online （业务）连接器参考文档](/connectors/excelonlinebusiness/)
