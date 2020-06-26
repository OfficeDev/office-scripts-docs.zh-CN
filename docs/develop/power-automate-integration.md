---
title: 将 Office 脚本与电源自动化相集成
description: 如何在使用 Power 自动工作流的网站上获取适用于 Excel 的 Office 脚本。
ms.date: 06/24/2020
localization_priority: Normal
ms.openlocfilehash: 977d9c88d75c8070eb729a443b4e8bc9a32e456d
ms.sourcegitcommit: aec3c971c6640429f89b6bb99d2c95ea06725599
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 06/25/2020
ms.locfileid: "44878733"
---
# <a name="integrate-office-scripts-with-power-automate"></a>将 Office 脚本与电源自动化相集成

[Power 自动](https://flow.microsoft.com)将脚本集成到更大的工作流中。 您可以使用 Power 自动执行操作，例如，将电子邮件的内容添加到工作表的表中，或在基于工作簿注释的项目管理工具中创建操作。 如果你刚开始使用 "电源自动化"，我们建议[使用 Power 自动化获取访问入门](/power-automate/getting-started)。 在这里，你可以了解有关跨多个服务自动化工作流的详细信息。

> [!IMPORTANT]
> 目前，不能从[共享流](/power-automate/share-buttons)中运行 Office 脚本。 只有创建脚本的用户才能运行它，甚至可以通过 Power 自动化。

## <a name="getting-started"></a>入门

若要开始结合使用电源自动化功能和 Office 脚本，请遵循教程[开始使用启用电源自动化的脚本](../tutorials/excel-power-automate-manual.md)。 这将教您如何创建调用简单脚本的流。 完成本教程和[使用 Power 自动化教程自动运行脚本](../tutorials/excel-power-automate-trigger.md)后，请返回此处了解有关平台集成的详细信息。

## <a name="excel-online-business-connector"></a>Excel Online （业务）连接器

[连接器](/connectors/connectors)是电源自动化和应用程序之间的桥梁。 [Excel Online （业务）连接器](/connectors/excelonlinebusiness)提供对 excel 工作簿的流访问。 "运行脚本" 操作允许您调用任何可通过所选工作簿访问的 Office 脚本。 您不仅可以通过流运行脚本，还可以通过脚本在工作簿之间传递数据。

> [!IMPORTANT]
> "运行脚本" 操作为使用 Excel connector 的用户提供对工作簿及其数据的有效访问权限。 此外，还存在一些使用脚本进行外部 API 调用的安全风险，如[Power 自动化中的外部调用](external-calls.md)中所述。 如果您的管理员担心暴露高度敏感的数据，则可以关闭 Excel Online 连接器或限制对 Office 脚本的访问，方法是通过[Office 脚本管理员控件](https://support.microsoft.com/office/19d3c51a-6ca2-40ab-978d-60fa49554dcf)。

## <a name="passing-data-from-power-automate-into-a-script"></a>将数据从电源自动化传递到脚本中

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

## <a name="returning-data-from-a-script-back-to-power-automate"></a>将数据从脚本返回到增强功能自动化

脚本可以返回工作簿中的数据，以用作电源自动化流中的动态内容。 与输入参数一样，Power 自动化将一些限制放在返回类型上。

1. 支持基本类型 `string` 、 `number` 、 `boolean` `void` 和 `undefined` 。

2. 用作返回类型的联合类型遵循与用作脚本参数时相同的限制。

3. 如果数组类型为类型 `string` 、或，则允许使用数组类型 `number` `boolean` 。 如果类型是受支持的联合或受支持的文本类型，也可以使用它们。

4. 用作返回类型的对象类型遵循与用作脚本参数时相同的限制。

5. 虽然支持隐式键入，但它必须遵循与定义的类型相同的规则。

## <a name="avoid-using-relative-references"></a>避免使用相对引用

Power 自动在所选的 Excel 工作簿中代表你运行脚本。 在这种情况下，工作簿可能会关闭。 在运行时，任何依赖用户的当前状态（如）的 API `Workbook.getActiveWorksheet` 都将在通过电源自动运行时失败。 在设计脚本时，请务必对工作表和区域使用绝对引用。

如果从 Power 自动流中的脚本调用，以下函数将引发错误并失败。

- `Chart.activate`
- `Range.select`
- `Workbook.getActiveCell`
- `Workbook.getActiveChart`
- `Workbook.getActiveChartOrNullObject`
- `Workbook.getActiveSlicer`
- `Workbook.getActiveSlicerOrNullObject`
- `Workbook.getActiveWorksheet`
- `Workbook.getSelectedRange`
- `Workbook.getSelectedRanges`
- `Worksheet.activate`

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
- [自动运行具有 Power 自动化功能的脚本](../tutorials/excel-power-automate-trigger.md)
- [Excel 网页版中 Office 脚本的脚本基础](scripting-fundamentals.md)
- [Power Automate 入门](/power-automate/getting-started)
- [Excel Online （业务）连接器参考文档](/connectors/excelonlinebusiness/)
