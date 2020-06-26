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
# <a name="integrate-office-scripts-with-power-automate"></a><span data-ttu-id="b09cd-103">将 Office 脚本与电源自动化相集成</span><span class="sxs-lookup"><span data-stu-id="b09cd-103">Integrate Office Scripts with Power Automate</span></span>

<span data-ttu-id="b09cd-104">[Power 自动](https://flow.microsoft.com)将脚本集成到更大的工作流中。</span><span class="sxs-lookup"><span data-stu-id="b09cd-104">[Power Automate](https://flow.microsoft.com) integrates your script into a larger workflow.</span></span> <span data-ttu-id="b09cd-105">您可以使用 Power 自动执行操作，例如，将电子邮件的内容添加到工作表的表中，或在基于工作簿注释的项目管理工具中创建操作。</span><span class="sxs-lookup"><span data-stu-id="b09cd-105">You can use Power Automate do things like add the contents of an email to a worksheet's table or create actions in your project management tools based on workbook comments.</span></span> <span data-ttu-id="b09cd-106">如果你刚开始使用 "电源自动化"，我们建议[使用 Power 自动化获取访问入门](/power-automate/getting-started)。</span><span class="sxs-lookup"><span data-stu-id="b09cd-106">If you are new to Power Automate, we recommend visiting [Get started with Power Automate](/power-automate/getting-started).</span></span> <span data-ttu-id="b09cd-107">在这里，你可以了解有关跨多个服务自动化工作流的详细信息。</span><span class="sxs-lookup"><span data-stu-id="b09cd-107">There, you can learn more about automating your workflows across multiple services.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="b09cd-108">目前，不能从[共享流](/power-automate/share-buttons)中运行 Office 脚本。</span><span class="sxs-lookup"><span data-stu-id="b09cd-108">Currently, you can't run Office Scripts from a [shared flow](/power-automate/share-buttons).</span></span> <span data-ttu-id="b09cd-109">只有创建脚本的用户才能运行它，甚至可以通过 Power 自动化。</span><span class="sxs-lookup"><span data-stu-id="b09cd-109">Only the user who created a script can run it, even through Power Automate.</span></span>

## <a name="getting-started"></a><span data-ttu-id="b09cd-110">入门</span><span class="sxs-lookup"><span data-stu-id="b09cd-110">Getting started</span></span>

<span data-ttu-id="b09cd-111">若要开始结合使用电源自动化功能和 Office 脚本，请遵循教程[开始使用启用电源自动化的脚本](../tutorials/excel-power-automate-manual.md)。</span><span class="sxs-lookup"><span data-stu-id="b09cd-111">To begin combining Power Automate and Office Scripts, follow the tutorial [Start using scripts with Power Automate](../tutorials/excel-power-automate-manual.md).</span></span> <span data-ttu-id="b09cd-112">这将教您如何创建调用简单脚本的流。</span><span class="sxs-lookup"><span data-stu-id="b09cd-112">This will teach you how to create a flow that calls a simple script.</span></span> <span data-ttu-id="b09cd-113">完成本教程和[使用 Power 自动化教程自动运行脚本](../tutorials/excel-power-automate-trigger.md)后，请返回此处了解有关平台集成的详细信息。</span><span class="sxs-lookup"><span data-stu-id="b09cd-113">After you've completed that tutorial and the [Automatically run scripts with Power Automate](../tutorials/excel-power-automate-trigger.md) tutorial, return here to learn details about the platform integrations.</span></span>

## <a name="excel-online-business-connector"></a><span data-ttu-id="b09cd-114">Excel Online （业务）连接器</span><span class="sxs-lookup"><span data-stu-id="b09cd-114">Excel Online (Business) connector</span></span>

<span data-ttu-id="b09cd-115">[连接器](/connectors/connectors)是电源自动化和应用程序之间的桥梁。</span><span class="sxs-lookup"><span data-stu-id="b09cd-115">[Connectors](/connectors/connectors) are the bridges between Power Automate and applications.</span></span> <span data-ttu-id="b09cd-116">[Excel Online （业务）连接器](/connectors/excelonlinebusiness)提供对 excel 工作簿的流访问。</span><span class="sxs-lookup"><span data-stu-id="b09cd-116">The [Excel Online (Business) connector](/connectors/excelonlinebusiness) gives your flows access to Excel workbooks.</span></span> <span data-ttu-id="b09cd-117">"运行脚本" 操作允许您调用任何可通过所选工作簿访问的 Office 脚本。</span><span class="sxs-lookup"><span data-stu-id="b09cd-117">The "Run script" action lets you call any Office Script accessible through the selected workbook.</span></span> <span data-ttu-id="b09cd-118">您不仅可以通过流运行脚本，还可以通过脚本在工作簿之间传递数据。</span><span class="sxs-lookup"><span data-stu-id="b09cd-118">Not only can you run scripts through a flow, you can pass data to and from the workbook with the flow through the scripts.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="b09cd-119">"运行脚本" 操作为使用 Excel connector 的用户提供对工作簿及其数据的有效访问权限。</span><span class="sxs-lookup"><span data-stu-id="b09cd-119">The "Run script" action gives people who use the Excel connector significant access to your workbook and its data.</span></span> <span data-ttu-id="b09cd-120">此外，还存在一些使用脚本进行外部 API 调用的安全风险，如[Power 自动化中的外部调用](external-calls.md)中所述。</span><span class="sxs-lookup"><span data-stu-id="b09cd-120">Additionally, there are security risks with scripts that make external API calls, as explained in [External calls from Power Automate](external-calls.md).</span></span> <span data-ttu-id="b09cd-121">如果您的管理员担心暴露高度敏感的数据，则可以关闭 Excel Online 连接器或限制对 Office 脚本的访问，方法是通过[Office 脚本管理员控件](https://support.microsoft.com/office/19d3c51a-6ca2-40ab-978d-60fa49554dcf)。</span><span class="sxs-lookup"><span data-stu-id="b09cd-121">If your admin is concerned with the exposure of highly sensitive data, they can either turn off the Excel Online connector or restrict access to Office Scripts through the [Office Scripts administrator controls](https://support.microsoft.com/office/19d3c51a-6ca2-40ab-978d-60fa49554dcf).</span></span>

## <a name="passing-data-from-power-automate-into-a-script"></a><span data-ttu-id="b09cd-122">将数据从电源自动化传递到脚本中</span><span class="sxs-lookup"><span data-stu-id="b09cd-122">Passing data from Power Automate into a script</span></span>

<span data-ttu-id="b09cd-123">所有脚本输入都被指定为函数的附加参数 `main` 。</span><span class="sxs-lookup"><span data-stu-id="b09cd-123">All script input is specified as additional parameters for the `main` function.</span></span> <span data-ttu-id="b09cd-124">例如，如果您希望脚本接受一个 `string` 表示输入名称的，则会将 `main` 签名更改为 `function main(workbook: ExcelScript.Workbook, name: string)` 。</span><span class="sxs-lookup"><span data-stu-id="b09cd-124">For example, if you wanted a script to accept a `string` that represents a name as input, you would change the `main` signature to `function main(workbook: ExcelScript.Workbook, name: string)`.</span></span>

<span data-ttu-id="b09cd-125">当您在电源自动化中配置流时，您可以将脚本输入指定为静态值、[表达式](/power-automate/use-expressions-in-conditions)或动态内容。</span><span class="sxs-lookup"><span data-stu-id="b09cd-125">When you're configuring a flow in Power Automate, you can specify script input as static values, [expressions](/power-automate/use-expressions-in-conditions), or dynamic content.</span></span> <span data-ttu-id="b09cd-126">有关单个服务连接器的详细信息，请参阅[Power 自动连接器文档](/connectors/)中的。</span><span class="sxs-lookup"><span data-stu-id="b09cd-126">Details on an individual service's connector can be found in the [Power Automate Connector documentation](/connectors/).</span></span>

<span data-ttu-id="b09cd-127">向脚本函数中添加输入参数时 `main` ，请考虑以下余量和限制。</span><span class="sxs-lookup"><span data-stu-id="b09cd-127">When adding input parameters to a script's `main` function, consider the following allowances and restrictions.</span></span>

1. <span data-ttu-id="b09cd-128">第一个参数的类型必须为 `ExcelScript.Workbook` 。</span><span class="sxs-lookup"><span data-stu-id="b09cd-128">The first parameter must be of type `ExcelScript.Workbook`.</span></span> <span data-ttu-id="b09cd-129">其参数名称无关紧要。</span><span class="sxs-lookup"><span data-stu-id="b09cd-129">Its parameter name does not matter.</span></span>

2. <span data-ttu-id="b09cd-130">每个参数都必须具有一个类型。</span><span class="sxs-lookup"><span data-stu-id="b09cd-130">Every parameter must have a type.</span></span>

3. <span data-ttu-id="b09cd-131">支持基本类型 `string` 、、、、、 `number` `boolean` `any` `unknown` `object` 和 `undefined` 。</span><span class="sxs-lookup"><span data-stu-id="b09cd-131">The basic types `string`, `number`, `boolean`, `any`, `unknown`, `object`, and `undefined` are supported.</span></span>

4. <span data-ttu-id="b09cd-132">支持前面列出的基本类型的数组。</span><span class="sxs-lookup"><span data-stu-id="b09cd-132">Arrays of the previously listed basic types are supported.</span></span>

5. <span data-ttu-id="b09cd-133">嵌套的数组支持作为参数（而不是返回类型）。</span><span class="sxs-lookup"><span data-stu-id="b09cd-133">Nested arrays are supported as parameters (but not as return types).</span></span>

6. <span data-ttu-id="b09cd-134">如果联合类型是属于单个类型（ `string` 、或）的文本的联合，则允许联合类型 `number` `boolean` 。</span><span class="sxs-lookup"><span data-stu-id="b09cd-134">Union types are allowed if they are a union of literals belonging to a single type (`string`, `number`, or `boolean`).</span></span> <span data-ttu-id="b09cd-135">此外，还支持具有未定义的受支持类型的联合。</span><span class="sxs-lookup"><span data-stu-id="b09cd-135">Unions of a supported type with undefined are also supported.</span></span>

7. <span data-ttu-id="b09cd-136">如果对象类型包含类型 `string` 、 `number` 、、支持的 `boolean` 数组或其他受支持的对象的属性，则允许这些对象类型。</span><span class="sxs-lookup"><span data-stu-id="b09cd-136">Object types are allowed if they contain properties of type `string`, `number`, `boolean`, supported arrays, or other supported objects.</span></span> <span data-ttu-id="b09cd-137">下面的示例演示受支持为参数类型的嵌套对象：</span><span class="sxs-lookup"><span data-stu-id="b09cd-137">The following example shows nested objects that are supported as parameter types:</span></span>

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

8. <span data-ttu-id="b09cd-138">对象必须在脚本中定义其接口或类定义。</span><span class="sxs-lookup"><span data-stu-id="b09cd-138">Objects must have their interface or class definition defined in the script.</span></span> <span data-ttu-id="b09cd-139">也可以以匿名方式直接定义对象，如下面的示例所示：</span><span class="sxs-lookup"><span data-stu-id="b09cd-139">An object can also be defined anonymously inline, as in the following example:</span></span>

    ```TypeScript
    function main(workbook: ExcelScript.Workbook): {name: string, email: string}
    ```

9. <span data-ttu-id="b09cd-140">可选参数是允许的，并且可以使用 optional 修饰符 `?` （例如，）来表示 `function main(workbook: ExcelScript.Workbook, Name?: string)` 。</span><span class="sxs-lookup"><span data-stu-id="b09cd-140">Optional parameters are allowed and can be denoted as such by using the optional modifier `?` (for example, `function main(workbook: ExcelScript.Workbook, Name?: string)`).</span></span>

10. <span data-ttu-id="b09cd-141">允许使用默认参数值（例如 `async function main(workbook: ExcelScript.Workbook, Name: string = 'Jane Doe')` 。</span><span class="sxs-lookup"><span data-stu-id="b09cd-141">Default parameter values are allowed (for example `async function main(workbook: ExcelScript.Workbook, Name: string = 'Jane Doe')`.</span></span>

## <a name="returning-data-from-a-script-back-to-power-automate"></a><span data-ttu-id="b09cd-142">将数据从脚本返回到增强功能自动化</span><span class="sxs-lookup"><span data-stu-id="b09cd-142">Returning data from a script back to Power Automate</span></span>

<span data-ttu-id="b09cd-143">脚本可以返回工作簿中的数据，以用作电源自动化流中的动态内容。</span><span class="sxs-lookup"><span data-stu-id="b09cd-143">Scripts can return data from the workbook to be used as dynamic content in a Power Automate flow.</span></span> <span data-ttu-id="b09cd-144">与输入参数一样，Power 自动化将一些限制放在返回类型上。</span><span class="sxs-lookup"><span data-stu-id="b09cd-144">As with input parameters, Power Automate places some restrictions on the return type.</span></span>

1. <span data-ttu-id="b09cd-145">支持基本类型 `string` 、 `number` 、 `boolean` `void` 和 `undefined` 。</span><span class="sxs-lookup"><span data-stu-id="b09cd-145">The basic types `string`, `number`, `boolean`, `void`, and `undefined` are supported.</span></span>

2. <span data-ttu-id="b09cd-146">用作返回类型的联合类型遵循与用作脚本参数时相同的限制。</span><span class="sxs-lookup"><span data-stu-id="b09cd-146">Union types used as return types follow the same restrictions as they do when used as script parameters.</span></span>

3. <span data-ttu-id="b09cd-147">如果数组类型为类型 `string` 、或，则允许使用数组类型 `number` `boolean` 。</span><span class="sxs-lookup"><span data-stu-id="b09cd-147">Array types are allowed if they are of type `string`, `number`, or `boolean`.</span></span> <span data-ttu-id="b09cd-148">如果类型是受支持的联合或受支持的文本类型，也可以使用它们。</span><span class="sxs-lookup"><span data-stu-id="b09cd-148">They are also allowed if the type is a supported union or supported literal type.</span></span>

4. <span data-ttu-id="b09cd-149">用作返回类型的对象类型遵循与用作脚本参数时相同的限制。</span><span class="sxs-lookup"><span data-stu-id="b09cd-149">Object types used as return types follow the same restrictions as they do when used as script parameters.</span></span>

5. <span data-ttu-id="b09cd-150">虽然支持隐式键入，但它必须遵循与定义的类型相同的规则。</span><span class="sxs-lookup"><span data-stu-id="b09cd-150">Implicit typing is supported, though it must follow the same rules as a defined type.</span></span>

## <a name="avoid-using-relative-references"></a><span data-ttu-id="b09cd-151">避免使用相对引用</span><span class="sxs-lookup"><span data-stu-id="b09cd-151">Avoid using relative references</span></span>

<span data-ttu-id="b09cd-152">Power 自动在所选的 Excel 工作簿中代表你运行脚本。</span><span class="sxs-lookup"><span data-stu-id="b09cd-152">Power Automate runs your script in the chosen Excel workbook on your behalf.</span></span> <span data-ttu-id="b09cd-153">在这种情况下，工作簿可能会关闭。</span><span class="sxs-lookup"><span data-stu-id="b09cd-153">The workbook might be closed when this happens.</span></span> <span data-ttu-id="b09cd-154">在运行时，任何依赖用户的当前状态（如）的 API `Workbook.getActiveWorksheet` 都将在通过电源自动运行时失败。</span><span class="sxs-lookup"><span data-stu-id="b09cd-154">Any API that relies on the user's current state, such as `Workbook.getActiveWorksheet`, will fail when run through Power Automate.</span></span> <span data-ttu-id="b09cd-155">在设计脚本时，请务必对工作表和区域使用绝对引用。</span><span class="sxs-lookup"><span data-stu-id="b09cd-155">When designing your scripts, be sure to use absolute references for worksheets and ranges.</span></span>

<span data-ttu-id="b09cd-156">如果从 Power 自动流中的脚本调用，以下函数将引发错误并失败。</span><span class="sxs-lookup"><span data-stu-id="b09cd-156">The following functions will throw an error and fail when called from a script in a Power Automate flow.</span></span>

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

## <a name="example"></a><span data-ttu-id="b09cd-157">示例</span><span class="sxs-lookup"><span data-stu-id="b09cd-157">Example</span></span>

<span data-ttu-id="b09cd-158">下面的屏幕截图显示了只要向您分配[GitHub](https://github.com/)问题时触发的电源自动化流。</span><span class="sxs-lookup"><span data-stu-id="b09cd-158">The following screenshot shows a Power Automate flow that's triggered whenever a [GitHub](https://github.com/) issue is assigned to you.</span></span> <span data-ttu-id="b09cd-159">流运行一个将问题添加到 Excel 工作簿中的表的脚本。</span><span class="sxs-lookup"><span data-stu-id="b09cd-159">The flow runs a script that adds the issue to a table in an Excel workbook.</span></span> <span data-ttu-id="b09cd-160">如果该表中有五个或更多问题，流将发送电子邮件提醒。</span><span class="sxs-lookup"><span data-stu-id="b09cd-160">If there are five or more issues in that table, the flow sends an email reminder.</span></span>

![示例流，如 Power 自动化流编辑器中所示。](../images/power-automate-parameter-return-sample.png)

<span data-ttu-id="b09cd-162">`main`脚本的功能将问题 ID 和问题标题指定为输入参数，脚本将返回 "问题" 表中的行数。</span><span class="sxs-lookup"><span data-stu-id="b09cd-162">The `main` function of the script specifies the issue ID and issue title as input parameters, and the script returns the number of rows in the issue table.</span></span>

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

## <a name="see-also"></a><span data-ttu-id="b09cd-163">另请参阅</span><span class="sxs-lookup"><span data-stu-id="b09cd-163">See also</span></span>

- [<span data-ttu-id="b09cd-164">在使用 Power 自动化的 web 上运行 Excel 中的 Office 脚本</span><span class="sxs-lookup"><span data-stu-id="b09cd-164">Run Office Scripts in Excel on the web with Power Automate</span></span>](../tutorials/excel-power-automate-manual.md)
- [<span data-ttu-id="b09cd-165">自动运行具有 Power 自动化功能的脚本</span><span class="sxs-lookup"><span data-stu-id="b09cd-165">Automatically run scripts with Power Automate</span></span>](../tutorials/excel-power-automate-trigger.md)
- [<span data-ttu-id="b09cd-166">Excel 网页版中 Office 脚本的脚本基础</span><span class="sxs-lookup"><span data-stu-id="b09cd-166">Scripting fundamentals for Office Scripts in Excel on the web</span></span>](scripting-fundamentals.md)
- [<span data-ttu-id="b09cd-167">Power Automate 入门</span><span class="sxs-lookup"><span data-stu-id="b09cd-167">Get started with Power Automate</span></span>](/power-automate/getting-started)
- [<span data-ttu-id="b09cd-168">Excel Online （业务）连接器参考文档</span><span class="sxs-lookup"><span data-stu-id="b09cd-168">Excel Online (Business) connector reference documentation</span></span>](/connectors/excelonlinebusiness/)
