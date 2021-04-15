---
title: 使用 Power Automate 运行 Office 脚本
description: 如何让适用于 Excel 网页的 Office 脚本与 Power Automate 工作流一起运行。
ms.date: 12/16/2020
localization_priority: Normal
ms.openlocfilehash: 1ca9aa14efe7cf2c91100a32fbc9a69054012f06
ms.sourcegitcommit: 45ffe3dbd2c834b78592ad35928cf8096f5e80bc
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 04/14/2021
ms.locfileid: "51755068"
---
# <a name="run-office-scripts-with-power-automate"></a><span data-ttu-id="20dda-103">使用 Power Automate 运行 Office 脚本</span><span class="sxs-lookup"><span data-stu-id="20dda-103">Run Office Scripts with Power Automate</span></span>

<span data-ttu-id="20dda-104">[Power Automate](https://flow.microsoft.com) 允许你将 Office 脚本添加到更大的自动化工作流。</span><span class="sxs-lookup"><span data-stu-id="20dda-104">[Power Automate](https://flow.microsoft.com) lets you add Office Scripts to a larger, automated workflow.</span></span> <span data-ttu-id="20dda-105">可以使用 Power Automate 执行一些操作，如将电子邮件内容添加到工作表表中，或在项目管理工具中基于工作簿注释创建操作。</span><span class="sxs-lookup"><span data-stu-id="20dda-105">You can use Power Automate do things like add the contents of an email to a worksheet's table or create actions in your project management tools based on workbook comments.</span></span>

## <a name="getting-started"></a><span data-ttu-id="20dda-106">入门</span><span class="sxs-lookup"><span data-stu-id="20dda-106">Getting started</span></span>

<span data-ttu-id="20dda-107">如果你刚开始使用 Power Automate，我们建议访问 Power [Automate 入门](/power-automate/getting-started)。</span><span class="sxs-lookup"><span data-stu-id="20dda-107">If you are new to Power Automate, we recommend visiting [Get started with Power Automate](/power-automate/getting-started).</span></span> <span data-ttu-id="20dda-108">在那里，你可以了解有关所有可用的自动化可能性的信息。</span><span class="sxs-lookup"><span data-stu-id="20dda-108">There, you can learn more about all the automation possibilities available to you.</span></span> <span data-ttu-id="20dda-109">此处的文档重点介绍 Office 脚本如何与 Power Automate 一起运行，以及这如何有助于改善 Excel 体验。</span><span class="sxs-lookup"><span data-stu-id="20dda-109">The documents here focus on how Office Scripts work with Power Automate and how that can help improve your Excel experience.</span></span>

<span data-ttu-id="20dda-110">若要开始组合 Power Automate 和 Office 脚本，请按照教程开始使用 Power [Automate 中的脚本](../tutorials/excel-power-automate-manual.md)。</span><span class="sxs-lookup"><span data-stu-id="20dda-110">To begin combining Power Automate and Office Scripts, follow the tutorial [Start using scripts with Power Automate](../tutorials/excel-power-automate-manual.md).</span></span> <span data-ttu-id="20dda-111">这将教您如何创建调用简单脚本的流。</span><span class="sxs-lookup"><span data-stu-id="20dda-111">This will teach you how to create a flow that calls a simple script.</span></span> <span data-ttu-id="20dda-112">完成本教程和自动运行的 [Power Automate](../tutorials/excel-power-automate-trigger.md) 流教程中的"将数据传递到脚本"教程后，请返回此处，详细了解如何连接 Office 脚本到 Power Automate 流。</span><span class="sxs-lookup"><span data-stu-id="20dda-112">After you've completed that tutorial and the [Pass data to scripts in an automatically-run Power Automate flow](../tutorials/excel-power-automate-trigger.md) tutorial, return here for detailed information about connecting Office Scripts to Power Automate flows.</span></span>

## <a name="excel-online-business-connector"></a><span data-ttu-id="20dda-113">Excel Online (Business) 连接器</span><span class="sxs-lookup"><span data-stu-id="20dda-113">Excel Online (Business) connector</span></span>

<span data-ttu-id="20dda-114">[连接器是](/connectors/connectors) Power Automate 和应用程序之间的桥梁。</span><span class="sxs-lookup"><span data-stu-id="20dda-114">[Connectors](/connectors/connectors) are the bridges between Power Automate and applications.</span></span> <span data-ttu-id="20dda-115">Excel [Online (Business) 连接器](/connectors/excelonlinebusiness) 可让你流访问 Excel 工作簿。</span><span class="sxs-lookup"><span data-stu-id="20dda-115">The [Excel Online (Business) connector](/connectors/excelonlinebusiness) gives your flows access to Excel workbooks.</span></span> <span data-ttu-id="20dda-116">通过"运行脚本"操作，您可以调用可通过所选工作簿访问的任何 Office 脚本。</span><span class="sxs-lookup"><span data-stu-id="20dda-116">The "Run script" action lets you call any Office Script accessible through the selected workbook.</span></span> <span data-ttu-id="20dda-117">还可以为脚本提供输入参数，以便流提供数据，或让脚本返回流中稍后步骤的信息。</span><span class="sxs-lookup"><span data-stu-id="20dda-117">You can also give your scripts input parameters so data can be provided by the flow, or have your script return information for later steps in the flow.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="20dda-118">"运行脚本"操作为使用 Excel 连接器的人提供对工作簿及其数据的重要访问权限。</span><span class="sxs-lookup"><span data-stu-id="20dda-118">The "Run script" action gives people who use the Excel connector significant access to your workbook and its data.</span></span> <span data-ttu-id="20dda-119">此外，执行外部 API 调用的脚本存在安全风险，如来自 [Power Automate 的外部调用中介绍](external-calls.md)。</span><span class="sxs-lookup"><span data-stu-id="20dda-119">Additionally, there are security risks with scripts that make external API calls, as explained in [External calls from Power Automate](external-calls.md).</span></span> <span data-ttu-id="20dda-120">如果你的管理员关注高度敏感数据的曝光，他们可以通过 Office 脚本管理员控件关闭 Excel Online 连接器或限制对 Office [脚本的访问](/microsoft-365/admin/manage/manage-office-scripts-settings)。</span><span class="sxs-lookup"><span data-stu-id="20dda-120">If your admin is concerned with the exposure of highly sensitive data, they can either turn off the Excel Online connector or restrict access to Office Scripts through the [Office Scripts administrator controls](/microsoft-365/admin/manage/manage-office-scripts-settings).</span></span>

## <a name="data-transfer-in-flows-for-scripts"></a><span data-ttu-id="20dda-121">脚本流中的数据传输</span><span class="sxs-lookup"><span data-stu-id="20dda-121">Data transfer in flows for scripts</span></span>

<span data-ttu-id="20dda-122">Power Automate 允许你在流的步骤之间传递数据片段。</span><span class="sxs-lookup"><span data-stu-id="20dda-122">Power Automate lets you pass pieces of data between steps of your flow.</span></span> <span data-ttu-id="20dda-123">可以将脚本配置为接受所需的任何类型的信息，并返回流中所需的工作簿中的内容。</span><span class="sxs-lookup"><span data-stu-id="20dda-123">Scripts can be configured to accept whatever types of information you need and return anything from your workbook that you want in your flow.</span></span> <span data-ttu-id="20dda-124">通过向函数添加参数来指定脚本的输入 (`main` 以及 `workbook: ExcelScript.Workbook`) 。</span><span class="sxs-lookup"><span data-stu-id="20dda-124">Input for your script is specified by adding parameters to the `main` function (in addition to `workbook: ExcelScript.Workbook`).</span></span> <span data-ttu-id="20dda-125">脚本的输出通过向 添加返回类型进行声明 `main` 。</span><span class="sxs-lookup"><span data-stu-id="20dda-125">Output from the script is declared by adding a return type to `main`.</span></span>

> [!NOTE]
> <span data-ttu-id="20dda-126">当您在流中创建"Run Script"块时，将填充接受的参数和返回的类型。</span><span class="sxs-lookup"><span data-stu-id="20dda-126">When you create a "Run Script" block in your flow, the accepted parameters and returned types are populated.</span></span> <span data-ttu-id="20dda-127">如果更改脚本的参数或返回类型，则需要恢复流的"运行脚本"块。</span><span class="sxs-lookup"><span data-stu-id="20dda-127">If you change the parameters or return types of your script, you'll need to redo the "Run script" block of your flow.</span></span> <span data-ttu-id="20dda-128">这可确保正确分析数据。</span><span class="sxs-lookup"><span data-stu-id="20dda-128">This ensures the data is being parsed correctly.</span></span>

<span data-ttu-id="20dda-129">以下各节介绍 Power Automate 中使用的脚本的输入和输出的详细信息。</span><span class="sxs-lookup"><span data-stu-id="20dda-129">The following sections cover the details of input and output for scripts used in Power Automate.</span></span> <span data-ttu-id="20dda-130">如果你想要实践学习本主题的方法，请尝试在自动运行的 [Power Automate](../tutorials/excel-power-automate-trigger.md) 流教程中将数据传递到脚本，或浏览自动 [任务](../resources/scenarios/task-reminders.md) 提醒示例方案。</span><span class="sxs-lookup"><span data-stu-id="20dda-130">If you'd like a hands-on approach to learning this topic, try out the [Pass data to scripts in an automatically-run Power Automate flow](../tutorials/excel-power-automate-trigger.md) tutorial or explore the [Automated task reminders](../resources/scenarios/task-reminders.md) sample scenario.</span></span>

### <a name="main-parameters-passing-data-to-a-script"></a><span data-ttu-id="20dda-131">`main` 参数：将数据传递给脚本</span><span class="sxs-lookup"><span data-stu-id="20dda-131">`main` Parameters: Passing data to a script</span></span>

<span data-ttu-id="20dda-132">所有脚本输入都指定为 函数的其他 `main` 参数。</span><span class="sxs-lookup"><span data-stu-id="20dda-132">All script input is specified as additional parameters for the `main` function.</span></span> <span data-ttu-id="20dda-133">例如，如果您希望脚本接受表示作为输入的名称的 ， `string` 则您需要将 `main` 签名更改为 `function main(workbook: ExcelScript.Workbook, name: string)` 。</span><span class="sxs-lookup"><span data-stu-id="20dda-133">For example, if you wanted a script to accept a `string` that represents a name as input, you would change the `main` signature to `function main(workbook: ExcelScript.Workbook, name: string)`.</span></span>

<span data-ttu-id="20dda-134">在 Power Automate 中配置流时，可以将脚本输入指定为静态值、 [表达式](/power-automate/use-expressions-in-conditions)或动态内容。</span><span class="sxs-lookup"><span data-stu-id="20dda-134">When you're configuring a flow in Power Automate, you can specify script input as static values, [expressions](/power-automate/use-expressions-in-conditions), or dynamic content.</span></span> <span data-ttu-id="20dda-135">有关单个服务连接器的详细信息，请参阅 [Power Automate Connector 文档](/connectors/)。</span><span class="sxs-lookup"><span data-stu-id="20dda-135">Details on an individual service's connector can be found in the [Power Automate Connector documentation](/connectors/).</span></span>

<span data-ttu-id="20dda-136">向脚本函数添加输入参数 `main` 时，请考虑以下允许和限制。</span><span class="sxs-lookup"><span data-stu-id="20dda-136">When adding input parameters to a script's `main` function, consider the following allowances and restrictions.</span></span>

1. <span data-ttu-id="20dda-137">第一个参数必须为 类型 `ExcelScript.Workbook` 。</span><span class="sxs-lookup"><span data-stu-id="20dda-137">The first parameter must be of type `ExcelScript.Workbook`.</span></span> <span data-ttu-id="20dda-138">其参数名称无关紧要。</span><span class="sxs-lookup"><span data-stu-id="20dda-138">Its parameter name does not matter.</span></span>

2. <span data-ttu-id="20dda-139">每个参数都必须具有类型 (，如 `string` 或 `number`) 。</span><span class="sxs-lookup"><span data-stu-id="20dda-139">Every parameter must have a type (such as `string` or `number`).</span></span>

3. <span data-ttu-id="20dda-140">支持基本类型 `string` `number` 、 、 、 、 `boolean` 、 `any` 和 `unknown` `object` `undefined` 。</span><span class="sxs-lookup"><span data-stu-id="20dda-140">The basic types `string`, `number`, `boolean`, `any`, `unknown`, `object`, and `undefined` are supported.</span></span>

4. <span data-ttu-id="20dda-141">支持前面列出的基本类型的数组。</span><span class="sxs-lookup"><span data-stu-id="20dda-141">Arrays of the previously listed basic types are supported.</span></span>

5. <span data-ttu-id="20dda-142">嵌套数组作为参数受支持， (作为返回类型) 。</span><span class="sxs-lookup"><span data-stu-id="20dda-142">Nested arrays are supported as parameters (but not as return types).</span></span>

6. <span data-ttu-id="20dda-143">如果联合类型是属于单个类型文本（如文本）的 (，则允许 `"Left" | "Right"`) 。</span><span class="sxs-lookup"><span data-stu-id="20dda-143">Union types are allowed if they are a union of literals belonging to a single type (such as `"Left" | "Right"`).</span></span> <span data-ttu-id="20dda-144">支持未定义类型的联合也受支持 (如 `string | undefined`) 。</span><span class="sxs-lookup"><span data-stu-id="20dda-144">Unions of a supported type with undefined are also supported (such as `string | undefined`).</span></span>

7. <span data-ttu-id="20dda-145">如果对象类型包含类型 、支持的数组或其他受支持对象的属性 `string` `number` ，则 `boolean` 允许这些对象类型。</span><span class="sxs-lookup"><span data-stu-id="20dda-145">Object types are allowed if they contain properties of type `string`, `number`, `boolean`, supported arrays, or other supported objects.</span></span> <span data-ttu-id="20dda-146">以下示例演示作为参数类型支持的嵌套对象：</span><span class="sxs-lookup"><span data-stu-id="20dda-146">The following example shows nested objects that are supported as parameter types:</span></span>

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

8. <span data-ttu-id="20dda-147">对象必须在脚本中定义其接口或类定义。</span><span class="sxs-lookup"><span data-stu-id="20dda-147">Objects must have their interface or class definition defined in the script.</span></span> <span data-ttu-id="20dda-148">也可以匿名内联定义对象，如以下示例所示：</span><span class="sxs-lookup"><span data-stu-id="20dda-148">An object can also be defined anonymously inline, as in the following example:</span></span>

    ```TypeScript
    function main(workbook: ExcelScript.Workbook): {name: string, email: string}
    ```

9. <span data-ttu-id="20dda-149">允许使用可选参数，并且可以使用可选修饰符参数进行 (`?` 例如 `function main(workbook: ExcelScript.Workbook, Name?: string)` ，) 。</span><span class="sxs-lookup"><span data-stu-id="20dda-149">Optional parameters are allowed and can be denoted as such by using the optional modifier `?` (for example, `function main(workbook: ExcelScript.Workbook, Name?: string)`).</span></span>

10. <span data-ttu-id="20dda-150">允许默认参数值 (例如 `async function main(workbook: ExcelScript.Workbook, Name: string = 'Jane Doe')` 。</span><span class="sxs-lookup"><span data-stu-id="20dda-150">Default parameter values are allowed (for example `async function main(workbook: ExcelScript.Workbook, Name: string = 'Jane Doe')`.</span></span>

### <a name="returning-data-from-a-script"></a><span data-ttu-id="20dda-151">从脚本返回数据</span><span class="sxs-lookup"><span data-stu-id="20dda-151">Returning data from a script</span></span>

<span data-ttu-id="20dda-152">脚本可以从工作簿中返回数据，以用作 Power Automate 流中的动态内容。</span><span class="sxs-lookup"><span data-stu-id="20dda-152">Scripts can return data from the workbook to be used as dynamic content in a Power Automate flow.</span></span> <span data-ttu-id="20dda-153">与输入参数一样，Power Automate 对返回类型施加了一些限制。</span><span class="sxs-lookup"><span data-stu-id="20dda-153">As with input parameters, Power Automate places some restrictions on the return type.</span></span>

1. <span data-ttu-id="20dda-154">支持 `string` 基本类型 、 `number` 、 、 `boolean` 和 `void` `undefined` 。</span><span class="sxs-lookup"><span data-stu-id="20dda-154">The basic types `string`, `number`, `boolean`, `void`, and `undefined` are supported.</span></span>

2. <span data-ttu-id="20dda-155">用作返回类型的联合类型遵循与用作脚本参数时相同的限制。</span><span class="sxs-lookup"><span data-stu-id="20dda-155">Union types used as return types follow the same restrictions as they do when used as script parameters.</span></span>

3. <span data-ttu-id="20dda-156">如果数组类型为 、 或 ，则 `string` `number` 允许使用数组类型 `boolean` 。</span><span class="sxs-lookup"><span data-stu-id="20dda-156">Array types are allowed if they are of type `string`, `number`, or `boolean`.</span></span> <span data-ttu-id="20dda-157">如果类型是受支持的联合或受支持的文字类型，则也允许它们。</span><span class="sxs-lookup"><span data-stu-id="20dda-157">They are also allowed if the type is a supported union or supported literal type.</span></span>

4. <span data-ttu-id="20dda-158">用作返回类型的对象类型遵循与用作脚本参数时相同的限制。</span><span class="sxs-lookup"><span data-stu-id="20dda-158">Object types used as return types follow the same restrictions as they do when used as script parameters.</span></span>

5. <span data-ttu-id="20dda-159">支持隐式键入，尽管它必须遵循与定义类型相同的规则。</span><span class="sxs-lookup"><span data-stu-id="20dda-159">Implicit typing is supported, though it must follow the same rules as a defined type.</span></span>

## <a name="example"></a><span data-ttu-id="20dda-160">示例</span><span class="sxs-lookup"><span data-stu-id="20dda-160">Example</span></span>

<span data-ttu-id="20dda-161">以下屏幕截图显示了每当分配 [GitHub](https://github.com/) 问题时触发的 Power Automate 流。</span><span class="sxs-lookup"><span data-stu-id="20dda-161">The following screenshot shows a Power Automate flow that's triggered whenever a [GitHub](https://github.com/) issue is assigned to you.</span></span> <span data-ttu-id="20dda-162">该流运行一个脚本，该脚本将问题添加到 Excel 工作簿的表中。</span><span class="sxs-lookup"><span data-stu-id="20dda-162">The flow runs a script that adds the issue to a table in an Excel workbook.</span></span> <span data-ttu-id="20dda-163">如果该表中存在五个或多个问题，则流将发送电子邮件提醒。</span><span class="sxs-lookup"><span data-stu-id="20dda-163">If there are five or more issues in that table, the flow sends an email reminder.</span></span>

:::image type="content" source="../images/power-automate-parameter-return-sample.png" alt-text="显示示例流的 Power Automate 流编辑器。":::

<span data-ttu-id="20dda-165">脚本函数将问题 ID 和问题标题指定为输入参数，脚本返回问题 `main` 表中的行数。</span><span class="sxs-lookup"><span data-stu-id="20dda-165">The `main` function of the script specifies the issue ID and issue title as input parameters, and the script returns the number of rows in the issue table.</span></span>

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

## <a name="see-also"></a><span data-ttu-id="20dda-166">另请参阅</span><span class="sxs-lookup"><span data-stu-id="20dda-166">See also</span></span>

- [<span data-ttu-id="20dda-167">使用 Power Automate 在 Excel 网页中运行 Office 脚本</span><span class="sxs-lookup"><span data-stu-id="20dda-167">Run Office Scripts in Excel on the web with Power Automate</span></span>](../tutorials/excel-power-automate-manual.md)
- [<span data-ttu-id="20dda-168">将数据传递到自动运行的 Power Automate 流中的脚本</span><span class="sxs-lookup"><span data-stu-id="20dda-168">Pass data to scripts in an automatically-run Power Automate flow</span></span>](../tutorials/excel-power-automate-trigger.md)
- [<span data-ttu-id="20dda-169">从脚本返回数据到自动运行 Power Automated 流</span><span class="sxs-lookup"><span data-stu-id="20dda-169">Return data from a script to an automatically-run Power Automate flow</span></span>](../tutorials/excel-power-automate-returns.md)
- [<span data-ttu-id="20dda-170">Power Automate with Office Scripts 疑难解答信息</span><span class="sxs-lookup"><span data-stu-id="20dda-170">Troubleshooting information for Power Automate with Office Scripts</span></span>](../testing/power-automate-troubleshooting.md)
- [<span data-ttu-id="20dda-171">Power Automate 入门</span><span class="sxs-lookup"><span data-stu-id="20dda-171">Get started with Power Automate</span></span>](/power-automate/getting-started)
- [<span data-ttu-id="20dda-172">Excel Online (Business) 连接器参考文档</span><span class="sxs-lookup"><span data-stu-id="20dda-172">Excel Online (Business) connector reference documentation</span></span>](/connectors/excelonlinebusiness/)
