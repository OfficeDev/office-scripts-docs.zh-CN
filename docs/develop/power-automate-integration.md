---
title: 使用 Power 自动运行 Office 脚本
description: 如何在使用 Power 自动工作流的网站上获取适用于 Excel 的 Office 脚本。
ms.date: 07/24/2020
localization_priority: Normal
ms.openlocfilehash: a427948847d7ab84962cdede7fb44d214592909f
ms.sourcegitcommit: ff7fde04ce5a66d8df06ed505951c8111e2e9833
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 08/11/2020
ms.locfileid: "46616673"
---
# <a name="run-office-scripts-with-power-automate"></a><span data-ttu-id="fe29e-103">使用 Power 自动运行 Office 脚本</span><span class="sxs-lookup"><span data-stu-id="fe29e-103">Run Office Scripts with Power Automate</span></span>

<span data-ttu-id="fe29e-104">通过使用[电源自动化](https://flow.microsoft.com)，可以将 Office 脚本添加到更大的自动化工作流中。</span><span class="sxs-lookup"><span data-stu-id="fe29e-104">[Power Automate](https://flow.microsoft.com) lets you add Office Scripts to a larger, automated workflow.</span></span> <span data-ttu-id="fe29e-105">您可以使用 Power 自动执行操作，例如，将电子邮件的内容添加到工作表的表中，或在基于工作簿注释的项目管理工具中创建操作。</span><span class="sxs-lookup"><span data-stu-id="fe29e-105">You can use Power Automate do things like add the contents of an email to a worksheet's table or create actions in your project management tools based on workbook comments.</span></span>

## <a name="getting-started"></a><span data-ttu-id="fe29e-106">入门</span><span class="sxs-lookup"><span data-stu-id="fe29e-106">Getting started</span></span>

<span data-ttu-id="fe29e-107">如果你刚开始使用 "电源自动化"，我们建议[使用 Power 自动化获取访问入门](/power-automate/getting-started)。</span><span class="sxs-lookup"><span data-stu-id="fe29e-107">If you are new to Power Automate, we recommend visiting [Get started with Power Automate](/power-automate/getting-started).</span></span> <span data-ttu-id="fe29e-108">在这里，你可以了解有关你可使用的所有自动化可能性的详细信息。</span><span class="sxs-lookup"><span data-stu-id="fe29e-108">There, you can learn more about all the automation possibilities available to you.</span></span> <span data-ttu-id="fe29e-109">此处的文档重点介绍 Office 脚本与电源自动化的工作方式，以及如何帮助改进 Excel 体验。</span><span class="sxs-lookup"><span data-stu-id="fe29e-109">The documents here focus on how Office Scripts work with Power Automate and how that can help improve your Excel experience.</span></span>

<span data-ttu-id="fe29e-110">若要开始结合使用电源自动化功能和 Office 脚本，请遵循教程[开始使用启用电源自动化的脚本](../tutorials/excel-power-automate-manual.md)。</span><span class="sxs-lookup"><span data-stu-id="fe29e-110">To begin combining Power Automate and Office Scripts, follow the tutorial [Start using scripts with Power Automate](../tutorials/excel-power-automate-manual.md).</span></span> <span data-ttu-id="fe29e-111">这将教您如何创建调用简单脚本的流。</span><span class="sxs-lookup"><span data-stu-id="fe29e-111">This will teach you how to create a flow that calls a simple script.</span></span> <span data-ttu-id="fe29e-112">在完成了教程和将[数据传递到自动运行电源自动化流教程中的脚本](../tutorials/excel-power-automate-trigger.md)之后，请返回此处以了解有关连接 Office 脚本以实现自动处理功能流的详细信息。</span><span class="sxs-lookup"><span data-stu-id="fe29e-112">After you've completed that tutorial and the [Pass data to scripts in an automatically-run Power Automate flow](../tutorials/excel-power-automate-trigger.md) tutorial, return here for detailed information about connecting Office Scripts to Power Automate flows.</span></span>

## <a name="excel-online-business-connector"></a><span data-ttu-id="fe29e-113">Excel Online (商业) 连接器</span><span class="sxs-lookup"><span data-stu-id="fe29e-113">Excel Online (Business) connector</span></span>

<span data-ttu-id="fe29e-114">[连接器](/connectors/connectors)是电源自动化和应用程序之间的桥梁。</span><span class="sxs-lookup"><span data-stu-id="fe29e-114">[Connectors](/connectors/connectors) are the bridges between Power Automate and applications.</span></span> <span data-ttu-id="fe29e-115">[Excel Online (Business) 连接器](/connectors/excelonlinebusiness)为您的流提供对 Excel 工作簿的访问权限。</span><span class="sxs-lookup"><span data-stu-id="fe29e-115">The [Excel Online (Business) connector](/connectors/excelonlinebusiness) gives your flows access to Excel workbooks.</span></span> <span data-ttu-id="fe29e-116">"运行脚本" 操作允许您调用任何可通过所选工作簿访问的 Office 脚本。</span><span class="sxs-lookup"><span data-stu-id="fe29e-116">The "Run script" action lets you call any Office Script accessible through the selected workbook.</span></span> <span data-ttu-id="fe29e-117">您还可以为脚本输入参数，以便流可以提供数据，或让脚本返回有关流中后续步骤的信息。</span><span class="sxs-lookup"><span data-stu-id="fe29e-117">You can also give your scripts input parameters so data can be provided by the flow, or have your script return information for later steps in the flow.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="fe29e-118">"运行脚本" 操作为使用 Excel connector 的用户提供对工作簿及其数据的有效访问权限。</span><span class="sxs-lookup"><span data-stu-id="fe29e-118">The "Run script" action gives people who use the Excel connector significant access to your workbook and its data.</span></span> <span data-ttu-id="fe29e-119">此外，还存在一些使用脚本进行外部 API 调用的安全风险，如[Power 自动化中的外部调用](external-calls.md)中所述。</span><span class="sxs-lookup"><span data-stu-id="fe29e-119">Additionally, there are security risks with scripts that make external API calls, as explained in [External calls from Power Automate](external-calls.md).</span></span> <span data-ttu-id="fe29e-120">如果您的管理员担心暴露高度敏感的数据，则可以关闭 Excel Online 连接器或限制对 Office 脚本的访问，方法是通过[Office 脚本管理员控件](https://support.microsoft.com/office/19d3c51a-6ca2-40ab-978d-60fa49554dcf)。</span><span class="sxs-lookup"><span data-stu-id="fe29e-120">If your admin is concerned with the exposure of highly sensitive data, they can either turn off the Excel Online connector or restrict access to Office Scripts through the [Office Scripts administrator controls](https://support.microsoft.com/office/19d3c51a-6ca2-40ab-978d-60fa49554dcf).</span></span>

## <a name="data-transfer-in-flows-for-scripts"></a><span data-ttu-id="fe29e-121">脚本流中的数据传输</span><span class="sxs-lookup"><span data-stu-id="fe29e-121">Data transfer in flows for scripts</span></span>

<span data-ttu-id="fe29e-122">利用电源自动化，可以在流的各个步骤之间传递数据片段。</span><span class="sxs-lookup"><span data-stu-id="fe29e-122">Power Automate lets you pass pieces of data between steps of your flow.</span></span> <span data-ttu-id="fe29e-123">可以将脚本配置为接受所需的任何类型的信息，并从您的工作簿中返回您想要的任何内容。</span><span class="sxs-lookup"><span data-stu-id="fe29e-123">Scripts can be configured to accept whatever types of information you need and return anything from your workbook that you want in your flow.</span></span> <span data-ttu-id="fe29e-124">脚本的输入通过将参数添加到函数中来指定， `main` (除了 `workbook: ExcelScript.Workbook`) 之外。</span><span class="sxs-lookup"><span data-stu-id="fe29e-124">Input for your script is specified by adding parameters to the `main` function (in addition to `workbook: ExcelScript.Workbook`).</span></span> <span data-ttu-id="fe29e-125">脚本中的输出通过将返回类型添加到来声明 `main` 。</span><span class="sxs-lookup"><span data-stu-id="fe29e-125">Output from the script is declared by adding a return type to `main`.</span></span>

> [!NOTE]
> <span data-ttu-id="fe29e-126">在流中创建 "运行脚本" 块时，将填充接受的参数和返回的类型。</span><span class="sxs-lookup"><span data-stu-id="fe29e-126">When you create a "Run Script" block in your flow, the accepted parameters and returned types are populated.</span></span> <span data-ttu-id="fe29e-127">如果更改了脚本的参数或返回类型，您将需要恢复流的 "运行脚本" 块。</span><span class="sxs-lookup"><span data-stu-id="fe29e-127">If you change the parameters or return types of your script, you'll need to redo the "Run script" block of your flow.</span></span> <span data-ttu-id="fe29e-128">这样可确保正确分析数据。</span><span class="sxs-lookup"><span data-stu-id="fe29e-128">This ensures the data is being parsed correctly.</span></span>

<span data-ttu-id="fe29e-129">以下各节介绍了用于 Power 自动化的脚本输入和输出的详细信息。</span><span class="sxs-lookup"><span data-stu-id="fe29e-129">The following sections cover the details of input and output for scripts used in Power Automate.</span></span> <span data-ttu-id="fe29e-130">如果你想要学习本主题的实践方法，请尝试[在自动运行电源自动化流教程中将数据传递到脚本](../tutorials/excel-power-automate-trigger.md)，或浏览[自动任务提醒](../resources/scenarios/task-reminders.md)示例方案。</span><span class="sxs-lookup"><span data-stu-id="fe29e-130">If you'd like a hands-on approach to learning this topic, try out the [Pass data to scripts in an automatically-run Power Automate flow](../tutorials/excel-power-automate-trigger.md) tutorial or explore the [Automated task reminders](../resources/scenarios/task-reminders.md) sample scenario.</span></span>

### <a name="main-parameters-passing-data-to-a-script"></a><span data-ttu-id="fe29e-131">`main`参数：将数据传递给脚本</span><span class="sxs-lookup"><span data-stu-id="fe29e-131">`main` Parameters: Passing data to a script</span></span>

<span data-ttu-id="fe29e-132">所有脚本输入都被指定为函数的附加参数 `main` 。</span><span class="sxs-lookup"><span data-stu-id="fe29e-132">All script input is specified as additional parameters for the `main` function.</span></span> <span data-ttu-id="fe29e-133">例如，如果您希望脚本接受一个 `string` 表示输入名称的，则会将 `main` 签名更改为 `function main(workbook: ExcelScript.Workbook, name: string)` 。</span><span class="sxs-lookup"><span data-stu-id="fe29e-133">For example, if you wanted a script to accept a `string` that represents a name as input, you would change the `main` signature to `function main(workbook: ExcelScript.Workbook, name: string)`.</span></span>

<span data-ttu-id="fe29e-134">当您在电源自动化中配置流时，您可以将脚本输入指定为静态值、[表达式](/power-automate/use-expressions-in-conditions)或动态内容。</span><span class="sxs-lookup"><span data-stu-id="fe29e-134">When you're configuring a flow in Power Automate, you can specify script input as static values, [expressions](/power-automate/use-expressions-in-conditions), or dynamic content.</span></span> <span data-ttu-id="fe29e-135">有关单个服务连接器的详细信息，请参阅[Power 自动连接器文档](/connectors/)中的。</span><span class="sxs-lookup"><span data-stu-id="fe29e-135">Details on an individual service's connector can be found in the [Power Automate Connector documentation](/connectors/).</span></span>

<span data-ttu-id="fe29e-136">向脚本函数中添加输入参数时 `main` ，请考虑以下余量和限制。</span><span class="sxs-lookup"><span data-stu-id="fe29e-136">When adding input parameters to a script's `main` function, consider the following allowances and restrictions.</span></span>

1. <span data-ttu-id="fe29e-137">第一个参数的类型必须为 `ExcelScript.Workbook` 。</span><span class="sxs-lookup"><span data-stu-id="fe29e-137">The first parameter must be of type `ExcelScript.Workbook`.</span></span> <span data-ttu-id="fe29e-138">其参数名称无关紧要。</span><span class="sxs-lookup"><span data-stu-id="fe29e-138">Its parameter name does not matter.</span></span>

2. <span data-ttu-id="fe29e-139">每个参数都必须具有一个类型 (，如 `string` 或 `number`) 。</span><span class="sxs-lookup"><span data-stu-id="fe29e-139">Every parameter must have a type (such as `string` or `number`).</span></span>

3. <span data-ttu-id="fe29e-140">支持基本类型 `string` 、、、、、 `number` `boolean` `any` `unknown` `object` 和 `undefined` 。</span><span class="sxs-lookup"><span data-stu-id="fe29e-140">The basic types `string`, `number`, `boolean`, `any`, `unknown`, `object`, and `undefined` are supported.</span></span>

4. <span data-ttu-id="fe29e-141">支持前面列出的基本类型的数组。</span><span class="sxs-lookup"><span data-stu-id="fe29e-141">Arrays of the previously listed basic types are supported.</span></span>

5. <span data-ttu-id="fe29e-142">嵌套的数组支持作为参数 (但不支持) 的返回类型。</span><span class="sxs-lookup"><span data-stu-id="fe29e-142">Nested arrays are supported as parameters (but not as return types).</span></span>

6. <span data-ttu-id="fe29e-143">如果联合类型是属于单个类型 (（如) ）的文本联合，则允许联合类型 `"Left" | "Right"` 。</span><span class="sxs-lookup"><span data-stu-id="fe29e-143">Union types are allowed if they are a union of literals belonging to a single type (such as `"Left" | "Right"`).</span></span> <span data-ttu-id="fe29e-144">包含未定义的受支持类型的联合也受支持 (如 `string | undefined`) 。</span><span class="sxs-lookup"><span data-stu-id="fe29e-144">Unions of a supported type with undefined are also supported (such as `string | undefined`).</span></span>

7. <span data-ttu-id="fe29e-145">如果对象类型包含类型 `string` 、 `number` 、、支持的 `boolean` 数组或其他受支持的对象的属性，则允许这些对象类型。</span><span class="sxs-lookup"><span data-stu-id="fe29e-145">Object types are allowed if they contain properties of type `string`, `number`, `boolean`, supported arrays, or other supported objects.</span></span> <span data-ttu-id="fe29e-146">下面的示例演示受支持为参数类型的嵌套对象：</span><span class="sxs-lookup"><span data-stu-id="fe29e-146">The following example shows nested objects that are supported as parameter types:</span></span>

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

8. <span data-ttu-id="fe29e-147">对象必须在脚本中定义其接口或类定义。</span><span class="sxs-lookup"><span data-stu-id="fe29e-147">Objects must have their interface or class definition defined in the script.</span></span> <span data-ttu-id="fe29e-148">也可以以匿名方式直接定义对象，如下面的示例所示：</span><span class="sxs-lookup"><span data-stu-id="fe29e-148">An object can also be defined anonymously inline, as in the following example:</span></span>

    ```TypeScript
    function main(workbook: ExcelScript.Workbook): {name: string, email: string}
    ```

9. <span data-ttu-id="fe29e-149">可选参数是允许的，并且可以使用 optional 修饰符以 `?` (例如， `function main(workbook: ExcelScript.Workbook, Name?: string)`) 。</span><span class="sxs-lookup"><span data-stu-id="fe29e-149">Optional parameters are allowed and can be denoted as such by using the optional modifier `?` (for example, `function main(workbook: ExcelScript.Workbook, Name?: string)`).</span></span>

10. <span data-ttu-id="fe29e-150">例如， (允许使用默认参数值 `async function main(workbook: ExcelScript.Workbook, Name: string = 'Jane Doe')` 。</span><span class="sxs-lookup"><span data-stu-id="fe29e-150">Default parameter values are allowed (for example `async function main(workbook: ExcelScript.Workbook, Name: string = 'Jane Doe')`.</span></span>

### <a name="returning-data-from-a-script"></a><span data-ttu-id="fe29e-151">从脚本中返回数据</span><span class="sxs-lookup"><span data-stu-id="fe29e-151">Returning data from a script</span></span>

<span data-ttu-id="fe29e-152">脚本可以返回工作簿中的数据，以用作电源自动化流中的动态内容。</span><span class="sxs-lookup"><span data-stu-id="fe29e-152">Scripts can return data from the workbook to be used as dynamic content in a Power Automate flow.</span></span> <span data-ttu-id="fe29e-153">与输入参数一样，Power 自动化将一些限制放在返回类型上。</span><span class="sxs-lookup"><span data-stu-id="fe29e-153">As with input parameters, Power Automate places some restrictions on the return type.</span></span>

1. <span data-ttu-id="fe29e-154">支持基本类型 `string` 、 `number` 、 `boolean` `void` 和 `undefined` 。</span><span class="sxs-lookup"><span data-stu-id="fe29e-154">The basic types `string`, `number`, `boolean`, `void`, and `undefined` are supported.</span></span>

2. <span data-ttu-id="fe29e-155">用作返回类型的联合类型遵循与用作脚本参数时相同的限制。</span><span class="sxs-lookup"><span data-stu-id="fe29e-155">Union types used as return types follow the same restrictions as they do when used as script parameters.</span></span>

3. <span data-ttu-id="fe29e-156">如果数组类型为类型 `string` 、或，则允许使用数组类型 `number` `boolean` 。</span><span class="sxs-lookup"><span data-stu-id="fe29e-156">Array types are allowed if they are of type `string`, `number`, or `boolean`.</span></span> <span data-ttu-id="fe29e-157">如果类型是受支持的联合或受支持的文本类型，也可以使用它们。</span><span class="sxs-lookup"><span data-stu-id="fe29e-157">They are also allowed if the type is a supported union or supported literal type.</span></span>

4. <span data-ttu-id="fe29e-158">用作返回类型的对象类型遵循与用作脚本参数时相同的限制。</span><span class="sxs-lookup"><span data-stu-id="fe29e-158">Object types used as return types follow the same restrictions as they do when used as script parameters.</span></span>

5. <span data-ttu-id="fe29e-159">虽然支持隐式键入，但它必须遵循与定义的类型相同的规则。</span><span class="sxs-lookup"><span data-stu-id="fe29e-159">Implicit typing is supported, though it must follow the same rules as a defined type.</span></span>

## <a name="avoid-using-relative-references"></a><span data-ttu-id="fe29e-160">避免使用相对引用</span><span class="sxs-lookup"><span data-stu-id="fe29e-160">Avoid using relative references</span></span>

<span data-ttu-id="fe29e-161">Power 自动在所选的 Excel 工作簿中代表你运行脚本。</span><span class="sxs-lookup"><span data-stu-id="fe29e-161">Power Automate runs your script in the chosen Excel workbook on your behalf.</span></span> <span data-ttu-id="fe29e-162">在这种情况下，工作簿可能会关闭。</span><span class="sxs-lookup"><span data-stu-id="fe29e-162">The workbook might be closed when this happens.</span></span> <span data-ttu-id="fe29e-163">在运行时，任何依赖用户的当前状态（如）的 API `Workbook.getActiveWorksheet` 都将在通过电源自动运行时失败。</span><span class="sxs-lookup"><span data-stu-id="fe29e-163">Any API that relies on the user's current state, such as `Workbook.getActiveWorksheet`, will fail when run through Power Automate.</span></span> <span data-ttu-id="fe29e-164">在设计脚本时，请务必对工作表和区域使用绝对引用。</span><span class="sxs-lookup"><span data-stu-id="fe29e-164">When designing your scripts, be sure to use absolute references for worksheets and ranges.</span></span>

<span data-ttu-id="fe29e-165">如果从 Power 自动流中的脚本调用，以下方法将引发错误并失败。</span><span class="sxs-lookup"><span data-stu-id="fe29e-165">The following methods will throw an error and fail when called from a script in a Power Automate flow.</span></span>

| <span data-ttu-id="fe29e-166">Class</span><span class="sxs-lookup"><span data-stu-id="fe29e-166">Class</span></span> | <span data-ttu-id="fe29e-167">方法</span><span class="sxs-lookup"><span data-stu-id="fe29e-167">Method</span></span> |
|--|--|
| [<span data-ttu-id="fe29e-168">Chart</span><span class="sxs-lookup"><span data-stu-id="fe29e-168">Chart</span></span>](/javascript/api/office-scripts/excelscript/excelscript.chart) | `activate` |
| [<span data-ttu-id="fe29e-169">区域</span><span class="sxs-lookup"><span data-stu-id="fe29e-169">Range</span></span>](/javascript/api/office-scripts/excelscript/excelscript.range) | `select` |
| [<span data-ttu-id="fe29e-170">Workbook</span><span class="sxs-lookup"><span data-stu-id="fe29e-170">Workbook</span></span>](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getActiveCell` |
| [<span data-ttu-id="fe29e-171">Workbook</span><span class="sxs-lookup"><span data-stu-id="fe29e-171">Workbook</span></span>](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getActiveChart` |
| [<span data-ttu-id="fe29e-172">Workbook</span><span class="sxs-lookup"><span data-stu-id="fe29e-172">Workbook</span></span>](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getActiveSlicer` |
| [<span data-ttu-id="fe29e-173">Workbook</span><span class="sxs-lookup"><span data-stu-id="fe29e-173">Workbook</span></span>](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getActiveWorksheet` |
| [<span data-ttu-id="fe29e-174">Workbook</span><span class="sxs-lookup"><span data-stu-id="fe29e-174">Workbook</span></span>](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getSelectedRange` |
| [<span data-ttu-id="fe29e-175">Workbook</span><span class="sxs-lookup"><span data-stu-id="fe29e-175">Workbook</span></span>](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getSelectedRanges` |
| [<span data-ttu-id="fe29e-176">Worksheet</span><span class="sxs-lookup"><span data-stu-id="fe29e-176">Worksheet</span></span>](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `activate` |

## <a name="example"></a><span data-ttu-id="fe29e-177">示例</span><span class="sxs-lookup"><span data-stu-id="fe29e-177">Example</span></span>

<span data-ttu-id="fe29e-178">下面的屏幕截图显示了只要向您分配[GitHub](https://github.com/)问题时触发的电源自动化流。</span><span class="sxs-lookup"><span data-stu-id="fe29e-178">The following screenshot shows a Power Automate flow that's triggered whenever a [GitHub](https://github.com/) issue is assigned to you.</span></span> <span data-ttu-id="fe29e-179">流运行一个将问题添加到 Excel 工作簿中的表的脚本。</span><span class="sxs-lookup"><span data-stu-id="fe29e-179">The flow runs a script that adds the issue to a table in an Excel workbook.</span></span> <span data-ttu-id="fe29e-180">如果该表中有五个或更多问题，流将发送电子邮件提醒。</span><span class="sxs-lookup"><span data-stu-id="fe29e-180">If there are five or more issues in that table, the flow sends an email reminder.</span></span>

![示例流，如 Power 自动化流编辑器中所示。](../images/power-automate-parameter-return-sample.png)

<span data-ttu-id="fe29e-182">`main`脚本的功能将问题 ID 和问题标题指定为输入参数，脚本将返回 "问题" 表中的行数。</span><span class="sxs-lookup"><span data-stu-id="fe29e-182">The `main` function of the script specifies the issue ID and issue title as input parameters, and the script returns the number of rows in the issue table.</span></span>

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

## <a name="see-also"></a><span data-ttu-id="fe29e-183">另请参阅</span><span class="sxs-lookup"><span data-stu-id="fe29e-183">See also</span></span>

- [<span data-ttu-id="fe29e-184">在使用 Power 自动化的 web 上运行 Excel 中的 Office 脚本</span><span class="sxs-lookup"><span data-stu-id="fe29e-184">Run Office Scripts in Excel on the web with Power Automate</span></span>](../tutorials/excel-power-automate-manual.md)
- [<span data-ttu-id="fe29e-185">将数据传递到自动运行的 Power Automate 流中的脚本</span><span class="sxs-lookup"><span data-stu-id="fe29e-185">Pass data to scripts in an automatically-run Power Automate flow</span></span>](../tutorials/excel-power-automate-trigger.md)
- [<span data-ttu-id="fe29e-186">Excel 网页版中 Office 脚本的脚本基础知识</span><span class="sxs-lookup"><span data-stu-id="fe29e-186">Scripting fundamentals for Office Scripts in Excel on the web</span></span>](scripting-fundamentals.md)
- [<span data-ttu-id="fe29e-187">Power Automate 入门</span><span class="sxs-lookup"><span data-stu-id="fe29e-187">Get started with Power Automate</span></span>](/power-automate/getting-started)
- [<span data-ttu-id="fe29e-188">Excel Online (Business) 连接器参考文档</span><span class="sxs-lookup"><span data-stu-id="fe29e-188">Excel Online (Business) connector reference documentation</span></span>](/connectors/excelonlinebusiness/)
