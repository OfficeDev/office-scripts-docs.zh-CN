---
title: Office 脚本入门
description: 有关 Office 脚本的基础知识，包括访问、环境和脚本模式。
ms.date: 04/01/2021
localization_priority: Normal
ms.openlocfilehash: f954ee67aa486e4b8185047738ef3d15319a94ae
ms.sourcegitcommit: 5d24e77df70aa2c1c982275d53213c2a9323ff86
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 04/02/2021
ms.locfileid: "51571112"
---
# <a name="getting-started"></a><span data-ttu-id="8b48b-103">入门</span><span class="sxs-lookup"><span data-stu-id="8b48b-103">Getting started</span></span>

<span data-ttu-id="8b48b-104">本节提供有关 Office 脚本的基础知识的详细信息，包括访问、环境、脚本基础知识和几个基本脚本模式。</span><span class="sxs-lookup"><span data-stu-id="8b48b-104">This section provides details about the basics of Office Scripts including access, environment, script fundamentals, and few basic script patterns.</span></span>

## <a name="environment-setup"></a><span data-ttu-id="8b48b-105">环境设置</span><span class="sxs-lookup"><span data-stu-id="8b48b-105">Environment setup</span></span>

<span data-ttu-id="8b48b-106">了解访问、环境和脚本编辑器的基础知识。</span><span class="sxs-lookup"><span data-stu-id="8b48b-106">Learn about the basics of access, environment, and script editor.</span></span>

<span data-ttu-id="8b48b-107">[![Office 脚本应用程序的基础知识](../../images/getting-started-env.png)](https://youtu.be/vvCtxsjPxo8 "Office 脚本应用程序的基础知识")</span><span class="sxs-lookup"><span data-stu-id="8b48b-107">[![Basics of Office Scripts application](../../images/getting-started-env.png)](https://youtu.be/vvCtxsjPxo8 "Basics of Office Scripts application")</span></span>

### <a name="access"></a><span data-ttu-id="8b48b-108">Access</span><span class="sxs-lookup"><span data-stu-id="8b48b-108">Access</span></span>

<span data-ttu-id="8b48b-109">Office Scripts requires admin settings available for Microsoft 365 administrator under **Settings**  >  **Org settings** Office  >  **Scripts**.</span><span class="sxs-lookup"><span data-stu-id="8b48b-109">Office Scripts requires admin settings available for Microsoft 365 administrator under **Settings** > **Org settings** > **Office Scripts**.</span></span> <span data-ttu-id="8b48b-110">默认情况下，会为所有用户打开它。</span><span class="sxs-lookup"><span data-stu-id="8b48b-110">By default, it's turned on for all users.</span></span> <span data-ttu-id="8b48b-111">有两个子设置，管理员可以打开和关闭它们。</span><span class="sxs-lookup"><span data-stu-id="8b48b-111">There are two sub-settings, which the admin can turn on and off.</span></span>

* <span data-ttu-id="8b48b-112">在组织内部共享脚本的能力</span><span class="sxs-lookup"><span data-stu-id="8b48b-112">Ability to share scripts within the organization</span></span>
* <span data-ttu-id="8b48b-113">在 Power Automate 中使用脚本的能力</span><span class="sxs-lookup"><span data-stu-id="8b48b-113">Ability to use scripts in Power Automate</span></span>

<span data-ttu-id="8b48b-114">您可以通过在 Excel 网页浏览器 (浏览器) 中打开文件并查看 Excel 功能区中是否显示"自动"选项卡，来判断您是否可以访问 Office 脚本。</span><span class="sxs-lookup"><span data-stu-id="8b48b-114">You can tell if you have access to Office Scripts by opening a file in Excel on the web (browser) and seeing if the **Automate** tab appears in the Excel ribbon or not.</span></span>
<span data-ttu-id="8b48b-115">如果仍然看不到"自动执行 **"选项卡，** 请查看 [此疑难解答部分](../../testing/troubleshooting.md#automate-tab-not-appearing-or-office-scripts-unavailable)。</span><span class="sxs-lookup"><span data-stu-id="8b48b-115">If you still can't see the **Automate** tab, check [this troubleshooting section](../../testing/troubleshooting.md#automate-tab-not-appearing-or-office-scripts-unavailable).</span></span>

### <a name="availability"></a><span data-ttu-id="8b48b-116">供应情况</span><span class="sxs-lookup"><span data-stu-id="8b48b-116">Availability</span></span>

<span data-ttu-id="8b48b-117">Office 脚本仅适用于 Excel 网页版中的企业版 E3+ 许可证 (用户和 E1 帐户不支持) 。</span><span class="sxs-lookup"><span data-stu-id="8b48b-117">Office Scripts is available only in the Excel on the web for Enterprise E3+ licenses (Consumer and E1 accounts are not supported).</span></span> <span data-ttu-id="8b48b-118">Windows 和 Mac 上的 Excel 尚不支持 Office 脚本。</span><span class="sxs-lookup"><span data-stu-id="8b48b-118">Office Scripts is not yet supported in Excel on Windows and Mac.</span></span>

### <a name="scripts-and-editor"></a><span data-ttu-id="8b48b-119">脚本和编辑程序</span><span class="sxs-lookup"><span data-stu-id="8b48b-119">Scripts and editor</span></span>

<span data-ttu-id="8b48b-120">代码编辑器内置于 Excel 网页版 (联机) 。</span><span class="sxs-lookup"><span data-stu-id="8b48b-120">The code editor is built right into Excel on the web (online version).</span></span> <span data-ttu-id="8b48b-121">如果你已使用编辑器（如Visual Studio Code 或 Sublime），则此编辑体验将非常相似。</span><span class="sxs-lookup"><span data-stu-id="8b48b-121">If you have used editors like Visual Studio Code or Sublime, this editing experience will be quite similar.</span></span>
<span data-ttu-id="8b48b-122">代码编辑器使用Visual Studio大多数快捷键在 Office 脚本编辑体验中也工作。</span><span class="sxs-lookup"><span data-stu-id="8b48b-122">Most of the shortcut keys that Visual Studio Code editor uses work in the Office Scripts editing experience as well.</span></span> <span data-ttu-id="8b48b-123">请查看以下快捷键讲义。</span><span class="sxs-lookup"><span data-stu-id="8b48b-123">Check out the following shortcut keys handouts.</span></span>

* [<span data-ttu-id="8b48b-124">macOS</span><span class="sxs-lookup"><span data-stu-id="8b48b-124">macOS</span></span>](https://code.visualstudio.com/shortcuts/keyboard-shortcuts-macos.pdf)
* [<span data-ttu-id="8b48b-125">Windows</span><span class="sxs-lookup"><span data-stu-id="8b48b-125">Windows</span></span>](https://code.visualstudio.com/shortcuts/keyboard-shortcuts-windows.pdf)

#### <a name="key-things-to-note"></a><span data-ttu-id="8b48b-126">要注意的关键内容</span><span class="sxs-lookup"><span data-stu-id="8b48b-126">Key things to note</span></span>

* <span data-ttu-id="8b48b-127">Office 脚本仅适用于存储在 OneDrive for Business、SharePoint 网站和团队网站中的文件。</span><span class="sxs-lookup"><span data-stu-id="8b48b-127">Office Scripts is only available for files stored in OneDrive for Business, SharePoint sites, and Team sites.</span></span>
* <span data-ttu-id="8b48b-128">编辑器不会显示脚本的扩展名。</span><span class="sxs-lookup"><span data-stu-id="8b48b-128">The editor doesn't show the script's extension.</span></span> <span data-ttu-id="8b48b-129">实际上，这些是 TypeScript 文件，但它们使用名为 的自定义扩展存储 `.osts` 。</span><span class="sxs-lookup"><span data-stu-id="8b48b-129">In reality, these are TypeScript files but they are stored with a custom extension called `.osts`.</span></span>
* <span data-ttu-id="8b48b-130">脚本存储在你自己的 OneDrive for Business 文件夹中 `My Files/Documents/OfficeScripts` 。</span><span class="sxs-lookup"><span data-stu-id="8b48b-130">The scripts are stored in your own OneDrive for Business folder `My Files/Documents/OfficeScripts`.</span></span> <span data-ttu-id="8b48b-131">无需管理此文件夹。</span><span class="sxs-lookup"><span data-stu-id="8b48b-131">You won't need to manage this folder.</span></span> <span data-ttu-id="8b48b-132">对于部件，你可以忽略这一方面，因为编辑器管理查看/编辑体验。</span><span class="sxs-lookup"><span data-stu-id="8b48b-132">For your part, you can ignore this aspect as the editor manages the viewing/editing experience.</span></span>
* <span data-ttu-id="8b48b-133">脚本不会存储为 Excel 文件的一部分。</span><span class="sxs-lookup"><span data-stu-id="8b48b-133">Scripts are not stored as part of Excel files.</span></span> <span data-ttu-id="8b48b-134">它们单独存储。</span><span class="sxs-lookup"><span data-stu-id="8b48b-134">They are stored separately.</span></span>
* <span data-ttu-id="8b48b-135">你可以与 Excel 文件共享脚本，这实际上意味着你将脚本与文件链接，而不是附加它。</span><span class="sxs-lookup"><span data-stu-id="8b48b-135">You can share the script with an Excel file which in effect means you are linking the script with the file, not attaching it.</span></span> <span data-ttu-id="8b48b-136">有权访问 Excel 文件的任何人也能够查看、运行或 **制作脚本** 副本。 </span><span class="sxs-lookup"><span data-stu-id="8b48b-136">Whoever has access to the Excel file will also be able to **view**, **run**, or **make a copy** of the script.</span></span> <span data-ttu-id="8b48b-137">与 VBA 宏相比，这是一个关键区别。</span><span class="sxs-lookup"><span data-stu-id="8b48b-137">This is a key difference compared to VBA macros.</span></span>
* <span data-ttu-id="8b48b-138">除非你共享脚本，否则其他人无法访问它，因为它驻留在你自己的库中。</span><span class="sxs-lookup"><span data-stu-id="8b48b-138">Unless you share your scripts, no one else can access it as it resides in your own library.</span></span>
* <span data-ttu-id="8b48b-139">无法从本地磁盘或自定义云位置链接脚本。</span><span class="sxs-lookup"><span data-stu-id="8b48b-139">Scripts can't be linked from a local disk or custom cloud locations.</span></span> <span data-ttu-id="8b48b-140">Office 脚本仅识别并运行上述 OneDrive 文件夹的预定义 (或共享脚本) 脚本。</span><span class="sxs-lookup"><span data-stu-id="8b48b-140">Office Scripts only recognizes and runs a script that is on predefined location (your OneDrive folder mentioned above) or shared scripts.</span></span>
* <span data-ttu-id="8b48b-141">在编辑过程中，文件会临时保存在浏览器中，但在关闭 Excel 窗口将其保存到 OneDrive 位置之前，必须保存脚本。</span><span class="sxs-lookup"><span data-stu-id="8b48b-141">During editing, files are temporarily saved in the browser but you'll have to save the script before closing the Excel window to save it to the OneDrive location.</span></span> <span data-ttu-id="8b48b-142">请不要忘记在编辑后保存文件。</span><span class="sxs-lookup"><span data-stu-id="8b48b-142">Don't forget to save the file after edits.</span></span>

## <a name="gentle-introduction-to-scripting"></a><span data-ttu-id="8b48b-143">脚本简介简介</span><span class="sxs-lookup"><span data-stu-id="8b48b-143">Gentle introduction to scripting</span></span>

<span data-ttu-id="8b48b-144">Office 脚本是使用 TypeScript 语言编写的独立脚本，其中包含对选定的 Excel 工作簿执行一些自动化操作的说明。</span><span class="sxs-lookup"><span data-stu-id="8b48b-144">Office Scripts are standalone scripts written in the TypeScript language that contain instructions to perform some automation against the selected Excel workbook.</span></span> <span data-ttu-id="8b48b-145">所有自动化指令都自包含在脚本中，脚本无法调用或调用其他脚本。</span><span class="sxs-lookup"><span data-stu-id="8b48b-145">All automation instructions are self-contained within a script and scripts can't invoke or call other scripts.</span></span> <span data-ttu-id="8b48b-146">所有脚本都存储在独立文件中，并存储在用户的 OneDrive 文件夹中。</span><span class="sxs-lookup"><span data-stu-id="8b48b-146">All scripts are stored in standalone files and stored on the user's OneDrive folder.</span></span> <span data-ttu-id="8b48b-147">可以录制新脚本、编辑录制的脚本或从头开始编写全新的脚本，所有这些都在内置编辑器界面中完成。</span><span class="sxs-lookup"><span data-stu-id="8b48b-147">You can record a new script, edit a recorded script, or write a whole new script from scratch, all within a built-in editor interface.</span></span> <span data-ttu-id="8b48b-148">Office 脚本的最好的一部分是，它们不需要用户进一步设置。</span><span class="sxs-lookup"><span data-stu-id="8b48b-148">The best part of Office Scripts is that they don't need any further setup from users.</span></span> <span data-ttu-id="8b48b-149">没有外部库、网页或 UI 元素、设置等。所有环境设置都由 Office 脚本处理，它允许通过简单的 API 界面轻松而快速地访问自动化。</span><span class="sxs-lookup"><span data-stu-id="8b48b-149">No external libraries, web pages, or UI elements, setup, etc. All the environment setup is handled by Office Scripts and it allows easy and fast access to automation through a simple API interface.</span></span>

<span data-ttu-id="8b48b-150">一些有助于了解如何编辑和浏览脚本的基本概念包括：</span><span class="sxs-lookup"><span data-stu-id="8b48b-150">Some of the basic concepts helpful to understand how to edit and navigate around scripts include:</span></span>

* <span data-ttu-id="8b48b-151">基本 TypeScript 语言语法</span><span class="sxs-lookup"><span data-stu-id="8b48b-151">Basic TypeScript language syntax</span></span>
* <span data-ttu-id="8b48b-152">了解 `main` 函数和参数</span><span class="sxs-lookup"><span data-stu-id="8b48b-152">Understanding of `main` function and arguments</span></span>
* <span data-ttu-id="8b48b-153">对象和层次结构、方法、属性</span><span class="sxs-lookup"><span data-stu-id="8b48b-153">Objects and hierarchy, methods, properties</span></span>
* <span data-ttu-id="8b48b-154">数组 (集合) 导航和操作</span><span class="sxs-lookup"><span data-stu-id="8b48b-154">Collection (array): navigation and operations</span></span>
* <span data-ttu-id="8b48b-155">类型定义</span><span class="sxs-lookup"><span data-stu-id="8b48b-155">Type definitions</span></span>
* <span data-ttu-id="8b48b-156">环境：记录/编辑、运行、检查结果、共享</span><span class="sxs-lookup"><span data-stu-id="8b48b-156">Environment: record/edit, run, examine results, share</span></span>

<span data-ttu-id="8b48b-157">本视频和部分详细介绍了其中一些概念。</span><span class="sxs-lookup"><span data-stu-id="8b48b-157">This video and section explain some of these concepts in detail.</span></span>

<span data-ttu-id="8b48b-158">[![Office 脚本基础知识](../../images/getting-started-v_script.png)](https://youtu.be/8Zsrc1uaiiU "脚本的基础知识")</span><span class="sxs-lookup"><span data-stu-id="8b48b-158">[![Basics of Office Scripts](../../images/getting-started-v_script.png)](https://youtu.be/8Zsrc1uaiiU "Basics of Scripts")</span></span>

### <a name="language-typescript"></a><span data-ttu-id="8b48b-159">语言：TypeScript</span><span class="sxs-lookup"><span data-stu-id="8b48b-159">Language: TypeScript</span></span>

<span data-ttu-id="8b48b-160">[Office 脚本](../../index.md) 是使用 [TypeScript](https://www.typescriptlang.org/)语言编写的，该语言是一种基于 JavaScript (之一的开放源代码语言，它通过添加静态类型定义) 最常用的语言之一。</span><span class="sxs-lookup"><span data-stu-id="8b48b-160">[Office Scripts](../../index.md) is written using the [TypeScript language](https://www.typescriptlang.org/), which is an open-source language that builds on JavaScript (one of the world's most used) by adding static type definitions.</span></span> <span data-ttu-id="8b48b-161">正如网站所说明的，提供一种方法来描述对象的形状，提供更好的文档，并允许 `Types` TypeScript 验证代码是否正常工作。</span><span class="sxs-lookup"><span data-stu-id="8b48b-161">As the website says, `Types` provide a way to describe the shape of an object, providing better documentation, and allowing TypeScript to validate that your code is working correctly.</span></span>

<span data-ttu-id="8b48b-162">语言语法本身使用 [JavaScript](https://developer.mozilla.org/docs/Web/JavaScript) 编写，其他键入使用 TypeScript 约定在脚本中定义。</span><span class="sxs-lookup"><span data-stu-id="8b48b-162">The language syntax itself is written using [JavaScript](https://developer.mozilla.org/docs/Web/JavaScript) with additional typings defined in the script using TypeScript conventions.</span></span> <span data-ttu-id="8b48b-163">在大多数情况下，您可以将 Office 脚本视为使用 JavaScript 编写的。</span><span class="sxs-lookup"><span data-stu-id="8b48b-163">For the most part, you can think of Office Scripts as written in JavaScript.</span></span> <span data-ttu-id="8b48b-164">必须了解 JavaScript 语言的基础知识，以开始 Office 脚本之旅;尽管你无需精通它，但可以开始你的自动化之旅。</span><span class="sxs-lookup"><span data-stu-id="8b48b-164">It is essential that you understand the basics of JavaScript language to begin your Office Scripts journey; though you don't need to be proficient at it to begin your automation journey.</span></span> <span data-ttu-id="8b48b-165">使用 Office 脚本的操作录制器，您可以了解脚本语句，因为包含代码注释，您可以遵循和进行小型编辑。</span><span class="sxs-lookup"><span data-stu-id="8b48b-165">With the Office Scripts' action recorder, you can understand the script statements because code comments are included and you can follow along and make small edits.</span></span>

<span data-ttu-id="8b48b-166">允许脚本与 Excel 交互的 Office 脚本 API 是为可能没有太多编码背景的最终用户设计的。</span><span class="sxs-lookup"><span data-stu-id="8b48b-166">Office Scripts APIs, which allow the script to interact with Excel, are designed for end-users who may not have much coding background.</span></span> <span data-ttu-id="8b48b-167">API 可以同步调用，你无需了解高级主题，如承诺或回调。</span><span class="sxs-lookup"><span data-stu-id="8b48b-167">APIs can be invoked synchronously and you don't need to know advanced topics such as promises or callbacks.</span></span> <span data-ttu-id="8b48b-168">Office 脚本 API 设计提供：</span><span class="sxs-lookup"><span data-stu-id="8b48b-168">Office Scripts API design provides:</span></span>

* <span data-ttu-id="8b48b-169">包含方法、getters/setters 的简单对象模型。</span><span class="sxs-lookup"><span data-stu-id="8b48b-169">Simple object model with methods, getters/setters.</span></span>
* <span data-ttu-id="8b48b-170">作为常规数组的易于访问的对象集合。</span><span class="sxs-lookup"><span data-stu-id="8b48b-170">Easy-to-access object collections as regular arrays.</span></span>
* <span data-ttu-id="8b48b-171">简单的错误处理选项。</span><span class="sxs-lookup"><span data-stu-id="8b48b-171">Simple error handling options.</span></span>
* <span data-ttu-id="8b48b-172">优化了选定方案的性能，帮助用户专注于当前方案。</span><span class="sxs-lookup"><span data-stu-id="8b48b-172">Optimized performance for select scenarios helping users to focus on the scenario at hand.</span></span>

### <a name="main-function-the-scripts-starting-point"></a><span data-ttu-id="8b48b-173">`main` function：脚本的起始点</span><span class="sxs-lookup"><span data-stu-id="8b48b-173">`main` function: The script's starting point</span></span>

<span data-ttu-id="8b48b-174">Office 脚本的执行从 函数 `main` 开始。</span><span class="sxs-lookup"><span data-stu-id="8b48b-174">Office Scripts' execution begins at the `main` function.</span></span> <span data-ttu-id="8b48b-175">脚本是包含一个或多个函数以及类型、接口、变量等声明的单个文件。若要随脚本一起操作，请从 函数开始，因为 Excel 始终在您执行任何脚本时首先 `main` `main` 调用 函数。</span><span class="sxs-lookup"><span data-stu-id="8b48b-175">A script is a single file containing one or many functions along with declarations of types, interfaces, variables, etc. To follow along with the script, begin with the `main` function as Excel always first invokes the `main` function when you execute any script.</span></span> <span data-ttu-id="8b48b-176">函数将始终具有至少一个名为 (参数或) 参数，该参数只是一个标识脚本所针对的当前工作簿的 `main` `workbook` 变量名称。</span><span class="sxs-lookup"><span data-stu-id="8b48b-176">The `main` function will always have at least one argument (or parameter) named `workbook`, which is just a variable name identifying the current workbook against which the script is running.</span></span> <span data-ttu-id="8b48b-177">你可以定义其他参数，以使用 Power Automate (脱机) 执行。</span><span class="sxs-lookup"><span data-stu-id="8b48b-177">You can define additional arguments for usage with Power Automate (offline) execution.</span></span>

* `function main(workbook: ExcelScript.Workbook)`

<span data-ttu-id="8b48b-178">可以将脚本组织为较小的函数，帮助实现代码的可重复性、清晰度等。其他函数可以位于主函数内部或外部，但始终位于同一文件中。</span><span class="sxs-lookup"><span data-stu-id="8b48b-178">A script can be organized into smaller functions to aid with code reusability, clarity, etc. Other functions can be inside or outside of the main function but always in the same file.</span></span> <span data-ttu-id="8b48b-179">脚本是自包含的，只能使用在同一文件中定义的函数。</span><span class="sxs-lookup"><span data-stu-id="8b48b-179">A script is self-contained and can only use functions defined in the same file.</span></span> <span data-ttu-id="8b48b-180">脚本无法调用或调用其他 Office 脚本。</span><span class="sxs-lookup"><span data-stu-id="8b48b-180">Scripts cannot invoke or call another Office Script.</span></span>

<span data-ttu-id="8b48b-181">因此，总之：</span><span class="sxs-lookup"><span data-stu-id="8b48b-181">So, in summary:</span></span>

* <span data-ttu-id="8b48b-182">`main`函数是任何脚本的入口点。</span><span class="sxs-lookup"><span data-stu-id="8b48b-182">The `main` function is the entry point for any script.</span></span> <span data-ttu-id="8b48b-183">执行函数时，Excel 应用程序通过提供工作簿作为其第一个参数来调用此主函数。</span><span class="sxs-lookup"><span data-stu-id="8b48b-183">When the function is executed, the Excel application invokes this main function by providing the workbook as its first parameter.</span></span>
* <span data-ttu-id="8b48b-184">在显示时保留第一个参数 `workbook` 及其类型声明很重要。</span><span class="sxs-lookup"><span data-stu-id="8b48b-184">It's important to keep the first argument `workbook` and its type declaration as it appears.</span></span> <span data-ttu-id="8b48b-185">你可以向函数添加新参数 (请参阅下一节) 但第一个参数保持 `main` 为正常。</span><span class="sxs-lookup"><span data-stu-id="8b48b-185">You can add new arguments to the `main` function (see the next section) but do keep the first argument as is.</span></span>

![主函数是脚本的入口点](../../images/getting-started-main-introduction.png)

#### <a name="send-or-receive-data-from-other-apps"></a><span data-ttu-id="8b48b-187">发送或接收来自其他应用的数据</span><span class="sxs-lookup"><span data-stu-id="8b48b-187">Send or receive data from other apps</span></span>

<span data-ttu-id="8b48b-188">可以通过在 Power Automate 中运行脚本将 Excel 连接到 [组织的其他部分](https://flow.microsoft.com)。</span><span class="sxs-lookup"><span data-stu-id="8b48b-188">You can connect Excel to other parts of your organization by running scripts in [Power Automate](https://flow.microsoft.com).</span></span> <span data-ttu-id="8b48b-189">详细了解在 [Power Automate 流中运行 Office 脚本](../../develop/power-automate-integration.md)。</span><span class="sxs-lookup"><span data-stu-id="8b48b-189">Learn more about [running Office Scripts in Power Automate flows](../../develop/power-automate-integration.md).</span></span>

<span data-ttu-id="8b48b-190">从 Excel 接收数据或将数据发送到 Excel 的方式是通过 `main` 函数。</span><span class="sxs-lookup"><span data-stu-id="8b48b-190">The way to receive or send data from and to Excel is through the `main` function.</span></span> <span data-ttu-id="8b48b-191">将它视为信息网关，允许在脚本中描述和使用传入和传出数据。</span><span class="sxs-lookup"><span data-stu-id="8b48b-191">Think of it as the information gateway that allows incoming and outgoing data to be described and used in the script.</span></span> <span data-ttu-id="8b48b-192">可以使用 数据类型 从脚本外部接收数据，并返回任何 TypeScript 识别的数据（如 、 、 或 在脚本中定义的接口形式的任何 `string` `string` `number` `boolean` 对象）。</span><span class="sxs-lookup"><span data-stu-id="8b48b-192">You can receive data from outside the script using the `string` data type and return any TypeScript-recognized data such as `string`, `number`, `boolean`, or any objects in the form of interfaces you define in the script.</span></span>

![脚本的输入和输出](../../images/getting-started-data-in-out.png)

#### <a name="use-functions-to-organize-and-reuse-code"></a><span data-ttu-id="8b48b-194">使用函数组织和重复使用代码</span><span class="sxs-lookup"><span data-stu-id="8b48b-194">Use functions to organize and reuse code</span></span>

<span data-ttu-id="8b48b-195">可以使用函数在脚本中组织和重复使用代码。</span><span class="sxs-lookup"><span data-stu-id="8b48b-195">You can use functions to organize and reuse code within your script.</span></span>

![在脚本中使用函数](../../images/getting-started-use-functions.png)

### <a name="objects-hierarchy-methods-properties-collections"></a><span data-ttu-id="8b48b-197">对象、层次结构、方法、属性、集合</span><span class="sxs-lookup"><span data-stu-id="8b48b-197">Objects, hierarchy, methods, properties, collections</span></span>

<span data-ttu-id="8b48b-198">Excel 的所有对象模型在对象的层次结构中定义，从类型 为 的 workbook 对象开始 `ExcelScript.Workbook` 。</span><span class="sxs-lookup"><span data-stu-id="8b48b-198">All of Excel's object model is defined in a hierarchical structure of objects, beginning with the workbook object of type `ExcelScript.Workbook`.</span></span> <span data-ttu-id="8b48b-199">对象可以包含方法、属性和其他对象。</span><span class="sxs-lookup"><span data-stu-id="8b48b-199">An object can contain methods, properties, and other objects within it.</span></span> <span data-ttu-id="8b48b-200">对象使用 方法相互链接。</span><span class="sxs-lookup"><span data-stu-id="8b48b-200">Objects are linked to each other using the methods.</span></span> <span data-ttu-id="8b48b-201">对象的方法可以返回另一个对象或对象集合。</span><span class="sxs-lookup"><span data-stu-id="8b48b-201">An object's method can return another object or collection of objects.</span></span> <span data-ttu-id="8b48b-202">使用代码编辑器的 IntelliSense (代码) 功能是浏览对象层次结构的一种很好的方法。</span><span class="sxs-lookup"><span data-stu-id="8b48b-202">Using the code editor's IntelliSense (code completion) feature is a great way to explore the object hierarchy.</span></span> <span data-ttu-id="8b48b-203">您还可以使用官方 [参考文档网站](/javascript/api/office-scripts/overview) 来跟踪对象之间的关系。</span><span class="sxs-lookup"><span data-stu-id="8b48b-203">You can also use the [official reference documentation site](/javascript/api/office-scripts/overview) to follow along with the relationships among objects.</span></span>

<span data-ttu-id="8b48b-204">对象 [是](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Object) 一组属性，而属性是名称或键 (值) 之间的关联。</span><span class="sxs-lookup"><span data-stu-id="8b48b-204">An [object](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Object) is a collection of properties, and a property is an association between a name (or key) and a value.</span></span> <span data-ttu-id="8b48b-205">属性的值可以是一个函数，在这种情况下，该属性称为方法。</span><span class="sxs-lookup"><span data-stu-id="8b48b-205">A property's value can be a function, in which case the property is known as a method.</span></span> <span data-ttu-id="8b48b-206">对于 Office 脚本对象模型，对象表示 Excel 文件中用户与之交互的内容，如图表、超链接、数据透视表等。它还可以表示对象的行为，如工作表的保护属性。</span><span class="sxs-lookup"><span data-stu-id="8b48b-206">In the case of the Office Scripts object model, an object represents a thing in the Excel file that users interact with such as a chart, hyperlink, pivot-table, etc. It can also represent the behavior of an object such as the protection attributes of a worksheet.</span></span>

<span data-ttu-id="8b48b-207">TypeScript 对象和属性与方法的主题相当深入。</span><span class="sxs-lookup"><span data-stu-id="8b48b-207">The topic of TypeScript objects and properties vs methods is quite deep.</span></span> <span data-ttu-id="8b48b-208">为了开始使用脚本并提高工作效率，你可以记住一些基本内容：</span><span class="sxs-lookup"><span data-stu-id="8b48b-208">In order to get started with the script and be productive, you can remember a few basic things:</span></span>

* <span data-ttu-id="8b48b-209">对象和属性均使用点 (点) 表示法访问，对象位于 的左侧，属性或方法位于 `.` `.` 右侧。</span><span class="sxs-lookup"><span data-stu-id="8b48b-209">Both objects and properties are accessed using `.` (dot) notation, with the object on the left side of the `.` and the property or method on the right side.</span></span> <span data-ttu-id="8b48b-210">示例 `hyperlink.address` `range.getAddress()` ：、。</span><span class="sxs-lookup"><span data-stu-id="8b48b-210">Examples: `hyperlink.address`, `range.getAddress()`.</span></span>
* <span data-ttu-id="8b48b-211">属性在本质上是标量 (字符串、布尔值、数字) 。</span><span class="sxs-lookup"><span data-stu-id="8b48b-211">Properties are scalar in nature (strings, booleans, numbers).</span></span> <span data-ttu-id="8b48b-212">例如，工作簿的名称、工作表的位置、表格是否具有页脚的值。</span><span class="sxs-lookup"><span data-stu-id="8b48b-212">For example, name of a workbook, position of a worksheet, the value of whether the table has a footer or not.</span></span>
* <span data-ttu-id="8b48b-213">方法使用开放关闭括号"调用"或"执行"。</span><span class="sxs-lookup"><span data-stu-id="8b48b-213">Methods are 'invoked' or 'executed' using the open-close parentheses.</span></span> <span data-ttu-id="8b48b-214">示例：`table.delete()`。</span><span class="sxs-lookup"><span data-stu-id="8b48b-214">Example: `table.delete()`.</span></span> <span data-ttu-id="8b48b-215">有时，参数在打开关闭的括号之间包含，以传递给函数 `range.setValue('Hello')` ：。</span><span class="sxs-lookup"><span data-stu-id="8b48b-215">Sometimes an argument is passed to a function by including them between open-close parentheses: `range.setValue('Hello')`.</span></span> <span data-ttu-id="8b48b-216">可以将许多参数传递给函数 (其协定/签名参数定义) 使用 分隔 `,` 。</span><span class="sxs-lookup"><span data-stu-id="8b48b-216">You can pass many arguments to a function (as defined by its contract/signature) and separate them using `,`.</span></span>  <span data-ttu-id="8b48b-217">例如：`worksheet.addTable('A1:D6', true)`。</span><span class="sxs-lookup"><span data-stu-id="8b48b-217">For example: `worksheet.addTable('A1:D6', true)`.</span></span> <span data-ttu-id="8b48b-218">您可以传递方法所需的任何类型的参数，如字符串、数字、布尔值，甚至是其他对象，例如 ，其中 是脚本中其他位置创建 `worksheet.addTable(targetRange, true)` `targetRange` 的对象。</span><span class="sxs-lookup"><span data-stu-id="8b48b-218">You can pass arguments of any type as required by the method such as strings, number, boolean, or even other objects, for example, `worksheet.addTable(targetRange, true)`, where `targetRange` is an object created elsewhere in the script.</span></span>
* <span data-ttu-id="8b48b-219">方法可以返回标量属性 (名称、地址等) 或其他对象 (区域、图表) ，或者不返回任何 (例如方法) 的情况。 `delete`</span><span class="sxs-lookup"><span data-stu-id="8b48b-219">Methods can return a thing such as a scalar property (name, address, etc.) or another object (range, chart), or not return anything at all (such as the case with `delete` methods).</span></span> <span data-ttu-id="8b48b-220">通过声明变量或分配给现有变量，您可以接收该方法返回的值。</span><span class="sxs-lookup"><span data-stu-id="8b48b-220">You receive what the method returns by declaring a variable or assigning to an existing variable.</span></span> <span data-ttu-id="8b48b-221">您可以在语句的左侧看到 ，例如 `const table = worksheet.addTable('A1:D6', true)` 。</span><span class="sxs-lookup"><span data-stu-id="8b48b-221">You can see that on the left hand side of statement such as `const table = worksheet.addTable('A1:D6', true)`.</span></span>
* <span data-ttu-id="8b48b-222">大多数情况下，Office 脚本对象模型由包含用于链接 Excel 对象模型的各个部分的方法的对象组成。</span><span class="sxs-lookup"><span data-stu-id="8b48b-222">For the most part, the Office Scripts object model consists of objects with methods that link various parts of the Excel object model.</span></span> <span data-ttu-id="8b48b-223">很少遇到标量或对象值的属性。</span><span class="sxs-lookup"><span data-stu-id="8b48b-223">Very rarely you'll come across properties that are of scalar or object values.</span></span>
* <span data-ttu-id="8b48b-224">在 Office 脚本中，Excel 对象模型方法必须包含开放式关闭括号。</span><span class="sxs-lookup"><span data-stu-id="8b48b-224">In Office Scripts, an Excel object model method has to contain open-close parentheses.</span></span> <span data-ttu-id="8b48b-225">不允许使用不带它们的方法 (例如将方法分配给变量) 。</span><span class="sxs-lookup"><span data-stu-id="8b48b-225">Using methods without them is not allowed (such as assigning a method to a variable).</span></span>

<span data-ttu-id="8b48b-226">让我们看一下对象上的一些 `workbook` 方法。</span><span class="sxs-lookup"><span data-stu-id="8b48b-226">Let's look at a few methods on the `workbook` object.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
    // Return a boolean (true or false) setting of whether the workbook is set to auto-save or not. 
    const autoSave = workbook.getAutoSave(); 
    // Get workbook name.
    const name = workbook.getName();
    // Get active cell range object.
    const cell = workbook.getActiveCell();
    // Get table named SALES.
    const cell = workbook.getTable('SALES');
    // Get all slicer objects.
    const slicers = workbook.getSlicers();
}
```

<span data-ttu-id="8b48b-227">在此示例中：</span><span class="sxs-lookup"><span data-stu-id="8b48b-227">In this example:</span></span>

* <span data-ttu-id="8b48b-228">对象的方法 `workbook` （如 和 `getAutoSave()` ）返回 `getName()` 标量属性 (string、number、boolean) 。</span><span class="sxs-lookup"><span data-stu-id="8b48b-228">The methods of the `workbook` object such as `getAutoSave()` and `getName()` return a scalar property (string, number, boolean).</span></span>
* <span data-ttu-id="8b48b-229">方法（如 `getActiveCell()` 返回另一个对象）。</span><span class="sxs-lookup"><span data-stu-id="8b48b-229">Methods such as `getActiveCell()` return another object.</span></span>
* <span data-ttu-id="8b48b-230">`getTable()`此方法接受一个 (表名称的参数，) 并返回工作簿中的特定表。</span><span class="sxs-lookup"><span data-stu-id="8b48b-230">The `getTable()` method accepts an argument (table name in this case) and returns a specific table in the workbook.</span></span>
* <span data-ttu-id="8b48b-231">该方法返回一个 (，该数组集合在很多位置) 工作簿中所有切片器对象 `getSlicers()` 的集合。</span><span class="sxs-lookup"><span data-stu-id="8b48b-231">The `getSlicers()` method returns an array (referred to in many places as a collection) of all slicer objects within the workbook.</span></span>

<span data-ttu-id="8b48b-232">你会注意到，所有这些方法都有前缀，这只是 Office 脚本对象模型中使用的约定，用于传达 `get` 该方法将返回某些内容。</span><span class="sxs-lookup"><span data-stu-id="8b48b-232">You'll notice that all of these methods have a `get` prefix, which is just a convention used in the Office Scripts object model to convey that the method is returning something.</span></span> <span data-ttu-id="8b48b-233">它们通常也称为"getters"。</span><span class="sxs-lookup"><span data-stu-id="8b48b-233">They are also commonly referred to as 'getters'.</span></span>

<span data-ttu-id="8b48b-234">有两种其他类型的方法，我们将在下一个示例中看到：</span><span class="sxs-lookup"><span data-stu-id="8b48b-234">There are two other types of methods that we'll now see in the next example:</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
    // Get a worksheet named 'Sheet1.
    const sheet = workbook.getWorksheet('Sheet1'); 
    // Set name to SALES.
    sheet.setName('SALES');
    // Position the worksheet at the beginning.
    sheet.setPosition(0);
}
```

<span data-ttu-id="8b48b-235">在此示例中：</span><span class="sxs-lookup"><span data-stu-id="8b48b-235">In this example:</span></span>

* <span data-ttu-id="8b48b-236">`setName()`方法为工作表设置一个新名称。</span><span class="sxs-lookup"><span data-stu-id="8b48b-236">The `setName()` method sets a new name to the worksheet.</span></span> <span data-ttu-id="8b48b-237">`setPosition()` 将位置设置到第一个单元格。</span><span class="sxs-lookup"><span data-stu-id="8b48b-237">`setPosition()` sets the position to the first cell.</span></span>
* <span data-ttu-id="8b48b-238">这些方法通过设置工作簿的属性或行为来修改 Excel 文件。</span><span class="sxs-lookup"><span data-stu-id="8b48b-238">Such methods modify the Excel file by setting a property or behavior of the workbook.</span></span> <span data-ttu-id="8b48b-239">这些方法称为"setters"。</span><span class="sxs-lookup"><span data-stu-id="8b48b-239">These methods are called 'setters'.</span></span>
* <span data-ttu-id="8b48b-240">通常，"setter"具有配套"getter"，例如 `worksheet.getPosition` `worksheet.setPosition` 和 ，两者都是方法。</span><span class="sxs-lookup"><span data-stu-id="8b48b-240">Typically 'setters' have a companion 'getter', for example, `worksheet.getPosition` and `worksheet.setPosition`, both of which are methods.</span></span>

#### <a name="undefined-and-null-primitive-types"></a><span data-ttu-id="8b48b-241">`undefined` 和 `null` 基元类型</span><span class="sxs-lookup"><span data-stu-id="8b48b-241">`undefined` and `null` primitive types</span></span>

<span data-ttu-id="8b48b-242">以下是您必须注意的两种基元数据类型：</span><span class="sxs-lookup"><span data-stu-id="8b48b-242">The following are two primitive data types that you must be aware of:</span></span>

1. <span data-ttu-id="8b48b-243">该值 [`null`](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/null) 表示有意缺少任何对象值。</span><span class="sxs-lookup"><span data-stu-id="8b48b-243">The value [`null`](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/null) represents the intentional absence of any object value.</span></span> <span data-ttu-id="8b48b-244">它是 JavaScript 的基元值之一，用于指示变量没有值。</span><span class="sxs-lookup"><span data-stu-id="8b48b-244">It is one of JavaScript's primitive values and is used to indicate that a variable has no value.</span></span>
1. <span data-ttu-id="8b48b-245">尚未分配值的变量的类型为 [`undefined`](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/undefined) 。</span><span class="sxs-lookup"><span data-stu-id="8b48b-245">A variable that has not been assigned a value is of type [`undefined`](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/undefined).</span></span> <span data-ttu-id="8b48b-246">如果要求值的变量没有分配的值，则方法或语句 `undefined` 也可以返回。</span><span class="sxs-lookup"><span data-stu-id="8b48b-246">A method or statement can also return `undefined` if the variable that's being evaluated doesn't have an assigned value.</span></span>

<span data-ttu-id="8b48b-247">这两种类型会作为错误处理的一部分出现，如果未正确处理，则可能导致非常麻烦。</span><span class="sxs-lookup"><span data-stu-id="8b48b-247">These two types crop up as part of error handling and can cause quite a bit of headache if not handled properly.</span></span> <span data-ttu-id="8b48b-248">幸运的是，TypeScript/JavaScript 提供了一种检查变量的类型或 `undefined` 的方法 `null` 。</span><span class="sxs-lookup"><span data-stu-id="8b48b-248">Fortunately, TypeScript/JavaScript offers a way to check if a variable is of type `undefined` or `null`.</span></span> <span data-ttu-id="8b48b-249">我们将在稍后部分讨论其中一些检查，包括错误处理。</span><span class="sxs-lookup"><span data-stu-id="8b48b-249">We will talk about some of those checks in later sections, including error handling.</span></span>

#### <a name="method-chaining"></a><span data-ttu-id="8b48b-250">方法链接</span><span class="sxs-lookup"><span data-stu-id="8b48b-250">Method chaining</span></span>

<span data-ttu-id="8b48b-251">可以使用点表示法连接从方法返回的对象，以缩短代码。</span><span class="sxs-lookup"><span data-stu-id="8b48b-251">You can use dot notation to connect objects being returned from a method to shorten your code.</span></span> <span data-ttu-id="8b48b-252">有时，此技术使代码易于阅读和管理。</span><span class="sxs-lookup"><span data-stu-id="8b48b-252">Sometimes this technique makes the code easy to read and manage.</span></span> <span data-ttu-id="8b48b-253">但是，要注意一些内容。</span><span class="sxs-lookup"><span data-stu-id="8b48b-253">However, there are few things to be aware of.</span></span> <span data-ttu-id="8b48b-254">让我们看一下以下示例。</span><span class="sxs-lookup"><span data-stu-id="8b48b-254">Let's look at the following examples.</span></span>

<span data-ttu-id="8b48b-255">下面的代码获取活动单元格和下一个单元格，然后设置值。</span><span class="sxs-lookup"><span data-stu-id="8b48b-255">The following code gets the active cell and the next cell, then sets the value.</span></span> <span data-ttu-id="8b48b-256">这是使用链接的良好候选项，因为此代码将一定会成功。</span><span class="sxs-lookup"><span data-stu-id="8b48b-256">This is a good candidate to use chaining as this code will succeed all the time.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
    workbook.getActiveCell().getOffsetRange(0,1).setValue('Next cell');
}
```

<span data-ttu-id="8b48b-257">但是，以下代码 (获取名为 **SALES** 的表，并打开其带状列样式) 一个问题。</span><span class="sxs-lookup"><span data-stu-id="8b48b-257">However, the following code (which gets a table named **SALES** and turns on its banded column style) has an issue.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  workbook.getTable('SALES').setShowBandedColumns(true);
}
```

<span data-ttu-id="8b48b-258">如果 **SALES** 表不存在，如何？</span><span class="sxs-lookup"><span data-stu-id="8b48b-258">What if the **SALES** table doesn't exist?</span></span> <span data-ttu-id="8b48b-259">脚本将失败，并返回错误 (如) ，因为返回 (这是一个指示没有表（如 `getTable('SALES')` `undefined` **SALES**) ）的 JavaScript 类型。</span><span class="sxs-lookup"><span data-stu-id="8b48b-259">The script will fail with an error (shown next) because `getTable('SALES')` returns `undefined` (which is a JavaScript type indicating that there is no table such as **SALES**).</span></span> <span data-ttu-id="8b48b-260">调用 `setShowBandedColumns` 上的 `undefined` 方法没有任何意义，即 ，因此脚本 `undefined.setShowBandedColumns(true)` 会出现错误。</span><span class="sxs-lookup"><span data-stu-id="8b48b-260">Calling the `setShowBandedColumns` method on `undefined` makes no sense, that is, `undefined.setShowBandedColumns(true)`, and hence the script ends in an error.</span></span>

```text
Line 2: Cannot read property 'setShowBandedColumns' of undefined
```

<span data-ttu-id="8b48b-261">当引用或方法[](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Operators/Optional_chaining)可能为 或 (这是 `undefined` JavaScript 用于指示未分配或非不存在的对象或结果) 来处理此情况时，可以使用可选的链接运算符，从而提供一种通过连接对象简化访问值的方法。 `null`</span><span class="sxs-lookup"><span data-stu-id="8b48b-261">You could use the [optional chaining operator](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Operators/Optional_chaining) that provides a way to simplify accessing values through connected objects when it's possible that a reference or method may be `undefined` or `null` (which is JavaScript's way of indicating an unassigned or nonexistent object or result) to handle this condition.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
    // This line will not fail as the setShowBandedColumns method is executed only if the SALES table is present.
    workbook.getTable('SALES')?.setShowBandedColumns(true); 
}
```

<span data-ttu-id="8b48b-262">如果要处理不存在的对象条件或由方法返回的类型，最好从 方法分配返回值并单独 `undefined` 处理。</span><span class="sxs-lookup"><span data-stu-id="8b48b-262">If you wish to handle nonexistent object conditions or `undefined` type being returned by a method, then it is better to assign the return value from the method and handle that separately.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
    const salesTable = workbook.getTable('SALES');
    if (salesTable) {
        salesTable.setShowBandedColumns(true);
    } else { 
        // Handle this condition.
    }
}
```

#### <a name="get-object-reference"></a><span data-ttu-id="8b48b-263">获取对象引用</span><span class="sxs-lookup"><span data-stu-id="8b48b-263">Get object reference</span></span>

<span data-ttu-id="8b48b-264">`workbook`对象在 函数中给予 `main` 你。</span><span class="sxs-lookup"><span data-stu-id="8b48b-264">The `workbook` object is given to you in the `main` function.</span></span> <span data-ttu-id="8b48b-265">您可以开始使用 对象 `workbook` 并直接访问其方法。</span><span class="sxs-lookup"><span data-stu-id="8b48b-265">You can begin to use the `workbook` object and access its methods directly.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
    // Get workbook name.
    const name = workbook.getName();
    // Display name to console.
    console.log(name);
}
```

<span data-ttu-id="8b48b-266">若要使用工作簿中的所有其他对象，请从 object 开始，然后向下转到层次结构，直到到达要 `workbook` 查找的对象。</span><span class="sxs-lookup"><span data-stu-id="8b48b-266">For using all other objects within the workbook, begin with `workbook` object and go down the hierarchy until you get to the object you are looking for.</span></span> <span data-ttu-id="8b48b-267">可以通过使用对象的方法获取对象或检索对象集合来获取对象引用， `get` 如下所示：</span><span class="sxs-lookup"><span data-stu-id="8b48b-267">You can get the object reference by fetching the object using its `get` method or by retrieving the collection of objects as shown below:</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
    // Get the active worksheet.
    const sheet = workbook.getActiveWorksheet();
    // Fetch using an ID or key.
    const sheet = workbook.getWorksheet('SomeSheetName');
    // Invoke methods on the object.
    sheet.setPosition(0); 
    
    // Get collection of methods.
    const tables = sheet.getTables();
    console.log('Total tables in this sheet: ' + tables.length);
}
```

#### <a name="check-if-an-object-exists-then-delete-and-add"></a><span data-ttu-id="8b48b-268">检查对象是否存在，然后删除并添加</span><span class="sxs-lookup"><span data-stu-id="8b48b-268">Check if an object exists, then delete, and add</span></span>

<span data-ttu-id="8b48b-269">对于创建对象（如使用预定义名称）来说，最好始终删除可能存在的类似对象，然后添加它。</span><span class="sxs-lookup"><span data-stu-id="8b48b-269">For creating an object, say with a predefined name, it is always better to remove a similar object that may exist and then add it.</span></span> <span data-ttu-id="8b48b-270">可以使用以下模式实现此要求。</span><span class="sxs-lookup"><span data-stu-id="8b48b-270">You can do that using the following pattern.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Name of the worksheet to be added. 
  let name = "Index";
  // Check if the worksheet already exists. If not, add the worksheet.
  let sheet = workbook.getWorksheet('Index');
  if (sheet) {
    console.log(`Worksheet by the name ${name} already exists. Deleting it.`);
    // Call the delete method on the object to remove it. 
    sheet.delete();
  } 
    // Add a blank worksheet. 
  console.log(`Adding the worksheet named  ${name}.`)
  const indexSheet = workbook.addWorksheet("Index");
}

```

<span data-ttu-id="8b48b-271">或者，要删除可能存在或不存在的对象，请使用以下模式。</span><span class="sxs-lookup"><span data-stu-id="8b48b-271">Alternatively, for deleting an object that may or may not exist, use the following pattern.</span></span>

```TypeScript
    // The ? preceding delete() will ensure that the API is only invoked if the object exists. 
    workbook.getWorksheet('Index')?.delete(); 
```

#### <a name="note-about-adding-an-object"></a><span data-ttu-id="8b48b-272">有关添加对象的注释</span><span class="sxs-lookup"><span data-stu-id="8b48b-272">Note about adding an object</span></span>

<span data-ttu-id="8b48b-273">若要创建、插入或添加对象（如切片器、数据透视表、 **工作表等** ），请使用相应的 add_Object_ 方法。</span><span class="sxs-lookup"><span data-stu-id="8b48b-273">To create, insert, or add an object such as a slicer, pivot table, worksheet, etc., use the corresponding **add_Object_** method.</span></span> <span data-ttu-id="8b48b-274">此方法可用于其父对象。</span><span class="sxs-lookup"><span data-stu-id="8b48b-274">Such a method is available on its parent object.</span></span> <span data-ttu-id="8b48b-275">例如， `addChart()` 方法可用于 `worksheet` 对象。</span><span class="sxs-lookup"><span data-stu-id="8b48b-275">For example, the `addChart()` method is available on `worksheet` object.</span></span> <span data-ttu-id="8b48b-276">the **add_Object_** method returns the object it creates.</span><span class="sxs-lookup"><span data-stu-id="8b48b-276">The **add_Object_** method returns the object it creates.</span></span> <span data-ttu-id="8b48b-277">接收返回的值，并稍后在脚本中使用它。</span><span class="sxs-lookup"><span data-stu-id="8b48b-277">Receive the returned value and use it later in your script.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Add object and get a reference to it. 
  const indexSheet = workbook.addWorksheet("Index");
  // Use it elsewhere in the script 
  console.log(indexSheet.getPosition());
}

```

<span data-ttu-id="8b48b-278">或者，要删除可能存在或不存在的对象，请使用此模式：</span><span class="sxs-lookup"><span data-stu-id="8b48b-278">Alternatively, for deleting an object that may or may not exist, use this pattern:</span></span>

```TypeScript
    workbook.getWorksheet('Index')?.delete(); // The ? preceding delete() will ensure that the API is only invoked if the object exists. 
```

#### <a name="collections"></a><span data-ttu-id="8b48b-279">收藏</span><span class="sxs-lookup"><span data-stu-id="8b48b-279">Collections</span></span>

<span data-ttu-id="8b48b-280">集合是表、图表、列等对象，可以检索为数组并经过访问进行处理。</span><span class="sxs-lookup"><span data-stu-id="8b48b-280">Collections are objects such as tables, charts, columns, etc. that can be retrieved as an array and iterated over for processing.</span></span> <span data-ttu-id="8b48b-281">可以使用相应的方法检索集合，然后使用多种 TypeScript 数组遍历技术之一在循环中处理数据， `get` 例如：</span><span class="sxs-lookup"><span data-stu-id="8b48b-281">You can retrieve a collection using the corresponding `get` method and process the data in a loop using one of many TypeScript array traversal techniques such as:</span></span>

* [<span data-ttu-id="8b48b-282">`for` 或 `while`</span><span class="sxs-lookup"><span data-stu-id="8b48b-282">`for` or `while`</span></span>](https://developer.mozilla.org/docs/Web/JavaScript/Guide/Loops_and_iteration)
* [`for..of`](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Statements/for...of)
* [`forEach`](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Array/forEach)

* [<span data-ttu-id="8b48b-283">数组的语言基础知识</span><span class="sxs-lookup"><span data-stu-id="8b48b-283">Language basics of arrays</span></span>](https://developer.mozilla.org//docs/Learn/JavaScript/First_steps/Arrays)

<span data-ttu-id="8b48b-284">此脚本演示如何使用 Office 脚本 API 中支持的集合。</span><span class="sxs-lookup"><span data-stu-id="8b48b-284">This script demonstrates how to use collections supported in Office Scripts APIs.</span></span> <span data-ttu-id="8b48b-285">它使用随机颜色为文件的每个工作表选项卡着色。</span><span class="sxs-lookup"><span data-stu-id="8b48b-285">It colors each worksheet tab in the file with a random color.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {

  // Get all sheets as a collection.
  const sheets = workbook.getWorksheets();
  const names = sheets.map ((sheet) => sheet.getName());
  console.log(names);
  console.log(`Total worksheets inside of this workbook: ${sheets.length}`);
  // Get information from specific sheets within the collection.
  console.log(`First sheet name is: ${names[0]}`);
  if (sheets.length > 1) {
    console.log(`Last sheet's Id is: ${sheets[sheets.length -1].getId()}`);
  }
  // Color each worksheet with random color.
  for (const sheet of sheets) {
    sheet.setTabColor(`#${Math.random().toString(16).substr(-6)}`);
  }
}
```

## <a name="type-declarations"></a><span data-ttu-id="8b48b-286">类型声明</span><span class="sxs-lookup"><span data-stu-id="8b48b-286">Type declarations</span></span>

<span data-ttu-id="8b48b-287">类型声明可帮助用户了解他们处理的变量的类型。</span><span class="sxs-lookup"><span data-stu-id="8b48b-287">Type declarations help users understand the type of variable they are dealing with.</span></span> <span data-ttu-id="8b48b-288">它有助于自动完成方法，并有助于进行开发时间质量检查。</span><span class="sxs-lookup"><span data-stu-id="8b48b-288">It helps with auto-completion of methods and assists in development time quality checks.</span></span>

<span data-ttu-id="8b48b-289">您可以在脚本中的不同位置找到类型声明，包括函数声明、变量声明、IntelliSense定义等。</span><span class="sxs-lookup"><span data-stu-id="8b48b-289">You can find type declarations in the script in various places including function declaration, variable declaration, IntelliSense definitions, etc.</span></span>

<span data-ttu-id="8b48b-290">示例：</span><span class="sxs-lookup"><span data-stu-id="8b48b-290">Examples:</span></span>

* `function main(workbook: ExcelScript.Workbook)`
* `let myRange: ExcelScript.Range;`
* `function getMaxAmount(range: ExcelScript.Range): number`

<span data-ttu-id="8b48b-291">您可以在代码编辑器中轻松识别类型，因为它通常以不同的颜色显示。</span><span class="sxs-lookup"><span data-stu-id="8b48b-291">You can identify the types easily in the code editor as it usually appears distinctly in a different color.</span></span> <span data-ttu-id="8b48b-292">通常， `:` 类型声明之前有一个冒号。</span><span class="sxs-lookup"><span data-stu-id="8b48b-292">A colon `:` usually precedes the type declaration.</span></span>  

<span data-ttu-id="8b48b-293">在 TypeScript 中，写入类型是可选的，因为类型推断允许你获得大量功能，而无需编写其他代码。</span><span class="sxs-lookup"><span data-stu-id="8b48b-293">Writing types can be optional in TypeScript because type inference allows you to get a lot of power without writing additional code.</span></span> <span data-ttu-id="8b48b-294">大多数情况下，TypeScript 语言非常适用于推断变量类型。</span><span class="sxs-lookup"><span data-stu-id="8b48b-294">For the most part, the TypeScript language is good at inferring the types of variables.</span></span> <span data-ttu-id="8b48b-295">但是，在某些情况下，如果语言无法清楚地标识类型，则 Office 脚本需要显式定义类型声明。</span><span class="sxs-lookup"><span data-stu-id="8b48b-295">However, in certain cases, Office Scripts require the type declarations to be explicitly defined if the language is unable to clearly identify the type.</span></span> <span data-ttu-id="8b48b-296">此外，Office 脚本 `any` 中不允许显式或隐式。</span><span class="sxs-lookup"><span data-stu-id="8b48b-296">Also, explicit or implicit `any` is not allowed in Office Script.</span></span> <span data-ttu-id="8b48b-297">稍后将详细了解这一点。</span><span class="sxs-lookup"><span data-stu-id="8b48b-297">More on that later.</span></span>

### <a name="excelscript-types"></a><span data-ttu-id="8b48b-298">`ExcelScript` types</span><span class="sxs-lookup"><span data-stu-id="8b48b-298">`ExcelScript` types</span></span>

<span data-ttu-id="8b48b-299">在 Office 脚本中，您将使用以下类型。</span><span class="sxs-lookup"><span data-stu-id="8b48b-299">In Office Scripts, you will use the following kinds of types.</span></span>

* <span data-ttu-id="8b48b-300">本机语言类型，如 `number` `string` `object` `boolean` 、、、、 `null` 等。</span><span class="sxs-lookup"><span data-stu-id="8b48b-300">Native language types such as `number`, `string`, `object`, `boolean`, `null`, etc.</span></span>
* <span data-ttu-id="8b48b-301">Excel API 类型。</span><span class="sxs-lookup"><span data-stu-id="8b48b-301">Excel API types.</span></span> <span data-ttu-id="8b48b-302">它们以 `ExcelScript` 开头。</span><span class="sxs-lookup"><span data-stu-id="8b48b-302">They begin with `ExcelScript`.</span></span> <span data-ttu-id="8b48b-303">例如， `ExcelScript.Range` `ExcelScript.Table` 、 等。</span><span class="sxs-lookup"><span data-stu-id="8b48b-303">For example, `ExcelScript.Range`, `ExcelScript.Table`, etc.</span></span>
* <span data-ttu-id="8b48b-304">您可能在脚本 using 语句中定义的任何自定义 `interface` 接口。</span><span class="sxs-lookup"><span data-stu-id="8b48b-304">Any custom interfaces you may have defined in the script using `interface` statements.</span></span>

<span data-ttu-id="8b48b-305">接下来，请参阅每个组的示例。</span><span class="sxs-lookup"><span data-stu-id="8b48b-305">See examples of each of these groups next.</span></span>

<span data-ttu-id="8b48b-306">**_本机语言类型_**</span><span class="sxs-lookup"><span data-stu-id="8b48b-306">**_Native language types_**</span></span>

<span data-ttu-id="8b48b-307">在下面的示例中，请注意已使用 `string` 、 `number` 和 `boolean` 的位置。</span><span class="sxs-lookup"><span data-stu-id="8b48b-307">In the following example, notice places where `string`, `number`, and `boolean` have been used.</span></span> <span data-ttu-id="8b48b-308">这些是本机 **TypeScript** 语言类型。</span><span class="sxs-lookup"><span data-stu-id="8b48b-308">These are native **TypeScript** language types.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook)
{
  const table = workbook.getActiveWorksheet().getTables()[0];
  const sales = table.getColumnByName('Sales').getRange().getValues();
  console.log(sales);
  // Add 100 to each value.
  const revisedSales = salesAs1DArray.map(data => data as number + 100);
  // Add a column.
  table.addColumn(-1, revisedSales);  
}
/**
 * Extract a column from 2D array and return result.
 */
function extractColumn(data: (string | number | boolean)[][], index: number): (string | number | boolean)[] {

  const column = data.map((row) => {
    return row[index];
  })
  return column;
}
/**
 * Convert a flat array into a 2D array that can be used as range column.
 */
function convertColumnTo2D(data: (string | number | boolean)[]): (string | number | boolean)[][] {

  const columnAs2D = data.map((row) => {
    return [row];
  })
  return columnAs2D;
}
```

<span data-ttu-id="8b48b-309">**_ExcelScript 类型_**</span><span class="sxs-lookup"><span data-stu-id="8b48b-309">**_ExcelScript types_**</span></span>

<span data-ttu-id="8b48b-310">在下面的示例中，帮助程序函数采用两个参数。</span><span class="sxs-lookup"><span data-stu-id="8b48b-310">In the following example, a helper function takes two arguments.</span></span> <span data-ttu-id="8b48b-311">第一个 `sheet` 类型为变量 `ExcelScript.Worksheet` 。</span><span class="sxs-lookup"><span data-stu-id="8b48b-311">The first one is the `sheet` variable which is of type `ExcelScript.Worksheet` type.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
    const sheet = workbook.getWorksheet('Sheet5');
    const data = ['2016', 'Bikes', 'Seats', '1500', .05];
    addRow(sheet, data);
    return;
}

function addRow(sheet: ExcelScript.Worksheet, data: (string | number | boolean)[]): void {

    const usedRange = sheet.getUsedRange();
    let startCell: ExcelScript.Range;
    // If the sheet is empty, then use A1 as starting cell for update.
    if (usedRange) { 
      startCell = usedRange.getLastRow().getCell(0, 0).getOffsetRange(1, 0);
    } else {
      startCell = sheet.getRange('A1');
    }
    console.log(startCell.getAddress());
    const targetRange = startCell.getResizedRange(0, data.length - 1);      
    targetRange.setValues([data]);
    return;
}
```

<span data-ttu-id="8b48b-312">**_自定义类型_**</span><span class="sxs-lookup"><span data-stu-id="8b48b-312">**_Custom types_**</span></span>

<span data-ttu-id="8b48b-313">自定义接口 `ReportImages` 用于将图像返回到另一个流操作。</span><span class="sxs-lookup"><span data-stu-id="8b48b-313">The custom interface `ReportImages` is used to return images to another flow action.</span></span> <span data-ttu-id="8b48b-314">该 `main` 函数声明包括指示 TypeScript 正在返回该 `: ReportImages` 类型的对象的指令。</span><span class="sxs-lookup"><span data-stu-id="8b48b-314">The `main` function declaration includes `: ReportImages` instruction to tell TypeScript that an object of that type is being returned.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook): ReportImages {
  let chart = workbook.getWorksheet("Sheet1").getCharts()[0];
  const table = workbook.getWorksheet('InvoiceAmounts').getTables()[0];
  
  const chartImage = chart.getImage();
  const tableImage = table.getRange().getImage();
  return {
    chartImage,
    tableImage
  }
}

interface ReportImages {
  chartImage: string
  tableImage: string
}
```

### <a name="type-assertion-overriding-the-type"></a><span data-ttu-id="8b48b-315">类型断言 (覆盖类型) </span><span class="sxs-lookup"><span data-stu-id="8b48b-315">Type assertion (overriding the type)</span></span>

<span data-ttu-id="8b48b-316">如 TypeScript [文档](https://www.typescriptlang.org/docs/handbook/basic-types.html#type-assertions) 所规定，"有时你最终会了解比 TypeScript 更多的值。</span><span class="sxs-lookup"><span data-stu-id="8b48b-316">As the TypeScript [documentation](https://www.typescriptlang.org/docs/handbook/basic-types.html#type-assertions) states, "Sometimes you'll end up in a situation where you'll know more about a value than TypeScript does.</span></span> <span data-ttu-id="8b48b-317">通常，当您知道某实体的类型可能比其当前类型更具体时，就会发生这种情况。</span><span class="sxs-lookup"><span data-stu-id="8b48b-317">Usually, this will happen when you know the type of some entity could be more specific than its current type.</span></span> <span data-ttu-id="8b48b-318">类型断言是告诉编译器"信任我，我了解我正在做什么"的一种方法。</span><span class="sxs-lookup"><span data-stu-id="8b48b-318">Type assertions are a way to tell the compiler “trust me, I know what I'm doing.”</span></span> <span data-ttu-id="8b48b-319">类型断言与其他语言中的类型转换类似，但它不执行数据的特殊检查或重构。</span><span class="sxs-lookup"><span data-stu-id="8b48b-319">A type assertion is like a type cast in other languages, but it performs no special checking or restructuring of data.</span></span> <span data-ttu-id="8b48b-320">它不会影响运行时，并且完全由编译器使用。"</span><span class="sxs-lookup"><span data-stu-id="8b48b-320">It has no runtime impact and is used purely by the compiler."</span></span>

<span data-ttu-id="8b48b-321">您可以使用 关键字或尖括号断言类型， `as` 如下面的代码所示。</span><span class="sxs-lookup"><span data-stu-id="8b48b-321">You can assert the type using the `as` keyword or using angle brackets as shown in following code.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  let data = workbook.getActiveCell().getValue();
  // Since the add10 function only accepts number, assert data's type as number, otherwise the script cannot be run.
  const answer1 = add10(data as number);
  const answer2 = add10(<number> data);
}

function add10(data: number) { 
  return data + 10;
}
```

#### <a name="any-type-in-the-script"></a><span data-ttu-id="8b48b-322">脚本中的"any"类型</span><span class="sxs-lookup"><span data-stu-id="8b48b-322">'any' type in the script</span></span>

<span data-ttu-id="8b48b-323">[TypeScript 网站指出](https://www.typescriptlang.org/docs/handbook/basic-types.html#any)：</span><span class="sxs-lookup"><span data-stu-id="8b48b-323">The [TypeScript website states](https://www.typescriptlang.org/docs/handbook/basic-types.html#any):</span></span>

  <span data-ttu-id="8b48b-324">在某些情况下，并非所有类型信息都可用，或者其声明会花一些不恰当的努力。</span><span class="sxs-lookup"><span data-stu-id="8b48b-324">In some situations, not all type information is available or its declaration would take an inappropriate amount of effort.</span></span> <span data-ttu-id="8b48b-325">对于在没有 TypeScript 或第三方库的情况下编写的代码中的值，可能会出现这些错误。</span><span class="sxs-lookup"><span data-stu-id="8b48b-325">These may occur for values from code that has been written without TypeScript or a 3rd party library.</span></span> <span data-ttu-id="8b48b-326">在这些情况下，我们可能需要选择退出类型检查。</span><span class="sxs-lookup"><span data-stu-id="8b48b-326">In these cases, we might want to opt-out of type checking.</span></span> <span data-ttu-id="8b48b-327">为此，我们将这些值标记为 `any` 以下类型：</span><span class="sxs-lookup"><span data-stu-id="8b48b-327">To do so, we label these values with the `any` type:</span></span>

  ```TypeScript
  declare function getValue(key: string): any;
  // OK, return value of 'getValue' is not checked
  const str: string = getValue("myString");
  ```

<span data-ttu-id="8b48b-328">**不允许 `any` 显式**</span><span class="sxs-lookup"><span data-stu-id="8b48b-328">**Explicit `any` is NOT allowed**</span></span>

```TypeScript
// This is not allowed
let someVariable: any; 
```

<span data-ttu-id="8b48b-329">类型为 Office 脚本处理 Excel API 的方式 `any` 带来了挑战。</span><span class="sxs-lookup"><span data-stu-id="8b48b-329">The `any` type presents challenges to the way Office Scripts processes the Excel APIs.</span></span> <span data-ttu-id="8b48b-330">当将变量发送到 Excel API 进行处理时，会导致问题。</span><span class="sxs-lookup"><span data-stu-id="8b48b-330">It causes issues when the variables are sent to Excel APIs for processing.</span></span> <span data-ttu-id="8b48b-331">了解脚本中使用的变量类型对于处理脚本至关重要，因此禁止对具有类型的任何变量进行 `any` 显式定义。</span><span class="sxs-lookup"><span data-stu-id="8b48b-331">Knowing the type of variables used in the script is essential to the processing of script and hence explicit definition of any variable with `any` type is prohibited.</span></span> <span data-ttu-id="8b48b-332">如果脚本中声明了类型的任何 (，在运行脚本脚本之前) 出现编译时 `any` 错误。</span><span class="sxs-lookup"><span data-stu-id="8b48b-332">You will receive a compile-time error (error prior to running the script) if there is any variable with `any` type declared in the script.</span></span> <span data-ttu-id="8b48b-333">You will see an error in the editor as well.</span><span class="sxs-lookup"><span data-stu-id="8b48b-333">You will see an error in the editor as well.</span></span>

![显式"任何"错误](../../images/getting-started-eanyi.png)

![输出中显示的显式"任何"错误](../../images/getting-started-expany.png)

<span data-ttu-id="8b48b-336">在上一图像中显示的代码中，指示第 5 行 `[5, 16] Explicit Any is not allowed` 16 列声明 `any` 类型。</span><span class="sxs-lookup"><span data-stu-id="8b48b-336">In the code displayed in the previous image, `[5, 16] Explicit Any is not allowed` indicates that line 5 column 16 declares the `any` type.</span></span> <span data-ttu-id="8b48b-337">这可以帮助您找到包含错误的代码行。</span><span class="sxs-lookup"><span data-stu-id="8b48b-337">This helps you locate the line of code that contains the error.</span></span>

<span data-ttu-id="8b48b-338">若要解决此问题，请始终声明变量的类型。</span><span class="sxs-lookup"><span data-stu-id="8b48b-338">To get around this issue, always declare the type of the variable.</span></span>

<span data-ttu-id="8b48b-339">如果你不确定变量的类型，TypeScript 中的一个酷技巧允许你定义 [联合类型](https://www.typescriptlang.org/docs/handbook/unions-and-intersections.html)。</span><span class="sxs-lookup"><span data-stu-id="8b48b-339">If you are uncertain about the type of a variable, one cool trick in TypeScript allows you to define [union types](https://www.typescriptlang.org/docs/handbook/unions-and-intersections.html).</span></span> <span data-ttu-id="8b48b-340">这可用于变量来保存区域值，可以是许多类型。</span><span class="sxs-lookup"><span data-stu-id="8b48b-340">This can be used for variables to hold a range values, which can be of many types.</span></span>

```TypeScript
// Define value as a union type rather than 'any' type.
let value: (string | number | boolean);
value = someValue_from_another_source;
//...
someRange.setValue(value);
```

### <a name="type-inference"></a><span data-ttu-id="8b48b-341">类型推断</span><span class="sxs-lookup"><span data-stu-id="8b48b-341">Type inference</span></span>

<span data-ttu-id="8b48b-342">在 TypeScript 中，当没有[](https://www.typescriptlang.org/docs/handbook/type-inference.html)显式类型注释时，有几种使用类型推断来提供类型信息的位置。</span><span class="sxs-lookup"><span data-stu-id="8b48b-342">In TypeScript, there are several places where [type inference](https://www.typescriptlang.org/docs/handbook/type-inference.html) is used to provide type information when there is no explicit type annotation.</span></span> <span data-ttu-id="8b48b-343">例如，x 变量的类型被推断为以下代码中的一个数字。</span><span class="sxs-lookup"><span data-stu-id="8b48b-343">For example, the type of the x variable is inferred to be a number in the following code.</span></span>

```TypeScript
let x = 3;
//  ^ = let x: number
```

<span data-ttu-id="8b48b-344">这种推断发生在初始化变量和成员、设置参数默认值以及确定函数返回类型时。</span><span class="sxs-lookup"><span data-stu-id="8b48b-344">This kind of inference takes place when initializing variables and members, setting parameter default values, and determining function return types.</span></span>

### <a name="no-implicit-any-rule"></a><span data-ttu-id="8b48b-345">no-implicit-any 规则</span><span class="sxs-lookup"><span data-stu-id="8b48b-345">no-implicit-any rule</span></span>

<span data-ttu-id="8b48b-346">脚本需要用于显式或隐式声明的变量类型。</span><span class="sxs-lookup"><span data-stu-id="8b48b-346">A script requires the types of the variables used to be explicitly or implicitly declared.</span></span> <span data-ttu-id="8b48b-347">如果 TypeScript 编译器无法确定变量 (或者因为类型未显式声明或类型推断无法进行) ，则你将在运行脚本) 之前收到编译时间错误 (错误。</span><span class="sxs-lookup"><span data-stu-id="8b48b-347">If the TypeScript compiler is unable to determine the type of a variable (either because type is not declared explicitly or type inference is not possible), then you will receive a compilation time error (error prior to running the script).</span></span> <span data-ttu-id="8b48b-348">You will see an error in the editor as well.</span><span class="sxs-lookup"><span data-stu-id="8b48b-348">You will see an error in the editor as well.</span></span>

![编辑器中显示的隐式"任何"错误](../../images/getting-started-iany.png)

<span data-ttu-id="8b48b-350">以下脚本有编译时间错误，因为声明变量时没有类型，并且 TypeScript 无法确定声明时的类型。</span><span class="sxs-lookup"><span data-stu-id="8b48b-350">The following scripts have compilation time errors because variables are declared without types and TypeScript cannot determine the type at the time of declaration.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
    // The variable 'value' gets 'any' type
    // because no type is declared.
    let value; 
    // Even when a number type is assigned,
    // the type of 'value' remains any.
    value = 10; 
    // The following statement fails because
    // Office Scripts can't send an argument
    // of type 'any' to Excel for processing.
    workbook.getActiveCell().setValue(value);
    return;
}
```

```TypeScript
function main(workbook: ExcelScript.Workbook) {
    // The variable 'cell' gets 'any' type
    // because no type is defined.
    let cell; 
    cell = workbook.getActiveCell().getValue();
    // Office Scripts can't assign Range type object
    // to a variable of 'any' type.
    console.log(cell.getValue());
    return;
}
```

<span data-ttu-id="8b48b-351">若要避免此错误，请改为使用以下模式。</span><span class="sxs-lookup"><span data-stu-id="8b48b-351">To avoid this error, use the following patterns instead.</span></span> <span data-ttu-id="8b48b-352">在每种情况下，变量及其类型都同时声明。</span><span class="sxs-lookup"><span data-stu-id="8b48b-352">In each case, the variable and its type are declared at the same time.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
    const value: number = 10; 
    workbook.getActiveCell().setValue(value);
    return;
}
```

```TypeScript
function main(workbook: ExcelScript.Workbook) {
    const cell: ExcelScript.Range = workbook.getActiveCell().getValue();
    console.log(cell.getValue()); 
    return;
}
```

## <a name="error-handling"></a><span data-ttu-id="8b48b-353">错误处理</span><span class="sxs-lookup"><span data-stu-id="8b48b-353">Error handling</span></span>

<span data-ttu-id="8b48b-354">Office 脚本错误可以分为以下类别之一。</span><span class="sxs-lookup"><span data-stu-id="8b48b-354">Office Scripts error can be classified into one of the following categories.</span></span>

1. <span data-ttu-id="8b48b-355">编辑器中显示的编译时警告</span><span class="sxs-lookup"><span data-stu-id="8b48b-355">Compile-time warning shown in the editor</span></span>
1. <span data-ttu-id="8b48b-356">编译时错误，该错误在您运行时出现，但在执行开始之前出现</span><span class="sxs-lookup"><span data-stu-id="8b48b-356">Compile-time error that appears when you run but occurs before execution begins</span></span>
1. <span data-ttu-id="8b48b-357">运行时错误</span><span class="sxs-lookup"><span data-stu-id="8b48b-357">Runtime error</span></span>

<span data-ttu-id="8b48b-358">可以在编辑器中用红色波浪下划线标识编辑器警告：</span><span class="sxs-lookup"><span data-stu-id="8b48b-358">Editor warnings can be identified using the wavy red underlines in the editor:</span></span>

![编辑器中显示的编译时警告](../../images/getting-started-eanyi.png)

<span data-ttu-id="8b48b-360">有时，还可能会看到橙色警告下划线和灰色信息性消息。</span><span class="sxs-lookup"><span data-stu-id="8b48b-360">At times, you may also see orange warning underlines and grey informational messages.</span></span> <span data-ttu-id="8b48b-361">应仔细检查它们，尽管它们不会导致错误。</span><span class="sxs-lookup"><span data-stu-id="8b48b-361">They should be examined closely though they are not going to cause errors.</span></span>

<span data-ttu-id="8b48b-362">无法区分编译时错误和运行时错误，因为两条错误消息看起来相同。</span><span class="sxs-lookup"><span data-stu-id="8b48b-362">It isn't possible to distinguish between compile-time and runtime errors as both error messages look identical.</span></span> <span data-ttu-id="8b48b-363">当您实际执行脚本时，这两者均会发生。</span><span class="sxs-lookup"><span data-stu-id="8b48b-363">They both occur when you actually execute the script.</span></span> <span data-ttu-id="8b48b-364">下图显示了编译时错误和运行时错误的示例。</span><span class="sxs-lookup"><span data-stu-id="8b48b-364">The following images show examples of a compile-time error and a runtime error.</span></span>

![编译时错误的示例](../../images/getting-started-expany.png)

![运行时错误示例](../../images/getting-started-error-basic.png)

<span data-ttu-id="8b48b-367">在这两种情况下，你将看到发生错误的行号。</span><span class="sxs-lookup"><span data-stu-id="8b48b-367">In both cases, you will see the line number where the error occurred.</span></span> <span data-ttu-id="8b48b-368">然后，你可以检查代码、修复问题，然后再次运行。</span><span class="sxs-lookup"><span data-stu-id="8b48b-368">You can then examine the code, fix the issue, and run again.</span></span>

<span data-ttu-id="8b48b-369">以下是避免运行时错误的一些最佳实践。</span><span class="sxs-lookup"><span data-stu-id="8b48b-369">Following are a few best practices to avoid runtime errors.</span></span>

### <a name="check-for-object-existence-before-deletion"></a><span data-ttu-id="8b48b-370">在删除之前检查对象是否存在</span><span class="sxs-lookup"><span data-stu-id="8b48b-370">Check for object existence before deletion</span></span>

<span data-ttu-id="8b48b-371">或者，要删除可能存在或不存在的对象，请使用此模式：</span><span class="sxs-lookup"><span data-stu-id="8b48b-371">Alternatively, for deleting an object that may or may not exist, use this pattern:</span></span>

```TypeScript
// The ? ensures that the delete() API is only invoked if the object exists.
workbook.getWorksheet('Index')?.delete();

// Alternative:
const indexSheet = workbook.getWorksheet('Index');
if (indexSheet) {
    indexSheet.delete();
}
```

### <a name="do-pre-checks-at-the-beginning-of-the-script"></a><span data-ttu-id="8b48b-372">在脚本开头进行预检查</span><span class="sxs-lookup"><span data-stu-id="8b48b-372">Do pre-checks at the beginning of the script</span></span>

<span data-ttu-id="8b48b-373">最佳做法是，在运行脚本之前，始终确保所有输入都存在于 Excel 文件中。</span><span class="sxs-lookup"><span data-stu-id="8b48b-373">As a best practice, always ensure that all your inputs are present in the Excel file prior to running your script.</span></span> <span data-ttu-id="8b48b-374">您可能已对工作簿中的对象进行了某些假设。</span><span class="sxs-lookup"><span data-stu-id="8b48b-374">You may have made certain assumptions about objects being present in the workbook.</span></span> <span data-ttu-id="8b48b-375">如果不存在这些对象，则脚本在读取对象或其数据时可能会遇到错误。</span><span class="sxs-lookup"><span data-stu-id="8b48b-375">If those objects don't exist, your script may encounter an error when you read the object or its data.</span></span> <span data-ttu-id="8b48b-376">与其在部分更新或处理完成后中间开始处理和错误，不如在脚本开头执行所有预检查。</span><span class="sxs-lookup"><span data-stu-id="8b48b-376">Rather than beginning the processing and erroring in the middle after part of the updates or processing has already finished, it is better to do all pre-checks at the start of the script.</span></span>

<span data-ttu-id="8b48b-377">例如，以下脚本需要显示名为 Table1 和 Table2 的两个表。</span><span class="sxs-lookup"><span data-stu-id="8b48b-377">For example, the following script requires two tables named Table1 and Table2 to be present.</span></span> <span data-ttu-id="8b48b-378">因此，脚本将检查其状态，并结束于语句和相应的 `return` 消息（如果不存在）。</span><span class="sxs-lookup"><span data-stu-id="8b48b-378">Hence the script checks for their presence and ends with the `return` statement and an appropriate message if they are not present.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {

  // Tables that should be in the workbook for the script to work:
  const TargetTableName = 'Table1';
  const SourceTableName = 'Table2';

  // Get the table objects.
  let targetTable = workbook.getTable(TargetTableName);
  let sourceTable = workbook.getTable(SourceTableName);

  if (!targetTable || !sourceTable) {
    console.log(`Required tables missing - Check that both source (${TargetTableName}) and target (${SourceTableName}) tables are present before running the script.`);
    return;
  }

  // Continue....
}
```

<span data-ttu-id="8b48b-379">如果验证以确保输入数据在单独的函数中发生，则通过从 函数发出 语句结束脚本 `return` `main` 非常重要。</span><span class="sxs-lookup"><span data-stu-id="8b48b-379">If the verification to ensure the presence of input data is happening in a separate function, it's important to end the script by issuing the `return` statement from the `main` function.</span></span>

<span data-ttu-id="8b48b-380">在下面的示例中， `main` 函数调用 `inputPresent` 函数以执行预检查。</span><span class="sxs-lookup"><span data-stu-id="8b48b-380">In the following example, the `main` function calls the `inputPresent` function to do the pre-checks.</span></span> <span data-ttu-id="8b48b-381">`inputPresent` 返回一个 boolean (`true` 或 `false`) 指示所有必需的输入是否全部存在。</span><span class="sxs-lookup"><span data-stu-id="8b48b-381">`inputPresent` returns a boolean (`true` or `false`) indicating whether all required inputs are present or not.</span></span> <span data-ttu-id="8b48b-382">然后，函数负责发出语句 (即从函数内) `main` `return` `main` 立即结束脚本。</span><span class="sxs-lookup"><span data-stu-id="8b48b-382">It's then the responsibility of the `main` function to issue the `return` statement (that is, from within the `main` function) to end the script immediately.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {

  // Get the table objects.
  if (!inputPresent(workbook)) {
    return;
  }

  // Continue....
}

function inputPresent( workbook: ExcelScript.Workbook): boolean {

  // Tables that should be in the workbook for the script to work:
  const TargetTableName = 'Table1';
  const SourceTableName = 'Table2';

  // Get the table objects.
  let targetTable = workbook.getTable(TargetTableName);
  let sourceTable = workbook.getTable(SourceTableName);

  if (!targetTable || !sourceTable) {
    console.log(`Required tables missing - Check that both source (${TargetTableName}) and target (${SourceTableName}) tables are present before running the script.`);
    return false;
  }
  return true;
}
```

### <a name="when-to-abort-throw-the-script"></a><span data-ttu-id="8b48b-383">何时中止 `throw` () 脚本</span><span class="sxs-lookup"><span data-stu-id="8b48b-383">When to abort (`throw`) the script</span></span>  

<span data-ttu-id="8b48b-384">大多数情况下，不需要中止脚本 () `throw` 脚本。</span><span class="sxs-lookup"><span data-stu-id="8b48b-384">For the most part, you don't need to abort (`throw`) from your script.</span></span> <span data-ttu-id="8b48b-385">这是因为脚本通常会通知用户脚本由于问题而无法运行。</span><span class="sxs-lookup"><span data-stu-id="8b48b-385">This is because the script's usually informs the user that the script failed to run due to an issue.</span></span> <span data-ttu-id="8b48b-386">在大多数情况下，用一条错误消息和函数中的语句结束脚本 `return` `main` 就足够了。</span><span class="sxs-lookup"><span data-stu-id="8b48b-386">In most case, it's sufficient to end the script with an error message and a `return` statement from the `main` function.</span></span>

<span data-ttu-id="8b48b-387">但是，如果脚本作为 Power Automate 的一部分运行，则当不满足某些条件时，您可能需要中止流。</span><span class="sxs-lookup"><span data-stu-id="8b48b-387">However, if your script is running as part of Power Automate, you may want to abort the flow if certain conditions are not met.</span></span> <span data-ttu-id="8b48b-388">因此，不要出错，而应发出语句来中止脚本，以便任何后续代码语句 `return` `throw` 不会运行，这一点很重要。</span><span class="sxs-lookup"><span data-stu-id="8b48b-388">It's therefore important to not `return` upon an error but rather issue a `throw` statement to abort the script so that any subsequent code statements don't run.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {

  // Tables that should be in the workbook for the script to work:
  const TargetTableName = 'Table1';
  const SourceTableName = 'Table2';

  // Get the table objects.
  let targetTable = workbook.getTable(TargetTableName);
  let sourceTable = workbook.getTable(SourceTableName);

  if (!targetTable || !sourceTable) {
    // Abort script.
    throw `Required tables missing - Check that both source (${TargetTableName}) and target (${SourceTableName}) tables are present before running the script.`;
  }
  
```

<span data-ttu-id="8b48b-389">如下一节所述，另一种情况是当你具有多个函数时 (调用调用 等) 这会使传播错误 `main` `functionX` `functionY` 变得困难。</span><span class="sxs-lookup"><span data-stu-id="8b48b-389">As mentioned in the following section, another scenario is when you have several functions involved (`main` calls `functionX` which calls `functionY`, etc.) which makes it hard to propagate the error.</span></span> <span data-ttu-id="8b48b-390">从带消息的嵌套函数中止/引发可能比返回错误一向多到多返回一条 `main` `main` 错误消息要容易。</span><span class="sxs-lookup"><span data-stu-id="8b48b-390">Aborting/throwing from the nested function with a message may be easier than returning an error all the way up to `main` and returning from `main` with an error message.</span></span>

### <a name="when-to-use-trycatch-throw-exception"></a><span data-ttu-id="8b48b-391">何时使用 try.。catch (throw exception) </span><span class="sxs-lookup"><span data-stu-id="8b48b-391">When to use try..catch (throw exception)</span></span>

<span data-ttu-id="8b48b-392">该技术 [`try..catch`](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Statements/try...catch) 是一种检测 API 调用是否失败的方法，并处理脚本中的该错误。</span><span class="sxs-lookup"><span data-stu-id="8b48b-392">The [`try..catch`](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Statements/try...catch) technique is a way to detect if an API call failed and handle that error in your script.</span></span> <span data-ttu-id="8b48b-393">检查 API 的返回值以验证其是否成功完成可能很重要。</span><span class="sxs-lookup"><span data-stu-id="8b48b-393">It may be important to check the return value of an API to verify that it was completed successfully.</span></span>

<span data-ttu-id="8b48b-394">请考虑以下示例代码段。</span><span class="sxs-lookup"><span data-stu-id="8b48b-394">Consider the following example snippet.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {

  // Somewhere in the script, perform a large data update.
  range.setValues(someLargeValues);

}
```

<span data-ttu-id="8b48b-395">`setValues()`调用可能会失败，并会导致脚本失败。</span><span class="sxs-lookup"><span data-stu-id="8b48b-395">The `setValues()` call may fail and result in the script failure.</span></span> <span data-ttu-id="8b48b-396">你可能希望在代码中处理此条件，并可能自定义错误消息或将更新分解为较小的单元等。在这种情况下，了解 API 返回错误并解释或处理该错误很重要。</span><span class="sxs-lookup"><span data-stu-id="8b48b-396">You may wish to handle this condition in your code and perhaps customize the error message or break up the update into smaller units, etc. In that case, it's important to know that the API returned an error and interpret or handle that error.</span></span>

```TypeScript
try {
    range.setValues(someLargeValues);
} catch (error) {
    console.log(`The script failed to update the values at location ____. Please inspect and run again.`);
    console.log(error);
    return; // End script (assuming this is in main function).
}

// OR...

try {
    range.setValues(someLargeValues);
} catch (error) {
    console.log(`The script failed to update the values at location ____. Trying a different approach`);
    handleUpdatesInSmallerChunks(someLargeValues);
}

// Continue...
}
```

<span data-ttu-id="8b48b-397">另一种情况是 main 函数调用另一个函数，而该函数又调用另一个 (函数，等等。) ，并且你关心的 API 调用在底部函数中向下发生。</span><span class="sxs-lookup"><span data-stu-id="8b48b-397">Another scenario is when main function calls another function, which in turn calls another function (and so on..), and the API call that you care about happens down in the bottom function.</span></span> <span data-ttu-id="8b48b-398">将错误一路传播到上可能不可行 `main` 或不方便。</span><span class="sxs-lookup"><span data-stu-id="8b48b-398">Propagating the error all the way up to `main` may not be feasible or convenient.</span></span> <span data-ttu-id="8b48b-399">在这种情况下，在底部函数中引发错误将最为方便。</span><span class="sxs-lookup"><span data-stu-id="8b48b-399">In that case, throwing an error in the bottom function will be most convenient.</span></span>

```TypeScript

function main(workbook: ExcelScript.Workbook) {
    ...
    updateRangeInChunks(sheet.getRange("B1"), data);
    ...
}

function updateRangeInChunks(
    ...
    updateNextChunk(startCell, values, rowsPerChunk, totalRowsUpdated);
    ...
}

function updateTargetRange(
      targetCell: ExcelScript.Range,
      values: (string | boolean | number)[][]
    ) {
    const targetRange = targetCell.getResizedRange(values.length - 1, values[0].length - 1);
    console.log(`Updating the range: ${targetRange.getAddress()}`);
    try {
      targetRange.setValues(values);
    } catch (e) {
      throw `Error while updating the whole range: ${JSON.stringify(e)}`;
    }
    return;
}
```

<span data-ttu-id="8b48b-400">*警告*： `try..catch` 在循环内使用会降低脚本的速度。</span><span class="sxs-lookup"><span data-stu-id="8b48b-400">*Warning*: Using `try..catch` inside of a loop will slow down your script.</span></span> <span data-ttu-id="8b48b-401">避免在循环内部或循环内使用它。</span><span class="sxs-lookup"><span data-stu-id="8b48b-401">Avoid using this inside of or around loops.</span></span>

## <a name="range-basics"></a><span data-ttu-id="8b48b-402">Range 基础知识</span><span class="sxs-lookup"><span data-stu-id="8b48b-402">Range basics</span></span>

<span data-ttu-id="8b48b-403">在继续您的旅程之前，请查看 Range [Basics。](range-basics.md)</span><span class="sxs-lookup"><span data-stu-id="8b48b-403">Check out [Range Basics](range-basics.md) before you go further on your journey.</span></span>

## <a name="basic-performance-considerations"></a><span data-ttu-id="8b48b-404">基本性能注意事项</span><span class="sxs-lookup"><span data-stu-id="8b48b-404">Basic performance considerations</span></span>

### <a name="avoid-slow-operations-in-the-loop"></a><span data-ttu-id="8b48b-405">避免循环中运行缓慢的操作</span><span class="sxs-lookup"><span data-stu-id="8b48b-405">Avoid slow operations in the loop</span></span>

<span data-ttu-id="8b48b-406">在循环语句内部/周围执行某些操作（如 、 等） `for` `for..of` `map` `forEach` 会导致性能变慢。</span><span class="sxs-lookup"><span data-stu-id="8b48b-406">Certain operations when done inside/around the loop statements such as `for`, `for..of`, `map`, `forEach`, etc. can lead to slow performance.</span></span> <span data-ttu-id="8b48b-407">避免以下 API 类别。</span><span class="sxs-lookup"><span data-stu-id="8b48b-407">Avoid the following API categories.</span></span>

* <span data-ttu-id="8b48b-408">`get*` API</span><span class="sxs-lookup"><span data-stu-id="8b48b-408">`get*` APIs</span></span>

<span data-ttu-id="8b48b-409">读取循环之外所需的全部数据，而不是在循环内读取数据。</span><span class="sxs-lookup"><span data-stu-id="8b48b-409">Read all the data you need outside of the loop rather than reading it inside of the loop.</span></span> <span data-ttu-id="8b48b-410">有时，很难避免在循环内读取;在这种情况下，请确保循环计数不是太大或分批管理它们，以避免必须循环访问大型数据结构。</span><span class="sxs-lookup"><span data-stu-id="8b48b-410">At times, it is hard to avoid reading inside of loops; in such a case, make sure your loop counts are not too large or manage them in batches to avoid having to loop through a large data structure.</span></span>

<span data-ttu-id="8b48b-411">注意：如果你处理的范围/数据非常大 (假设 >100，000 个单元格) ，你可能需要使用高级技术，例如将读取/写入分成多个区块。</span><span class="sxs-lookup"><span data-stu-id="8b48b-411">**Note**: If the range/data you are dealing with is quite large (say >100K cells), you may need to use advanced techniques like breaking up your read/writes into multiple chunks.</span></span> <span data-ttu-id="8b48b-412">以下视频确实适用于中小型数据设置。</span><span class="sxs-lookup"><span data-stu-id="8b48b-412">The following video is really for a small-mid sized data setup.</span></span> <span data-ttu-id="8b48b-413">对于大型数据集，请参阅 [高级数据写入方案](write-large-dataset.md)。</span><span class="sxs-lookup"><span data-stu-id="8b48b-413">For a large dataset, refer to [advanced data write scenario](write-large-dataset.md).</span></span>

<span data-ttu-id="8b48b-414">[![提供读写优化提示的视频](../../images/getting-started-v_perf.jpg)](https://youtu.be/lsR_GvVW3Pg "显示读写优化提示的视频")</span><span class="sxs-lookup"><span data-stu-id="8b48b-414">[![Video providing a read-and-write optimization tip](../../images/getting-started-v_perf.jpg)](https://youtu.be/lsR_GvVW3Pg "Video showing read-and-write optimization tip")</span></span>

* <span data-ttu-id="8b48b-415">`console.log` 语句 (请参阅以下示例) </span><span class="sxs-lookup"><span data-stu-id="8b48b-415">`console.log` statement (see the following example)</span></span>

```TypeScript
// Color each cell with random color.
for (let row = 0; row < rows; row++) {
    for (let col = 0; col < cols; col++) {
        range
            .getCell(row, col)
            .getFormat()
            .getFill()
            .setColor(`#${Math.random().toString(16).substr(-6)}`);
        /* Avoid such console.log inside loop */
        // console.log("Updating" + range.getCell(row, col).getAddress());
    }
}
```

* <span data-ttu-id="8b48b-416">`try {} catch ()` 语句</span><span class="sxs-lookup"><span data-stu-id="8b48b-416">`try {} catch ()` statement</span></span>

<span data-ttu-id="8b48b-417">避免异常 `for` 处理循环。</span><span class="sxs-lookup"><span data-stu-id="8b48b-417">Avoid exception handling `for` loops.</span></span> <span data-ttu-id="8b48b-418">内部循环和外部循环。</span><span class="sxs-lookup"><span data-stu-id="8b48b-418">Both inside and outside loops.</span></span>

## <a name="note-to-vba-developers"></a><span data-ttu-id="8b48b-419">VBA 开发人员注意事项</span><span class="sxs-lookup"><span data-stu-id="8b48b-419">Note to VBA developers</span></span>

<span data-ttu-id="8b48b-420">TypeScript 语言在语法和命名约定方面均与 VBA 不同。</span><span class="sxs-lookup"><span data-stu-id="8b48b-420">The TypeScript language differs from VBA both syntactically as well as in naming conventions.</span></span>

<span data-ttu-id="8b48b-421">查看以下等效代码段。</span><span class="sxs-lookup"><span data-stu-id="8b48b-421">Check out the following equivalent snippets.</span></span>

```vba
Worksheets("Sheet1").Range("A1:G37").Clear
```

```TypeScript
workbook.getWorksheet('Sheet1').getRange('A1:G37').clear(ExcelScript.ClearApplyTo.all);
```

<span data-ttu-id="8b48b-422">有关 TypeScript 的一些说明：</span><span class="sxs-lookup"><span data-stu-id="8b48b-422">A few things to call out about TypeScript:</span></span>

* <span data-ttu-id="8b48b-423">您可能会注意到，所有方法都需要有要执行的开放式关闭括号。</span><span class="sxs-lookup"><span data-stu-id="8b48b-423">You may notice that all methods need to have open-close parentheses to execute.</span></span> <span data-ttu-id="8b48b-424">参数的传递方式相同，但某些参数可能需要执行 (，即必需参数与可选) 。</span><span class="sxs-lookup"><span data-stu-id="8b48b-424">Arguments are passed identically but some arguments may be required for execution (that is, required vs optional).</span></span>
* <span data-ttu-id="8b48b-425">命名约定遵循 camelCase 而不是 PascalCase 约定。</span><span class="sxs-lookup"><span data-stu-id="8b48b-425">The naming convention follows camelCase instead of PascalCase convention.</span></span>
* <span data-ttu-id="8b48b-426">方法通常 `get` 具有 `set` 或 前缀，用于指示它是读取还是写入对象成员。</span><span class="sxs-lookup"><span data-stu-id="8b48b-426">Methods usually have `get` or `set` prefixes indicating whether it is reading or writing object members.</span></span>
* <span data-ttu-id="8b48b-427">代码块由左大括号定义和标识 `{` `}` ：。</span><span class="sxs-lookup"><span data-stu-id="8b48b-427">The code blocks are defined and identified by open-close curly braces: `{` `}`.</span></span> <span data-ttu-id="8b48b-428">条件、语句、 `if` `while` `for` 循环、函数定义等需要块。</span><span class="sxs-lookup"><span data-stu-id="8b48b-428">Blocks are required for `if` conditions, `while` statements, `for` loops, function definitions, etc.</span></span>
* <span data-ttu-id="8b48b-429">函数可以调用其他函数，甚至可以在函数中定义函数。</span><span class="sxs-lookup"><span data-stu-id="8b48b-429">Functions can call other functions and you can even define functions within a function.</span></span>

<span data-ttu-id="8b48b-430">总的来说，TypeScript 是一种不同的语言，它们之间有几个相似之处。</span><span class="sxs-lookup"><span data-stu-id="8b48b-430">Overall, TypeScript is a different language and there are few similarities between them.</span></span> <span data-ttu-id="8b48b-431">但是，Office 脚本 API 本身使用类似的术语和数据模型 (对象模型) 层次结构作为 VBA API，这应该可以帮助您四处导航。</span><span class="sxs-lookup"><span data-stu-id="8b48b-431">However, the Office Scripts API themselves use similar terminology and data-model (object model) hierarchy as VBA APIs and that should help you navigate around.</span></span>
