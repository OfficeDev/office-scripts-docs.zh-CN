---
title: Office脚本代码编辑器环境
description: Excel web 版 中脚本Office的先决条件和环境Excel web 版。
ms.date: 05/10/2021
localization_priority: Normal
ms.openlocfilehash: aa54939826f8dda2a068df0f3fabf0fd3a2c842b
ms.sourcegitcommit: 4687693f02fc90a57ba30c461f35046e02e6f5fb
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 05/19/2021
ms.locfileid: "52545820"
---
# <a name="office-scripts-code-editor-environment"></a><span data-ttu-id="49768-103">Office脚本代码编辑器环境</span><span class="sxs-lookup"><span data-stu-id="49768-103">Office Scripts Code Editor environment</span></span>

<span data-ttu-id="49768-104">Office脚本使用 TypeScript 或 JavaScript 编写，并使用 Office 脚本 JavaScript API 与 Excel 工作簿进行交互。</span><span class="sxs-lookup"><span data-stu-id="49768-104">Office Scripts are written in either TypeScript or JavaScript and use the Office Scripts JavaScript APIs to interact with an Excel workbook.</span></span> <span data-ttu-id="49768-105">代码编辑器基于Visual Studio Code，因此如果你之前使用过该环境，则感觉像在家一样。</span><span class="sxs-lookup"><span data-stu-id="49768-105">The Code Editor is based on Visual Studio Code, so if you've used that environment before, you'll feel right at home.</span></span>

## <a name="scripting-language-typescript-or-javascript"></a><span data-ttu-id="49768-106">脚本语言：TypeScript 或 JavaScript</span><span class="sxs-lookup"><span data-stu-id="49768-106">Scripting language: TypeScript or JavaScript</span></span>

<span data-ttu-id="49768-107">Office脚本是用[TypeScript](https://www.typescriptlang.org/docs/home.html)编写的，它是[JavaScript](https://developer.mozilla.org/docs/Web/JavaScript)的超集。</span><span class="sxs-lookup"><span data-stu-id="49768-107">Office Scripts are written in [TypeScript](https://www.typescriptlang.org/docs/home.html), which is a superset of [JavaScript](https://developer.mozilla.org/docs/Web/JavaScript).</span></span> <span data-ttu-id="49768-108">操作录制器在 TypeScript 中生成代码，Office脚本文档使用 TypeScript。</span><span class="sxs-lookup"><span data-stu-id="49768-108">The Action Recorder generates code in TypeScript and the Office Scripts documentation uses TypeScript.</span></span> <span data-ttu-id="49768-109">由于 TypeScript 是 JavaScript 的超集，因此在 JavaScript 中编写的任何脚本代码都运行正常。</span><span class="sxs-lookup"><span data-stu-id="49768-109">Since TypeScript is a superset of JavaScript, any scripting code that you write in JavaScript will work just fine.</span></span>

<span data-ttu-id="49768-110">Office脚本在很大程度上是自包含的代码片段。</span><span class="sxs-lookup"><span data-stu-id="49768-110">Office Scripts are largely self-contained pieces of code.</span></span> <span data-ttu-id="49768-111">只使用了 TypeScript 功能的一小部分。</span><span class="sxs-lookup"><span data-stu-id="49768-111">Only a small part of TypeScript's functionality is used.</span></span> <span data-ttu-id="49768-112">因此，您可以编辑脚本，而无需了解 TypeScript 的不一样。</span><span class="sxs-lookup"><span data-stu-id="49768-112">Therefore, you can edit scripts without having to learn the intricacies of TypeScript.</span></span> <span data-ttu-id="49768-113">代码编辑器还处理代码的安装、编译和执行，因此，你无需担心脚本本身。</span><span class="sxs-lookup"><span data-stu-id="49768-113">The Code Editor also handles the installation, compilation, and execution of code, so you don't need to worry about anything but the script itself.</span></span> <span data-ttu-id="49768-114">可以学习语言并创建脚本，而无需以前的编程知识。</span><span class="sxs-lookup"><span data-stu-id="49768-114">It's possible to learn the language and create scripts without previous programming knowledge.</span></span> <span data-ttu-id="49768-115">但是，如果你对编程很新，我们建议先学习一些基础，然后再继续Office脚本：</span><span class="sxs-lookup"><span data-stu-id="49768-115">However, if you're new to programming, we recommend learning some fundamentals before proceeding with Office Scripts:</span></span>

[!INCLUDE [Preview note](../includes/coding-basics-references.md)]

## <a name="office-scripts-javascript-api"></a><span data-ttu-id="49768-116">Office脚本 JavaScript API</span><span class="sxs-lookup"><span data-stu-id="49768-116">Office Scripts JavaScript API</span></span>

<span data-ttu-id="49768-117">Office脚本使用适用于加载项的 Office JavaScript API [Office版本](/office/dev/add-ins/overview/index)。虽然两个 API 存在相似之处，但不应假定代码可以在两个平台之间移植。</span><span class="sxs-lookup"><span data-stu-id="49768-117">Office Scripts use a specialized version of the Office JavaScript APIs for [Office Add-ins](/office/dev/add-ins/overview/index). While there are similarities in the two APIs, you should not assume code can be ported between the two platforms.</span></span> <span data-ttu-id="49768-118">The differences between the two platforms are described in the [Differences between Office Scripts and Office Add-ins](../resources/add-ins-differences.md#apis) article.</span><span class="sxs-lookup"><span data-stu-id="49768-118">The differences between the two platforms are described in the [Differences between Office Scripts and Office Add-ins](../resources/add-ins-differences.md#apis) article.</span></span> <span data-ttu-id="49768-119">可以在脚本 API 参考文档中查看可用于Office[的所有 API。](/javascript/api/office-scripts/overview)</span><span class="sxs-lookup"><span data-stu-id="49768-119">You can view all the APIs available to your script in the [Office Scripts API reference documentation](/javascript/api/office-scripts/overview).</span></span>

## <a name="external-library-support"></a><span data-ttu-id="49768-120">外部库支持</span><span class="sxs-lookup"><span data-stu-id="49768-120">External library support</span></span>

<span data-ttu-id="49768-121">Office脚本不支持使用外部第三方 JavaScript 库。</span><span class="sxs-lookup"><span data-stu-id="49768-121">Office Scripts does not support the usage of external, third-party JavaScript libraries.</span></span> <span data-ttu-id="49768-122">目前，您无法从脚本调用 Office脚本 API 外的任何库。</span><span class="sxs-lookup"><span data-stu-id="49768-122">Currently, you cannot call any library other than the Office Scripts APIs from a script.</span></span> <span data-ttu-id="49768-123">你仍然可以访问任何内置的 [JavaScript](../develop/javascript-objects.md)对象，如 [数学](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Math)。</span><span class="sxs-lookup"><span data-stu-id="49768-123">You do still have access to any [built-in JavaScript object](../develop/javascript-objects.md), such as [Math](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Math).</span></span>

## <a name="intellisense"></a><span data-ttu-id="49768-124">IntelliSense</span><span class="sxs-lookup"><span data-stu-id="49768-124">IntelliSense</span></span>

<span data-ttu-id="49768-125">IntelliSense是一种代码编辑器功能，可帮助您在编辑脚本时防止拼写错误和语法错误。</span><span class="sxs-lookup"><span data-stu-id="49768-125">IntelliSense is a Code Editor feature that helps prevent typos and syntax errors as you edit your script.</span></span> <span data-ttu-id="49768-126">它显示在键入时可能的对象和字段名称，以及每个 API 的内联文档。</span><span class="sxs-lookup"><span data-stu-id="49768-126">It displays possible object and field names as you type, as well as inline documentation for every API.</span></span>

<span data-ttu-id="49768-127">代码Excel编辑器使用与IntelliSense相同的Visual Studio Code。</span><span class="sxs-lookup"><span data-stu-id="49768-127">The Excel Code Editor uses the same IntelliSense engine as Visual Studio Code.</span></span> <span data-ttu-id="49768-128">若要了解有关该功能的更多信息，请访问[Visual Studio Code的IntelliSense功能。](https://code.visualstudio.com/docs/editor/intellisense#_intellisense-features)</span><span class="sxs-lookup"><span data-stu-id="49768-128">To learn more about the feature, visit [Visual Studio Code's IntelliSense Features](https://code.visualstudio.com/docs/editor/intellisense#_intellisense-features).</span></span>

## <a name="keyboard-shortcuts"></a><span data-ttu-id="49768-129">键盘快捷方式</span><span class="sxs-lookup"><span data-stu-id="49768-129">Keyboard shortcuts</span></span>

<span data-ttu-id="49768-130">大多数用于自定义脚本的Visual Studio Code也可在 Office 脚本代码编辑器中运行。</span><span class="sxs-lookup"><span data-stu-id="49768-130">Most of the keyboard shortcuts for Visual Studio Code also work in the Office Scripts Code Editor.</span></span> <span data-ttu-id="49768-131">使用以下 PDF 了解可用选项并充分利用代码编辑器：</span><span class="sxs-lookup"><span data-stu-id="49768-131">Use the following PDFs to learn about the available options and get the most out of the Code Editor:</span></span>

- <span data-ttu-id="49768-132">[macOS 的键盘快捷方式](https://code.visualstudio.com/shortcuts/keyboard-shortcuts-macos.pdf)。</span><span class="sxs-lookup"><span data-stu-id="49768-132">[Keyboard shortcuts for macOS](https://code.visualstudio.com/shortcuts/keyboard-shortcuts-macos.pdf).</span></span>
- <span data-ttu-id="49768-133">[键盘快捷方式的 Windows](https://code.visualstudio.com/shortcuts/keyboard-shortcuts-windows.pdf)。</span><span class="sxs-lookup"><span data-stu-id="49768-133">[Keyboard shortcuts for Windows](https://code.visualstudio.com/shortcuts/keyboard-shortcuts-windows.pdf).</span></span>

## <a name="see-also"></a><span data-ttu-id="49768-134">另请参阅</span><span class="sxs-lookup"><span data-stu-id="49768-134">See also</span></span>

- [<span data-ttu-id="49768-135">Office 脚本 API 参考</span><span class="sxs-lookup"><span data-stu-id="49768-135">Office Scripts API reference</span></span>](/javascript/api/office-scripts/overview)
- [<span data-ttu-id="49768-136">Office 脚本疑难解答</span><span class="sxs-lookup"><span data-stu-id="49768-136">Troubleshooting Office Scripts</span></span>](../testing/troubleshooting.md)
- [<span data-ttu-id="49768-137">在 Office 脚本中使用内置的 JavaScript 对象</span><span class="sxs-lookup"><span data-stu-id="49768-137">Using built-in JavaScript objects in Office Scripts</span></span>](../develop/javascript-objects.md)
