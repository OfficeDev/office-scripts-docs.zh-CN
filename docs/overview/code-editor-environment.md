---
title: Office 脚本代码编辑器环境
description: Excel web 版 中 Office 脚本的先决条件和环境信息。
ms.date: 11/08/2022
ms.localizationpriority: medium
ms.openlocfilehash: a5a7601285553b1da4001a1870b6120f21bf5f2c
ms.sourcegitcommit: 7cadf2b637bf62874e43b6e595286101816662aa
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 11/09/2022
ms.locfileid: "68891251"
---
# <a name="office-scripts-code-editor-environment"></a>Office 脚本代码编辑器环境

Office 脚本以 TypeScript 或 JavaScript 编写，并使用 Office 脚本 JavaScript API 与 Excel 工作簿交互。 代码编辑器基于Visual Studio Code，因此，如果你以前使用过该环境，你会感到很自在。

> [!TIP]
> 如果熟悉Visual Studio Code，现在可以使用它来编写脚本。 访问 [office 脚本 (预览) Visual Studio Code](../develop/vscode-for-scripts.md)试用此功能。

## <a name="scripting-language-typescript-or-javascript"></a>脚本语言：TypeScript 或 JavaScript

Office 脚本以 [TypeScript](https://www.typescriptlang.org/docs/home.html) 编写，它是 [JavaScript](https://developer.mozilla.org/docs/Web/JavaScript) 的一个超集。 操作记录器在 TypeScript 中生成代码，Office 脚本文档使用 TypeScript。 由于 TypeScript 是 JavaScript 的超集，因此在 JavaScript 中编写的任何脚本代码都将正常工作。

Office 脚本主要是独立的代码片段。 仅使用 TypeScript 功能的一小部分。 因此，无需了解 TypeScript 的复杂之处即可编辑脚本。 代码编辑器还处理代码的安装、编译和执行，因此除了脚本本身之外，你无需担心任何内容。 无需以前的编程知识即可学习语言并创建脚本。 但是，如果你不熟悉编程，我们建议先学习一些基础知识，然后再继续学习 Office 脚本：

[!INCLUDE [Recommended coding resources](../includes/coding-basics-references.md)]

## <a name="office-scripts-javascript-api"></a>Office 脚本 JavaScript API

Office 脚本使用 Office [加载项](/office/dev/add-ins/overview/index)的 Office JavaScript API 的专用版本。虽然这两个 API 存在相似之处，但不应假设代码可以在两个平台之间移植。 [Office 脚本和 Office 加载项](../resources/add-ins-differences.md#apis)之间的差异一文介绍了这两个平台之间的差异。 可以在 Office 脚本 API [参考文档中](/javascript/api/office-scripts/overview)查看脚本可用的所有 API。

## <a name="external-library-support"></a>外部库支持

Office 脚本不支持使用外部第三方 JavaScript 库。 目前，不能从脚本调用除 Office 脚本 API 以外的任何库。 你仍有权访问任何 [内置 JavaScript 对象](../develop/javascript-objects.md)，例如 [Math](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Math)。

## <a name="intellisense"></a>IntelliSense

IntelliSense 是一组代码编辑器功能，可帮助你编写代码。 它提供自动完成、语法错误突出显示和内联 API 文档。

IntelliSense 在键入时提供建议，类似于 Excel 中的建议文本。 按 Tab 或 Enter 键将插入建议的成员。 按 Ctrl+空格键在当前光标位置触发 IntelliSense。 这些建议在完成方法时特别有用。 IntelliSense 显示的方法签名包含它所需的参数列表、每个参数的类型、给定参数是必需参数还是可选参数以及方法的返回类型。

将光标悬停在方法、类或其他代码对象上可查看详细信息。 将鼠标悬停在语法错误或代码建议（由红色或黄色波浪线表示）上，以查看有关如何解决问题的建议。 通常，IntelliSense 提供“快速修复”选项来自动更改代码。

:::image type="content" source="../images/implicit-any-editor-message.png" alt-text="代码编辑器的悬停文本中带有“快速修复”按钮的错误消息。":::

Office 脚本代码编辑器使用与 Visual Studio Code 相同的 IntelliSense 引擎。 若要了解有关该功能的详细信息，请访问 [Visual Studio Code 的 IntelliSense 功能](https://code.visualstudio.com/docs/editor/intellisense#_intellisense-features)。

## <a name="keyboard-shortcuts"></a>键盘快捷方式

Visual Studio Code的大多数键盘快捷方式也适用于 Office 脚本代码编辑器。 使用以下 PDF 了解可用选项并充分利用代码编辑器：

- [macOS 的键盘快捷方式](https://code.visualstudio.com/shortcuts/keyboard-shortcuts-macos.pdf)。
- [适用于 Windows 的键盘快捷方式](https://code.visualstudio.com/shortcuts/keyboard-shortcuts-windows.pdf)。

## <a name="see-also"></a>另请参阅

- [Office 脚本 API 参考](/javascript/api/office-scripts/overview)
- [Office 脚本疑难解答](../testing/troubleshooting.md)
- [在 Office 脚本中使用内置的 JavaScript 对象](../develop/javascript-objects.md)
- [office 脚本Visual Studio Code (预览版) ](../develop/vscode-for-scripts.md)
