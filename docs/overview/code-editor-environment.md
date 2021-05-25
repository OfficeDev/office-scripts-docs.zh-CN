---
title: Office脚本代码编辑器环境
description: Excel web 版 中脚本Office的先决条件和环境Excel web 版。
ms.date: 05/24/2021
localization_priority: Normal
ms.openlocfilehash: aca97c31ba970617a9fa270021a5b5b976ae4a57
ms.sourcegitcommit: 90ca8cdf30f2065f63938f6bb6780d024c128467
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 05/25/2021
ms.locfileid: "52639878"
---
# <a name="office-scripts-code-editor-environment"></a>Office脚本代码编辑器环境

Office脚本使用 TypeScript 或 JavaScript 编写，并使用 Office 脚本 JavaScript API 与 Excel 工作簿进行交互。 代码编辑器基于Visual Studio Code，因此如果你之前使用过该环境，则感觉像在家一样。

## <a name="scripting-language-typescript-or-javascript"></a>脚本语言：TypeScript 或 JavaScript

Office 脚本以 [TypeScript](https://www.typescriptlang.org/docs/home.html) 编写，它是 [JavaScript](https://developer.mozilla.org/docs/Web/JavaScript) 的一个超集。 操作录制器在 TypeScript 中生成代码，Office脚本文档使用 TypeScript。 由于 TypeScript 是 JavaScript 的超集，因此在 JavaScript 中编写的任何脚本代码都运行正常。

Office脚本在很大程度上是自包含的代码片段。 只使用了 TypeScript 功能的一小部分。 因此，您可以编辑脚本，而无需了解 TypeScript 的不一样。 代码编辑器还处理代码的安装、编译和执行，因此，你无需担心脚本本身。 可以学习语言并创建脚本，而无需以前的编程知识。 但是，如果你对编程很新，我们建议先学习一些基础，然后再继续Office脚本：

[!INCLUDE [Recommended coding resources](../includes/coding-basics-references.md)]

## <a name="office-scripts-javascript-api"></a>Office脚本 JavaScript API

Office脚本使用适用于加载项的 Office JavaScript API [Office版本](/office/dev/add-ins/overview/index)。虽然两个 API 存在相似之处，但不应假定代码可以在两个平台之间移植。 The differences between the two platforms are described in the [Differences between Office Scripts and Office Add-ins](../resources/add-ins-differences.md#apis) article. 可以在脚本 API 参考文档中查看可用于Office[的所有 API。](/javascript/api/office-scripts/overview)

## <a name="external-library-support"></a>外部库支持

Office脚本不支持使用外部第三方 JavaScript 库。 目前，您无法从脚本调用 Office脚本 API 外的任何库。 你仍然可以访问任何内置的 [JavaScript](../develop/javascript-objects.md)对象，如 [数学](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Math)。

## <a name="intellisense"></a>IntelliSense

IntelliSense是一种代码编辑器功能，可帮助您在编辑脚本时防止拼写错误和语法错误。 它显示在键入时可能的对象和字段名称，以及每个 API 的内联文档。

代码Excel编辑器使用与IntelliSense相同的Visual Studio Code。 若要了解有关该功能的更多信息，请访问[Visual Studio Code的IntelliSense功能。](https://code.visualstudio.com/docs/editor/intellisense#_intellisense-features)

## <a name="keyboard-shortcuts"></a>键盘快捷方式

大多数用于自定义脚本的Visual Studio Code也可在 Office 脚本代码编辑器中运行。 使用以下 PDF 了解可用选项并充分利用代码编辑器：

- [macOS 的键盘快捷方式](https://code.visualstudio.com/shortcuts/keyboard-shortcuts-macos.pdf)。
- [键盘快捷方式的 Windows](https://code.visualstudio.com/shortcuts/keyboard-shortcuts-windows.pdf)。

## <a name="see-also"></a>另请参阅

- [Office 脚本 API 参考](/javascript/api/office-scripts/overview)
- [Office 脚本疑难解答](../testing/troubleshooting.md)
- [在 Office 脚本中使用内置的 JavaScript 对象](../develop/javascript-objects.md)
