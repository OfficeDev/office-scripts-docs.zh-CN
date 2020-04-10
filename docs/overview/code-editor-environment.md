---
title: Office 脚本代码编辑器环境
description: Excel 网页版中 Office 脚本的先决条件和环境信息。
ms.date: 04/08/2020
localization_priority: Normal
ms.openlocfilehash: 6b26adf886172f085980bed0488b4aa7a6815991
ms.sourcegitcommit: b13dedb5ee2048f0a244aa2294bf2c38697cb62c
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 04/10/2020
ms.locfileid: "43215264"
---
# <a name="office-scripts-code-editor-environment"></a>Office 脚本代码编辑器环境

Office 脚本在[TypeScript 或 JavaScript](#scripting-language-typescript-or-javascript)中编写，并使用[Office 脚本 JavaScript api](#office-scripts-javascript-api)与 Excel 工作簿进行交互。

## <a name="scripting-language-typescript-or-javascript"></a>脚本语言： TypeScript 或 JavaScript

Office 脚本是用[TypeScript](https://www.typescriptlang.org/docs/home.html)或[JavaScript](https://developer.mozilla.org/docs/Web/JavaScript)编写的。 操作记录器在 TypeScript （JavaScript 的超集）中生成代码。 Office 脚本文档使用 TypeScript，但如果您更熟悉 JavaScript，则可以改为使用。

Office 脚本主要是自包含的代码段。 仅使用 TypeScript 的功能的一小部分。 因此，您可以编辑脚本，而无需了解 TypeScript 的复杂性。 代码编辑器还处理安装、编译和代码的执行，因此您无需担心脚本本身。 您可以了解语言并创建脚本，而无需以前的编程知识。 但是，如果您是编程新手，我们建议您先了解一些基础知识，然后再继续使用 Office 脚本：

- 了解 JavaScript 的基础知识。 您应熟悉像变量、控制流、函数和数据类型这样的概念。 [Mozilla 提供了有关 JavaScript 的一个完善的综合性教程](https://developer.mozilla.org/docs/Web/JavaScript/Guide/Introduction)。
- 了解 TypeScript 中的类型。 通过确保在编译时使用正确的类型进行方法调用和分配，可以在 JavaScript 上构建 TypeScript。 有关[接口](https://www.typescriptlang.org/docs/handbook/interfaces.html)、[类](https://www.typescriptlang.org/docs/handbook/classes.html)、[类型推理](https://www.typescriptlang.org/docs/handbook/type-inference.html)和[类型兼容性](https://www.typescriptlang.org/docs/handbook/type-compatibility.html)的 TypeScript 文档将是最有用的。

## <a name="office-scripts-javascript-api"></a>Office 脚本 JavaScript API

Office 脚本使用专用版本[Office 外接程序](/office/dev/add-ins/overview/index)使用的 Office JavaScript api。[Office 脚本与 Office 外接程序一](../resources/add-ins-differences.md#apis)文中的区别介绍了两个平台之间的差异。 您可以在[Office 脚本 API 参考文档](/javascript/api/office-scripts/overview)中查看脚本的所有可用 api。

## <a name="intellisense"></a>IntelliSense

智能感知是一项代码编辑器功能，可帮助防止在编辑脚本时键入错误和语法错误。 它在您键入时显示可能的对象和字段名称，以及每个 API 的内联文档。

Excel 代码编辑器使用与 Visual Studio Code 相同的智能感知引擎。 若要了解有关此功能的详细信息，请访问[Visual Studio Code 的智能感知功能](https://code.visualstudio.com/docs/editor/intellisense#_intellisense-features)。

## <a name="external-library-support"></a>外部库支持

Office 脚本不支持使用外部第三方 JavaScript 库。 您当前无法从脚本调用 Office 脚本 Api 之外的任何其他库。 您仍有权访问任何[内置 JavaScript 对象](../develop/javascript-objects.md)，如[数学](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Math)。

## <a name="see-also"></a>另请参阅

- [Office 脚本 API 参考](/javascript/api/office-scripts/overview)
- [Office 脚本疑难解答](../testing/troubleshooting.md)
- [在 Office 脚本中使用内置的 JavaScript 对象](../develop/javascript-objects.md)
