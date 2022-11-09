---
title: Office 脚本的平台限制和要求
description: 与 Excel web 版 一起使用时，对 Office 脚本的资源限制和浏览器支持。
ms.date: 11/07/2022
ms.localizationpriority: medium
ms.openlocfilehash: 764d1eddaf303a941a098ec1d3f3056d63e8693f
ms.sourcegitcommit: 7cadf2b637bf62874e43b6e595286101816662aa
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 11/09/2022
ms.locfileid: "68891244"
---
# <a name="platform-limits-and-requirements-with-office-scripts"></a>Office 脚本的平台限制和要求

开发 Office 脚本时，应注意一些平台限制。 本文详细介绍了适用于Excel web 版的 Office 脚本的浏览器支持和数据限制。

## <a name="browser-support"></a>浏览器支持

Office 脚本适用于[支持Office 网页版](https://support.microsoft.com/office/ad1303e0-a318-47aa-b409-d3a5eb44e452)的任何浏览器。 但是，Internet Explorer 11 (IE 11) 不支持某些 JavaScript 功能。 [ES6 或更高版本](https://www.w3schools.com/Js/js_es6.asp)中引入的任何功能都不适用于 IE 11。 如果组织中的人员仍在使用该浏览器，请确保在共享脚本时在该环境中测试脚本。

[!INCLUDE [Teams support note](../includes/teams-support-note.md)]

### <a name="third-party-cookies"></a>第三方 Cookie

浏览器需要启用第三方 Cookie 才能在 Excel web 版 中显示“**自动”** 选项卡。 如果未显示选项卡，请检查浏览器设置。 如果使用专用浏览器会话，则每次可能需要重新启用此设置。

> [!NOTE]
> 某些浏览器将此设置称为“所有 Cookie”，而不是“第三方 Cookie”。

#### <a name="instructions-for-adjusting-cookie-settings-in-popular-browsers"></a>有关在常用浏览器中调整 Cookie 设置的说明

- [Chrome](https://support.google.com/chrome/answer/95647)
- [Microsoft Edge](https://support.microsoft.com/microsoft-edge/597f04f2-c0ce-f08c-7c2b-541086362bd2)
- [Firefox](https://support.mozilla.org/kb/disable-third-party-cookies)
- [Safari](https://support.apple.com/guide/safari/manage-cookies-and-website-data-sfri11471/mac)

## <a name="data-limits"></a>数据限制

一次可以传输的 Excel 数据量以及可以执行单个 Power Automate 事务的次数存在限制。

### <a name="excel"></a>Excel

Excel 网页版在通过脚本调用工作簿时具有以下限制：

- 请求和响应限制为 **5MB**。
- 范围限制为 **500 万个单元格**。

如果在处理大型数据集时遇到错误，请尝试使用多个较小的区域，而不是更大的范围。 有关示例，请参阅 [编写大型数据集](../resources/samples/write-large-dataset.md) 示例。 还可以使用 [Range.getSpecialCells](/javascript/api/office-scripts/excelscript/excelscript.range#excelscript-excelscript-range-getspecialcells-member(1)) 等 API 来面向特定单元格，而不是大范围。

Excel 规范和限制一文中可以找到不特定于 Office 脚本的 [Excel 限制](https://support.microsoft.com/office/excel-specifications-and-limits-1672b34d-7043-467e-8e27-269d656771c3)。

### <a name="power-automate"></a>Power Automate

将 Office 脚本与 Power Automate 配合使用时，每个用户 **每天只能调用 1，600 次运行脚本操作**。 此限制在 UTC 凌晨 12：00 重置。

Power Automate 平台也有使用限制，可在以下文章中找到。

- [Power Automate 中的限制和配置](/power-automate/limits-and-config)
- [Excel Online (Business) 连接器的已知问题和限制](/connectors/excelonlinebusiness/#known-issues-and-limitations)

> [!NOTE]
> 如果有长时间运行的脚本，请注意 [同步 Power Automate 操作的 120 秒超时](/power-automate/limits-and-config#timeout)。 需要 [优化脚本](../develop/web-client-performance.md) 或将 Excel 自动化拆分为多个脚本。

## <a name="see-also"></a>另请参阅

- [Excel 规范和限制](https://support.microsoft.com/office/excel-specifications-and-limits-1672b34d-7043-467e-8e27-269d656771c3)
- [Office 脚本疑难解答](troubleshooting.md)
- [消除 Office 脚本的影响](undo.md)
- [提高 Office 脚本的性能](../develop/web-client-performance.md)
