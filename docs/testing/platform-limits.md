---
title: Office 脚本的平台限制和要求
description: 与脚本一Office脚本的资源限制和浏览器Excel web 版
ms.date: 05/17/2021
ms.localizationpriority: medium
ms.openlocfilehash: 6db32cd1c72850f02ee9e2a714c1a0a3bccd727d
ms.sourcegitcommit: ea9c3685a7407a07acf802705f4f44fb007c7825
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 12/08/2021
ms.locfileid: "61333489"
---
# <a name="platform-limits-and-requirements-with-office-scripts"></a>Office 脚本的平台限制和要求

开发脚本时应注意一些平台Office限制。 本文详细介绍了 Office Scripts for Excel web 版 的浏览器支持和Excel web 版。

## <a name="browser-support"></a>浏览器支持

Office脚本在任何支持自定义脚本的[浏览器中Office 网页版。](https://support.microsoft.com/office/ad1303e0-a318-47aa-b409-d3a5eb44e452) 但是，IE 11 版本 11 Internet Explorer不支持 (JavaScript) 。 [ES6 或更高版本中引入](https://www.w3schools.com/Js/js_es6.asp)的任何功能将不能与 IE 11 一起使用。 如果组织成员仍使用该浏览器，请务必在共享脚本时测试该环境中脚本。

[!INCLUDE [Teams support note](../includes/teams-support-note.md)]

### <a name="third-party-cookies"></a>第三方 Cookie

浏览器需要启用第三方 Cookie，以在浏览器中显示"**自动Excel web 版。** 如果未显示选项卡，请检查浏览器设置。 如果使用的是专用浏览器会话，可能需要每次重新启用此设置。

> [!NOTE]
> 某些浏览器将此设置视为"所有 Cookie"，而不是"第三方 Cookie"。

#### <a name="instructions-for-adjusting-cookie-settings-in-popular-browsers"></a>在热门浏览器中调整 Cookie 设置的说明

- [Chrome](https://support.google.com/chrome/answer/95647)
- [Microsoft Edge](https://support.microsoft.com/microsoft-edge/597f04f2-c0ce-f08c-7c2b-541086362bd2)
- [Firefox](https://support.mozilla.org/kb/disable-third-party-cookies)
- [Safari](https://support.apple.com/guide/safari/manage-cookies-and-website-data-sfri11471/mac)

## <a name="data-limits"></a>数据限制

对于一次可Excel的数据量以及可以执行单个数据传输Power Automate存在限制。

### <a name="excel"></a>Excel

Excel 网页版脚本调用工作簿时，该限制具有以下限制：

- 请求和响应限制为 **5MB。**
- 范围限制为五百 **万个单元格**。

如果在处理大型数据集时遇到错误，请尝试使用多个较小的范围，而不是较大的区域。 有关示例，请参阅编写 [大型数据集](../resources/samples/write-large-dataset.md) 示例。 您还可以使用 [Range.getSpecialCells](/javascript/api/office-scripts/excelscript/excelscript.range#getSpecialCells_cellType__cellValueType_) 等 API 来定位特定单元格，而不是大型区域。

### <a name="power-automate"></a>Power Automate

将Office脚本Power Automate时，每个用户每天只能调用 **800** 次 Run Script 操作。 此限制在 UTC 时间上午 12：00 重置。

Power Automate平台还具有使用限制，可在以下文章中找到这些限制：

- [中限制和Power Automate](/power-automate/limits-and-config)
- [Excel Online (Business) 连接器的已知问题和限制](/connectors/excelonlinebusiness/#known-issues-and-limitations)

## <a name="see-also"></a>另请参阅

- [脚本Office疑难解答](troubleshooting.md)
- [消除 Office 脚本的影响](undo.md)
- [提高脚本Office性能](../develop/web-client-performance.md)
- [脚本编写基础Office中的脚本Excel web 版](../develop/scripting-fundamentals.md)
