---
title: Office 脚本的平台限制和要求
description: 在 web 上与 Excel 一起使用时，Office 脚本的资源限制和浏览器支持
ms.date: 10/23/2020
localization_priority: Normal
ms.openlocfilehash: 61f5c55be278ae056014d3b01e4176354d913f87
ms.sourcegitcommit: d3e7681e262bdccc281fcb7b3c719494202e846b
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 11/06/2020
ms.locfileid: "48930076"
---
# <a name="platform-limits-and-requirements-with-office-scripts"></a>Office 脚本的平台限制和要求

开发 Office 脚本时，应注意一些平台限制。 本文详细介绍了 web 上的适用于 Excel 的 Office 脚本的浏览器支持和数据限制。

## <a name="browser-support"></a>浏览器支持

Office 脚本在任何 [支持 Web Office 的](https://support.microsoft.com/office/ad1303e0-a318-47aa-b409-d3a5eb44e452)浏览器中工作。 但是，Internet Explorer 11 (IE 11) 中不支持某些 JavaScript 功能。 [ES6 或更高版本](https://www.w3schools.com/Js/js_es6.asp)中引入的任何功能将不适用于 IE 11。 如果组织中的人员仍在使用该浏览器，请务必在共享这些脚本时在该环境中对其进行测试。

[!INCLUDE [Teams support note](../includes/teams-support-note.md)]

### <a name="third-party-cookies"></a>第三方 cookie

你的浏览器需要启用了第三方 cookie，才能在 Excel 网页上显示 " **自动** " 选项卡。 如果不显示该选项卡，请检查您的浏览器设置。 如果使用的是专用浏览器会话，则每次可能需要重新启用此设置。

> [!NOTE]
> 某些浏览器将此设置称为 "所有 cookie"，而不是 "第三方 cookie"。

#### <a name="instructions-for-adjusting-cookie-settings-in-popular-browsers"></a>在常见浏览器中调整 cookie 设置的说明

- [Chrome](https://support.google.com/chrome/answer/95647)
- [Microsoft Edge](https://support.microsoft.com/microsoft-edge/temporarily-allow-cookies-and-site-data-in-microsoft-edge-597f04f2-c0ce-f08c-7c2b-541086362bd2)
- [Firefox](https://support.mozilla.org/kb/disable-third-party-cookies)
- [Safari](https://support.apple.com/guide/safari/manage-cookies-and-website-data-sfri11471/mac)

## <a name="data-limits"></a>数据限制

对可以一次传输多少个 Excel 数据以及可以执行多少个单独的电源自动化事务的操作有限制。

### <a name="excel"></a>Excel

在通过脚本调用工作簿时，网站的 Excel 具有以下限制：

- 请求和响应限制为 **5mb** 。
- 范围限制为5000000个 **单元格** 。

如果在处理大型数据集时遇到错误，请尝试使用多个较小的范围，而不是更大的范围。 您还可以将 [getSpecialCells](/javascript/api/office-scripts/excelscript/excelscript.range#getspecialcells-celltype--cellvaluetype-) 作为目标单元格（而不是大型区域）的 api。

### <a name="power-automate"></a>Power Automate

在使用带电自动化的 Office 脚本时， **每日限制为200个呼叫** 。 此限制在 UTC 时间重置为 12:00 AM。

电源自动化平台还有使用限制，可在 [电源自动化的文章限制和配置](/power-automate/limits-and-config)中找到。

## <a name="see-also"></a>另请参阅

- [Office 脚本疑难解答](troubleshooting.md)
- [消除 Office 脚本的影响](undo.md)
- [提高 Office 脚本的性能](../develop/web-client-performance.md)
- [Web 上的 Excel 中 Office 脚本的脚本基础](../develop/scripting-fundamentals.md)
