---
title: 在内容中添加Excel
description: 了解如何使用 Office 脚本在工作表中添加注释。
ms.date: 04/28/2021
localization_priority: Normal
ms.openlocfilehash: d592b37c3af8e475c81e8650dda44921fee7aeaf
ms.sourcegitcommit: f7a7aebfb687f2a35dbed07ed62ff352a114525a
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 05/06/2021
ms.locfileid: "52232506"
---
# <a name="add-comments-in-excel"></a>在内容中添加Excel

本示例演示如何向单元格添加注释，包括 [@mentioning添加注释](https://support.microsoft.com/office/90701709-5dc1-41c7-aa48-b01d4a46e8c7) 。

## <a name="example-scenario"></a>示例应用场景

* 团队主管负责维护轮班计划。 团队主管向轮班记录分配员工 ID。
* 团队主管希望通知员工。 通过添加一个@mentions注释，员工将通过电子邮件从工作表收到自定义邮件。
* 随后，员工可以在方便时查看工作簿并回复注释。

## <a name="solution"></a>解决方案

1. 该脚本从员工工作表中提取员工信息。
1. 然后，脚本添加注释 (包括相关的员工电子邮件) 到班次记录中的相应单元格。
1. 添加新注释之前，将删除单元格中的现有注释。

## <a name="sample-code-add-comments"></a>示例代码：添加注释

下载此示例 <a href="excel-comments.xlsx">excel-comments.xlsx</a> 使用的文件，然后自己试用！

```TypeScript
function main(workbook: ExcelScript.Workbook) {
    const employees = workbook.getWorksheet('Employees').getUsedRange().getTexts();
    console.log(employees); 

    const scheduleSheet = workbook.getWorksheet('Schedule');
    const table = scheduleSheet.getTables()[0];
    const range = table.getRangeBetweenHeaderAndTotal();
    const scheduleData = range.getTexts();

    for (let i=0; i < scheduleData.length; i++) {
      let eId = scheduleData[i][3];

      let employeeInfo = employees.find(e => e[0] === eId);
      if (employeeInfo) {
        console.log("Found a match " + employeeInfo);
        let adminNotes = scheduleData[i][4];
        try { 
          let comment = workbook.getCommentByCell(range.getCell(i, 5));
          comment.delete();
        } catch {
            console.log("Ignore if there is no existing comment in the cell");
        }
        workbook.addComment(range.getCell(i,5), {
          mentions: [{
            email: employeeInfo[1],
            id: 0,
            name: employeeInfo[2]
          }],
          richContent: `<at id=\"0\">${employeeInfo[2]}</at> ${adminNotes}`
        }, ExcelScript.ContentType.mention);        
        
      } else {
        console.log("No match for: " + eId);
      }
    }
    return;
}
```

## <a name="training-video-add-comments"></a>培训视频：添加注释

[观看 Sudhi Ramamurthy 在 YouTube 上演练此示例](https://youtu.be/CpR78nkaOFw)。
