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
# <a name="add-comments-in-excel"></a><span data-ttu-id="1c2c0-103">在内容中添加Excel</span><span class="sxs-lookup"><span data-stu-id="1c2c0-103">Add comments in Excel</span></span>

<span data-ttu-id="1c2c0-104">本示例演示如何向单元格添加注释，包括 [@mentioning添加注释](https://support.microsoft.com/office/90701709-5dc1-41c7-aa48-b01d4a46e8c7) 。</span><span class="sxs-lookup"><span data-stu-id="1c2c0-104">This sample shows how to add comments to a cell including [@mentioning](https://support.microsoft.com/office/90701709-5dc1-41c7-aa48-b01d4a46e8c7) a colleague.</span></span>

## <a name="example-scenario"></a><span data-ttu-id="1c2c0-105">示例应用场景</span><span class="sxs-lookup"><span data-stu-id="1c2c0-105">Example scenario</span></span>

* <span data-ttu-id="1c2c0-106">团队主管负责维护轮班计划。</span><span class="sxs-lookup"><span data-stu-id="1c2c0-106">The team lead maintains the shift schedule.</span></span> <span data-ttu-id="1c2c0-107">团队主管向轮班记录分配员工 ID。</span><span class="sxs-lookup"><span data-stu-id="1c2c0-107">The team lead assigns an employee ID to the shift record.</span></span>
* <span data-ttu-id="1c2c0-108">团队主管希望通知员工。</span><span class="sxs-lookup"><span data-stu-id="1c2c0-108">The team lead wishes to notify the employee.</span></span> <span data-ttu-id="1c2c0-109">通过添加一个@mentions注释，员工将通过电子邮件从工作表收到自定义邮件。</span><span class="sxs-lookup"><span data-stu-id="1c2c0-109">By adding a comment that @mentions the employee, the employee is emailed with a custom message from the worksheet.</span></span>
* <span data-ttu-id="1c2c0-110">随后，员工可以在方便时查看工作簿并回复注释。</span><span class="sxs-lookup"><span data-stu-id="1c2c0-110">Subsequently, the employee can view the workbook and respond to the comment at their convenience.</span></span>

## <a name="solution"></a><span data-ttu-id="1c2c0-111">解决方案</span><span class="sxs-lookup"><span data-stu-id="1c2c0-111">Solution</span></span>

1. <span data-ttu-id="1c2c0-112">该脚本从员工工作表中提取员工信息。</span><span class="sxs-lookup"><span data-stu-id="1c2c0-112">The script extracts employee information from the employee worksheet.</span></span>
1. <span data-ttu-id="1c2c0-113">然后，脚本添加注释 (包括相关的员工电子邮件) 到班次记录中的相应单元格。</span><span class="sxs-lookup"><span data-stu-id="1c2c0-113">The script then adds a comment (including the relevant employee email) to the appropriate cell in the shift record.</span></span>
1. <span data-ttu-id="1c2c0-114">添加新注释之前，将删除单元格中的现有注释。</span><span class="sxs-lookup"><span data-stu-id="1c2c0-114">Existing comments in the cell are removed before adding the new comment.</span></span>

## <a name="sample-code-add-comments"></a><span data-ttu-id="1c2c0-115">示例代码：添加注释</span><span class="sxs-lookup"><span data-stu-id="1c2c0-115">Sample code: Add comments</span></span>

<span data-ttu-id="1c2c0-116">下载此示例 <a href="excel-comments.xlsx">excel-comments.xlsx</a> 使用的文件，然后自己试用！</span><span class="sxs-lookup"><span data-stu-id="1c2c0-116">Download the file <a href="excel-comments.xlsx">excel-comments.xlsx</a> used in this sample and try it out yourself!</span></span>

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

## <a name="training-video-add-comments"></a><span data-ttu-id="1c2c0-117">培训视频：添加注释</span><span class="sxs-lookup"><span data-stu-id="1c2c0-117">Training video: Add comments</span></span>

<span data-ttu-id="1c2c0-118">[观看 Sudhi Ramamurthy 在 YouTube 上演练此示例](https://youtu.be/CpR78nkaOFw)。</span><span class="sxs-lookup"><span data-stu-id="1c2c0-118">[Watch Sudhi Ramamurthy walk through this sample on YouTube](https://youtu.be/CpR78nkaOFw).</span></span>
