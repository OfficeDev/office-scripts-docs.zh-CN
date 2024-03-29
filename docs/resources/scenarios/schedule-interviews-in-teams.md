---
title: 在 Teams 中安排面试
description: 了解如何使用 Office 脚本从 Excel 数据发送 Teams 会议。
ms.date: 06/29/2021
ms.localizationpriority: medium
ms.openlocfilehash: 8e8c4af40398842e219dc3e2a80c6d2ee72d6b83
ms.sourcegitcommit: a6504f8b0d6b717457c6e0b5306c35ad3900914e
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/02/2022
ms.locfileid: "67572575"
---
# <a name="office-scripts-sample-scenario-schedule-interviews-in-teams"></a>Office 脚本示例方案：在 Teams 中安排面试

在此方案中，你是一名人力资源招聘人员，在 Teams 中安排与候选人的面试会议。 你管理 Excel 文件中应聘者的采访计划。 你需要向候选人和面试官发送 Teams 会议邀请。 然后，需要更新 Excel 文件，确认已发送 Teams 会议。

该解决方案有三个在单个 Power Automate 流中组合的步骤。

1. 脚本从表中提取数据，并将对象数组作为 [JSON](https://www.w3schools.com/whatis/whatis_json.asp) 数据返回。
1. 然后，数据将发送到 Teams **创建 Teams 会议** 操作以发送邀请。
1. 相同的 JSON 数据将发送到另一个脚本，以更新邀请的状态。

有关使用 JSON 的详细信息，请阅读 [“使用 JSON 向 Office 脚本传递数据](../../develop/use-json.md)”。

## <a name="scripting-skills-covered"></a>所涵盖的脚本技能

* Power Automate 流
* Teams 集成
* 表分析

## <a name="sample-excel-file"></a>示例 Excel 文件

下载此解决方案中使用的文件 [hr-schedule.xlsx](hr-schedule.xlsx) 并自行试用！ 请务必更改至少一个电子邮件地址，以便收到邀请。

## <a name="sample-code-extract-table-data-to-schedule-invites"></a>示例代码：提取表数据以计划邀请

将此脚本添加到脚本集合。 将其命名为 **计划流的访谈** 。

```TypeScript
function main(workbook: ExcelScript.Workbook): InterviewInvite[] {
  const MEETING_DURATION = workbook.getWorksheet("Constants").getRange("B1").getValue() as number;
  const MESSAGE_TEMPLATE = workbook.getWorksheet("Constants").getRange("B2").getValue() as string;

  // Get the interview candidate information.
  const sheet = workbook.getWorksheet("Interviews");
  const table = sheet.getTables()[0];
  const dataRows = table.getRangeBetweenHeaderAndTotal().getValues();

  // Convert the table rows into InterviewInvite objects for the flow.
  let invites: InterviewInvite[] = [];
  dataRows.forEach((row) => {
    const inviteSent = row[1] as boolean;
    if (!inviteSent) {
      const startTime = new Date(Math.round(((row[6] as number) - 25569) * 86400 * 1000));
      const finishTime = new Date(startTime.getTime() + MEETING_DURATION * 60 * 1000);
      const candidateName = row[2] as string;
      const interviewerName = row[4] as string;

      invites.push({
        ID: row[0] as string,
        Candidate: candidateName,
        CandidateEmail: row[3] as string,
        Interviewer: row[4] as string,
        InterviewerEmail: row[5] as string,
        StartTime: startTime.toISOString(),
        FinishTime: finishTime.toISOString(),
        Message: generateInviteMessage(MESSAGE_TEMPLATE, candidateName, interviewerName)
      });
    }    
  });

  console.log(JSON.stringify(invites));
  return invites;
}

function generateInviteMessage(
  messageTemplate: string,
   candidate: string,
   interviewer: string) : string {
  return messageTemplate.replace("_Candidate_", candidate).replace("_Interviewer_", interviewer);
}

// The interview invite information.
interface InterviewInvite {
  ID: string
  Candidate: string
  CandidateEmail: string
  Interviewer: string
  InterviewerEmail: string
  StartTime: string
  FinishTime: string
  Message: string
}
```

## <a name="sample-code-mark-rows-as-invited"></a>示例代码：将行标记为受邀

将此脚本添加到脚本集合。 将其命名为 **记录流的发送邀请** 。

```TypeScript
function main(workbook: ExcelScript.Workbook, invites: InterviewInvite[]) {
  const table = workbook.getWorksheet("Interviews").getTables()[0];

  // Get the ID and Invite Sent columns from the table.
  const idColumn = table.getColumnByName("ID");
  const idRange = idColumn.getRangeBetweenHeaderAndTotal().getValues();
  const inviteSentColumn = table.getColumnByName("Invite Sent?");

  const dataRowCount = idRange.length;

  // Find matching IDs to mark the correct row.
  for (let row = 0; row < dataRowCount; row++){
    let inviteSent = invites.find((invite) => {
      return invite.ID == idRange[row][0] as string;
    });

    if (inviteSent) {
      inviteSentColumn.getRangeBetweenHeaderAndTotal().getCell(row, 0).setValue(true);
      console.log(`Invite for ${inviteSent.Candidate} has been sent.`);
    }
  } 
}

// The interview invite information.
interface InterviewInvite {
  ID: string
  Candidate: string
  CandidateEmail: string
  Interviewer: string
  InterviewerEmail: string
  StartTime: string
  FinishTime: string
  Message: string
}
```

## <a name="sample-flow-run-the-interview-scheduling-scripts-and-send-the-teams-meetings"></a>示例流：运行面试日程安排脚本并发送 Teams 会议

1. 创建新的 **即时云流**。
1. 选择 **“手动触发流** ”，然后选择 **“创建**”。
1. 添加使用 **Excel Online (Business)** 连接器和 **运行脚本** 操作 **的新步骤**。 使用以下值完成连接器。
    1. **位置**：OneDrive for Business
    1. **文档库**：OneDrive
    1. **文件**：通过 *文件浏览器) 选择hr-interviews.xlsx (*
    1. **脚本**：已 :::image type="content" source="../../images/schedule-interviews-1.png" alt-text="完成的 Excel Online (Business) 连接器的“计划访谈”屏幕截图，用于从 Power Automate 中的工作簿获取面试数据。":::
1. 添加使用 **“创建 Teams 会议**”操作 **的新步骤**。 从 Excel 连接器中选择动态内容时，将为流生成 **对每个** 块的应用。 使用以下值完成连接器。
    1. **日历 ID**：日历
    1. **主题**：Contoso 访谈
    1. **消息**： **消息** (Excel 值) 
    1. **时区**：太平洋标准时间
    1. **开始时间**： **StartTime** (Excel 值) 
    1. **结束时间**： **FinishTime** (Excel 值) 
    1. **必需与会者**： **CandidateEmail** ; **InterviewerEmail** (已 :::image type="content" source="../../images/schedule-interviews-2.png" alt-text="完成的 Teams 连接器的 Excel 值) 屏幕截图，用于在 Power Automate 中安排会议。":::
1. 在同一个 **应用到每个** 块中，使用 **“运行脚本**”操作添加另一个 **Excel Online (Business)** 连接器。 使用以下值。
    1. **位置**：OneDrive for Business
    1. **文档库**：OneDrive
    1. **文件**：通过 *文件浏览器) 选择hr-interviews.xlsx (*
    1. **脚本**：记录发送的邀请
    1. **邀请**： **结果** (Excel 值) :::image type="content" source="../../images/schedule-interviews-3.png" alt-text="已完成的 Excel Online (Business) 连接器的屏幕截图，用于记录已在 Power Automate 中发送的邀请。":::
1. 保存流并试用。使用流编辑器页上的 **“测试** ”按钮，或通过“ **我的流** ”选项卡运行流。出现提示时，请务必允许访问。

## <a name="training-video-send-a-teams-meeting-from-excel-data"></a>培训视频：从 Excel 数据发送 Teams 会议

[观看 Sudhi Ramamurthy 在 YouTube 上浏览此示例的版本](https://youtu.be/HyBdx52NOE8)。 他的版本使用更可靠的脚本来处理更改的列和过时的会议时间。
