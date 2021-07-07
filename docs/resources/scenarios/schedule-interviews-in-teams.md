---
title: 在 Teams 中安排面试
description: 了解如何使用 Office 脚本从Teams发送Excel会议。
ms.date: 06/29/2021
localization_priority: Normal
ms.openlocfilehash: cb24da12637add805d86da4d07ce878509c6a5f6
ms.sourcegitcommit: 211c157ca746e266eeb079f5fa1925a1e35ab702
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 07/07/2021
ms.locfileid: "53313727"
---
# <a name="office-scripts-sample-scenario-schedule-interviews-in-teams"></a><span data-ttu-id="a0cb1-103">Office脚本示例方案：安排在 Teams</span><span class="sxs-lookup"><span data-stu-id="a0cb1-103">Office Scripts sample scenario: Schedule interviews in Teams</span></span>

<span data-ttu-id="a0cb1-104">在此方案中，你是一名 HR 招聘人员，负责安排与Teams。</span><span class="sxs-lookup"><span data-stu-id="a0cb1-104">In this scenario, you're an HR recruiter scheduling interview meetings with candidates in Teams.</span></span> <span data-ttu-id="a0cb1-105">在一个管理文件中管理应聘者的Excel计划。</span><span class="sxs-lookup"><span data-stu-id="a0cb1-105">You manage the interview schedule of candidates in an Excel file.</span></span> <span data-ttu-id="a0cb1-106">你需要向候选人和Teams发送会议邀请。</span><span class="sxs-lookup"><span data-stu-id="a0cb1-106">You'll need to send the Teams meeting invite to both the candidate and interviewers.</span></span> <span data-ttu-id="a0cb1-107">然后，你需要更新Excel文件，并确认Teams会议已发送。</span><span class="sxs-lookup"><span data-stu-id="a0cb1-107">You then need to update the Excel file with the confirmation that Teams meetings have been sent.</span></span>

<span data-ttu-id="a0cb1-108">解决方案有三个步骤组合在单个流Power Automate流。</span><span class="sxs-lookup"><span data-stu-id="a0cb1-108">The solution has three steps that are combined in a single Power Automate flow.</span></span>

1. <span data-ttu-id="a0cb1-109">脚本从表中提取数据，并返回对象数组作为 JSON 数据。</span><span class="sxs-lookup"><span data-stu-id="a0cb1-109">A script extracts data from a table and returns an array of objects as JSON data.</span></span>
1. <span data-ttu-id="a0cb1-110">然后，数据将发送到Teams **创建Teams会议操作** 以发送邀请。</span><span class="sxs-lookup"><span data-stu-id="a0cb1-110">The data is then sent to the Teams **Create a Teams meeting** action to send invites.</span></span>
1. <span data-ttu-id="a0cb1-111">相同的 JSON 数据将发送到另一个脚本以更新邀请的状态。</span><span class="sxs-lookup"><span data-stu-id="a0cb1-111">The same JSON data is sent to another script to update the status of the invitation.</span></span>

## <a name="scripting-skills-covered"></a><span data-ttu-id="a0cb1-112">涵盖的脚本编写技能</span><span class="sxs-lookup"><span data-stu-id="a0cb1-112">Scripting skills covered</span></span>

* <span data-ttu-id="a0cb1-113">Power Automate流</span><span class="sxs-lookup"><span data-stu-id="a0cb1-113">Power Automate flows</span></span>
* <span data-ttu-id="a0cb1-114">Teams集成</span><span class="sxs-lookup"><span data-stu-id="a0cb1-114">Teams integration</span></span>
* <span data-ttu-id="a0cb1-115">表分析</span><span class="sxs-lookup"><span data-stu-id="a0cb1-115">Table parsing</span></span>

## <a name="sample-excel-file"></a><span data-ttu-id="a0cb1-116">示例Excel文件</span><span class="sxs-lookup"><span data-stu-id="a0cb1-116">Sample Excel file</span></span>

<span data-ttu-id="a0cb1-117">下载此 <a href="hr-schedule.xlsx">hr-schedule.xlsx</a> 中使用的文件，然后自己试用！</span><span class="sxs-lookup"><span data-stu-id="a0cb1-117">Download the file <a href="hr-schedule.xlsx">hr-schedule.xlsx</a> used in this solution and try it out yourself!</span></span> <span data-ttu-id="a0cb1-118">请务必更改至少一个电子邮件地址，以便收到邀请。</span><span class="sxs-lookup"><span data-stu-id="a0cb1-118">Be sure to change at least one of the email addresses so that you receive an invite.</span></span>

## <a name="sample-code-extract-table-data-to-schedule-invites"></a><span data-ttu-id="a0cb1-119">示例代码：提取表数据以计划邀请</span><span class="sxs-lookup"><span data-stu-id="a0cb1-119">Sample code: Extract table data to schedule invites</span></span>

<span data-ttu-id="a0cb1-120">将此脚本添加到脚本集合。</span><span class="sxs-lookup"><span data-stu-id="a0cb1-120">Add this script to your script collection.</span></span> <span data-ttu-id="a0cb1-121">将流程 **命名为"安排** 面试"。</span><span class="sxs-lookup"><span data-stu-id="a0cb1-121">Name it **Schedule Interviews** for the flow.</span></span>

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

## <a name="sample-code-mark-rows-as-invited"></a><span data-ttu-id="a0cb1-122">示例代码：将行标记为受邀</span><span class="sxs-lookup"><span data-stu-id="a0cb1-122">Sample code: Mark rows as invited</span></span>

<span data-ttu-id="a0cb1-123">将此脚本添加到脚本集合。</span><span class="sxs-lookup"><span data-stu-id="a0cb1-123">Add this script to your script collection.</span></span> <span data-ttu-id="a0cb1-124">Name it **Record Sent Invites** for the flow.</span><span class="sxs-lookup"><span data-stu-id="a0cb1-124">Name it **Record Sent Invites** for the flow.</span></span>

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

## <a name="sample-flow-run-the-interview-scheduling-scripts-and-send-the-teams-meetings"></a><span data-ttu-id="a0cb1-125">示例流程：运行访谈计划脚本并发送Teams会议</span><span class="sxs-lookup"><span data-stu-id="a0cb1-125">Sample flow: Run the interview scheduling scripts and send the Teams meetings</span></span>

1. <span data-ttu-id="a0cb1-126">创建新的即时 **云流**。</span><span class="sxs-lookup"><span data-stu-id="a0cb1-126">Create a new **Instant cloud flow**.</span></span>
1. <span data-ttu-id="a0cb1-127">选择 **"手动触发流"，** 然后选择"创建 **"。**</span><span class="sxs-lookup"><span data-stu-id="a0cb1-127">Choose **Manually trigger a flow** and select **Create**.</span></span>
1. <span data-ttu-id="a0cb1-128">添加使用 **Excel** Online (**Business) 连接器** 和 **"运行** 脚本"操作的新步骤。</span><span class="sxs-lookup"><span data-stu-id="a0cb1-128">Add a **New step** that uses the **Excel Online (Business)** connector and the **Run script** action.</span></span> <span data-ttu-id="a0cb1-129">使用下列值完成连接器。</span><span class="sxs-lookup"><span data-stu-id="a0cb1-129">Complete the connector with the following values.</span></span>
    1. <span data-ttu-id="a0cb1-130">**位置**：OneDrive for Business</span><span class="sxs-lookup"><span data-stu-id="a0cb1-130">**Location**: OneDrive for Business</span></span>
    1. <span data-ttu-id="a0cb1-131">**文档库**：OneDrive</span><span class="sxs-lookup"><span data-stu-id="a0cb1-131">**Document Library**: OneDrive</span></span>
    1. <span data-ttu-id="a0cb1-132">\**文件\*\*\*：hr-interviews.xlsx (浏览器选项选择)*</span><span class="sxs-lookup"><span data-stu-id="a0cb1-132">**File**: hr-interviews.xlsx *(Chosen through the file browser)*</span></span>
    1. **脚本**：计划面试 已完成的 Excel Online (Business) 连接器的屏幕截图，用于从 Power Automate 中的工作簿 :::image type="content" source="../../images/schedule-interviews-1.png" alt-text="获取面试Power Automate。":::
1. <span data-ttu-id="a0cb1-134">添加一 **个使用**"创建会议 **Teams的新** 步骤。</span><span class="sxs-lookup"><span data-stu-id="a0cb1-134">Add a **New step** that uses the **Create a Teams meeting** action.</span></span> <span data-ttu-id="a0cb1-135">当你从连接器选择动态内容Excel，将针对你的流生成应用到每个块。</span><span class="sxs-lookup"><span data-stu-id="a0cb1-135">As you select dynamic content from the Excel connector, an **Apply to each** block will be generated for your flow.</span></span> <span data-ttu-id="a0cb1-136">使用下列值完成连接器。</span><span class="sxs-lookup"><span data-stu-id="a0cb1-136">Complete the connector with the following values.</span></span>
    1. <span data-ttu-id="a0cb1-137">**日历 ID**：日历</span><span class="sxs-lookup"><span data-stu-id="a0cb1-137">**Calendar id**: Calendar</span></span>
    1. <span data-ttu-id="a0cb1-138">**主题**：Contoso Interview</span><span class="sxs-lookup"><span data-stu-id="a0cb1-138">**Subject**: Contoso Interview</span></span>
    1. <span data-ttu-id="a0cb1-139">**邮件\*\*\*\*： (Excel** 值) </span><span class="sxs-lookup"><span data-stu-id="a0cb1-139">**Message**: **Message** (the Excel value)</span></span>
    1. <span data-ttu-id="a0cb1-140">**时区：** 太平洋标准时间</span><span class="sxs-lookup"><span data-stu-id="a0cb1-140">**Time zone**: Pacific Standard Time</span></span>
    1. <span data-ttu-id="a0cb1-141">**开始时间\*\*\*\*：StartTime** (Excel值) </span><span class="sxs-lookup"><span data-stu-id="a0cb1-141">**Start time**: **StartTime** (the Excel value)</span></span>
    1. <span data-ttu-id="a0cb1-142">**结束时间\*\*\*\*：FinishTime** (Excel值) </span><span class="sxs-lookup"><span data-stu-id="a0cb1-142">**End time**: **FinishTime** (the Excel value)</span></span>
    1. **必选与会者**： **CandidateEmail** ;**ScreenshotEmail** (Excel值) 已完成的 Teams 连接器的屏幕截图，:::image type="content" source="../../images/schedule-interviews-2.png" alt-text="以在 Power Automate 中安排会议。":::
1. <span data-ttu-id="a0cb1-144">在同一 **个"应用到每个** 块"中，使用"运行脚本Excel添加 (**Business)** 连接器。 </span><span class="sxs-lookup"><span data-stu-id="a0cb1-144">In the same **Apply to each** block, add another **Excel Online (Business)** connector with the **Run script** action.</span></span> <span data-ttu-id="a0cb1-145">使用以下值。</span><span class="sxs-lookup"><span data-stu-id="a0cb1-145">Use the following values.</span></span>
    1. <span data-ttu-id="a0cb1-146">**位置**：OneDrive for Business</span><span class="sxs-lookup"><span data-stu-id="a0cb1-146">**Location**: OneDrive for Business</span></span>
    1. <span data-ttu-id="a0cb1-147">**文档库**：OneDrive</span><span class="sxs-lookup"><span data-stu-id="a0cb1-147">**Document Library**: OneDrive</span></span>
    1. <span data-ttu-id="a0cb1-148">\**文件\*\*\*：hr-interviews.xlsx (浏览器选项选择)*</span><span class="sxs-lookup"><span data-stu-id="a0cb1-148">**File**: hr-interviews.xlsx *(Chosen through the file browser)*</span></span>
    1. <span data-ttu-id="a0cb1-149">**脚本**：记录已发送的邀请</span><span class="sxs-lookup"><span data-stu-id="a0cb1-149">**Script**: Record Sent Invites</span></span>
    1. **invites** **：result** (the Excel value) :::image type="content" source="../../images/schedule-interviews-3.png" alt-text="Screenshot of the completed Excel Online (Business) connector to record that invites have been sent in Power Automate.":::
1. <span data-ttu-id="a0cb1-151">保存流并试用。使用" **流** 编辑器"页上的"测试"按钮，或通过"我的流" **选项卡运行** 流。请务必在系统提示时允许访问。</span><span class="sxs-lookup"><span data-stu-id="a0cb1-151">Save the flow and try it out. Use the **Test** button on the flow editor page or run the flow through your **My flows** tab. Be sure to allow access when prompted.</span></span>

## <a name="training-video-send-a-teams-meeting-from-excel-data"></a><span data-ttu-id="a0cb1-152">培训视频：从Teams发送Excel会议</span><span class="sxs-lookup"><span data-stu-id="a0cb1-152">Training video: Send a Teams meeting from Excel data</span></span>

<span data-ttu-id="a0cb1-153">[观看 Sudhi Ramamurthy 在 YouTube](https://youtu.be/HyBdx52NOE8)上演练此示例的版本。</span><span class="sxs-lookup"><span data-stu-id="a0cb1-153">[Watch Sudhi Ramamurthy walk through a version of this sample on YouTube](https://youtu.be/HyBdx52NOE8).</span></span> <span data-ttu-id="a0cb1-154">他的版本使用更强大的脚本来处理更改的列和过时的会议时间。</span><span class="sxs-lookup"><span data-stu-id="a0cb1-154">His version uses a more robust script that handles changing columns and obsolete meeting times.</span></span>
