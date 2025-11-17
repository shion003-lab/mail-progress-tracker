Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
      console.log("Office.js ready");
      document.getElementById("saveButton").onclick = saveProgress;
  }
});

const siteUrl = "https://openloopcojp.sharepoint.com/sites/msteams_ed64e5";
const listName = "MailProgressTracker";

async function saveProgress() {
  try {
    const item = Office.context.mailbox.item;

    // メール情報の取得（Graph API不要）
    const messageId = item.internetMessageId;
    const subject = item.subject;
    const from = item.from ? item.from.displayName : "";
    const receivedTime = item.dateTimeCreated;

    const progress = document.getElementById("progress").value;
    const comment = document.getElementById("comment").value;
    const updatedBy = Office.context.mailbox.userProfile.displayName;
    const updatedAt = new Date().toISOString();

    // SharePoint Digest を取得
    const digest = await getRequestDigest();

    const itemData = {
      __metadata: { type: "SP.Data.MailProgressTrackerListItem" },
      MessageID: messageId,
      Subject: subject,
      From: from,
      ReceivedTime: receivedTime,
      Progress: progress,
      Comment: comment,
      UpdatedBy: updatedBy,
      UpdatedAt: updatedAt
    };

    // 既存アイテムを検索
    const existing = await fetch(
      `${siteUrl}/_api/web/lists/getbytitle('${listName}')/items?$filter=MessageID eq '${messageId}'`,
      {
        headers: { Accept: "application/json;odata=verbose" }
      }
    );

    const result = await existing.json();
    let response;

    if (result.d.results.length > 0) {
      // 既存 → MERGE（更新）
      const id = result.d.results[0].Id;
      response = await fetch(
        `${siteUrl}/_api/web/lists/getbytitle('${listName}')/items(${id})`,
        {
          method: "MERGE",
          headers: {
            "IF-MATCH": "*",
            "X-RequestDigest": digest,
            "Accept": "application/json;odata=verbose",
            "Content-Type": "application/json;odata=verbose"
          },
          body: JSON.stringify(itemData)
        }
      );
    } else {
      // 新規 → POST（追加）
      response = await fetch(
        `${siteUrl}/_api/web/lists/getbytitle('${listName}')/items`,
        {
          method: "POST",
          headers: {
            "X-RequestDigest": digest,
            "Accept": "application/json;odata=verbose",
            "Content-Type": "application/json;odata=verbose"
          },
          body: JSON.stringify(itemData)
        }
      );
    }

    if (response.ok) {
      Office.context.mailbox.item.notificationMessages.replaceAsync(
        "progressSaved",
        {
          type: "informationalMessage",
          message: "進捗情報を SharePoint に保存しました",
          icon: "icon16",
          persistent: false
        }
      );
    } else {
      console.error(await response.text());
      alert("❌ SharePoint 保存に失敗しました。コンソールをご確認ください。");
    }

  } catch (e) {
    console.error(e);
    alert("❌ 保存中にエラーが発生しました");
  }
}

async function getRequestDigest() {
  const res = await fetch(`${siteUrl}/_api/contextinfo`, {
    method: "POST",
    headers: { Accept: "application/json;odata=verbose" }
  });
  const data = await res.json();
  return data.d.GetContextWebInformation.FormDigestValue;
}
