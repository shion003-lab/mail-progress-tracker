Office.onReady(async (info) => {
  if (info.host === Office.HostType.Outlook) {
    document.getElementById("saveBtn").onclick = saveProgress;
  }
});

const siteUrl = "https://openloopcojp.sharepoint.com/sites/msteams_ed64e5";
const listName = "MailProgressTracker";

// 保存処理
async function saveProgress() {
  try {
    const item = Office.context.mailbox.item;
    const messageId = item.internetMessageId; // メール固有ID
    const status = document.getElementById("status").value;
    const progress = document.getElementById("progress").value;
    const comment = document.getElementById("comment").value;

    const me = Office.context.mailbox.userProfile.displayName;
    const now = new Date().toISOString();

    const digest = await getRequestDigest();

    const itemData = {
      __metadata: { type: "SP.Data.MailProgressTrackerListItem" },
      MessageID: messageId,
      Status: status,
      Progress: progress,
      Comment: comment,
      UpdatedBy: me,
      UpdatedAt: now
    };

    // 既存レコードがあるか確認
    const existing = await fetch(
      `${siteUrl}/_api/web/lists/getbytitle('${listName}')/items?$filter=MessageID eq '${messageId}'`,
      { headers: { Accept: "application/json;odata=verbose" } }
    );
    const result = await existing.json();

    let response;
    if (result.d.results.length > 0) {
      const id = result.d.results[0].Id;
      response = await fetch(
        `${siteUrl}/_api/web/lists/getbytitle('${listName}')/items(${id})`,
        {
          method: "MERGE",
          headers: {
            "X-RequestDigest": digest,
            "IF-MATCH": "*",
            "Accept": "application/json;odata=verbose",
            "Content-Type": "application/json;odata=verbose"
          },
          body: JSON.stringify(itemData)
        }
      );
    } else {
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
      Office.context.mailbox.item.notificationMessages.replaceAsync("progressSaved", {
        type: "informationalMessage",
        message: "進捗情報を保存しました。",
        icon: "icon16",
        persistent: false
      });
    } else {
      console.error(await response.text());
      alert("保存に失敗しました。");
    }
  } catch (e) {
    console.error(e);
    alert("エラーが発生しました。");
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
