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

    // メールID（検索キー）
    const mailId = item.internetMessageId;

    // UI入力
    const status = document.getElementById("status").value;
    const progressValue = parseInt(document.getElementById("progress").value, 10);
    const comment = document.getElementById("comment").value;

    const updatedBy = Office.context.mailbox.userProfile.displayName;
    const updatedAt = new Date().toISOString();

    // SP Digest
    const digest = await getRequestDigest();

    // データ本体
    const itemData = {
      __metadata: { type: "SP.Data.MailProgressTrackerListItem" },
      MailID: mailId,
      Status: status,
      Progress: progressValue,
      Comment: comment,
      UpdatedBy: updatedBy,
      UpdatedAt: updatedAt
    };

    // 同一MailIDで既存件数検索
    const check = await fetch(
      `${siteUrl}/_api/web/lists/getbytitle('${listName}')/items?$filter=MailID eq '${mailId}'`,
      { headers: { Accept: "application/json;odata=verbose" } }
    );
    const checkJson = await check.json();

    let response;
    if (checkJson.d.results.length > 0) {
      // 更新(MERGE)
      const id = checkJson.d.results[0].Id;

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
      // 新規追加(POST)
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
          message: "SharePoint に保存しました",
          icon: "icon16",
          persistent: false
        }
      );
    } else {
      console.error(await response.text());
      alert("SharePoint 保存に失敗しました。コンソールを確認してください。");
    }

  } catch (err) {
    console.error(err);
    alert("保存処理中にエラーが発生しました。");
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
