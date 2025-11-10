Office.onReady(() => {
  document.getElementById("updateBtn").onclick = updateProgress;
  loadProgress();
});

const siteUrl = "https://openloopcojp.sharepoint.com/sites/msteams_ed64e5";
const listName = "MailProgress";

async function getAccessToken() {
  return new Promise((resolve, reject) => {
    Office.auth.getAccessTokenAsync({ forceConsent: false }, (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        resolve(result.value);
      } else {
        reject(result.error);
      }
    });
  });
}

async function updateProgress() {
  const user = document.getElementById("userName").value || "未入力";
  const status = document.getElementById("status").value;
  const now = new Date().toLocaleString("ja-JP");
  const mail = Office.context.mailbox.item;
  const mailId = mail.itemId;

  try {
    const token = await getAccessToken();
    const webUrl = `${siteUrl}/_api/web/lists/GetByTitle('${listName}')/items`;

    // 既存レコード確認
    const checkResp = await fetch(
      `${webUrl}?$filter=MailID eq '${mailId}'`,
      { headers: { Authorization: `Bearer ${token}` } }
    );
    const data = await checkResp.json();

    let method = "POST";
    let url = webUrl;
    let body = {
      MailID: mailId,
      担当者: user,
      状態: status,
      更新日時: now
    };

    if (data.value.length > 0) {
      // 更新
      method = "PATCH";
      const itemId = data.value[0].Id;
      url = `${webUrl}(${itemId})`;
    }

    const response = await fetch(url, {
      method,
      headers: {
        "Authorization": `Bearer ${token}`,
        "Content-Type": "application/json;odata=verbose",
        "Accept": "application/json;odata=verbose",
        "IF-MATCH": "*"
      },
      body: JSON.stringify(body)
    });

    if (response.ok) {
      document.getElementById("result").innerText = "進捗をSharePointに保存しました。";
    } else {
      const errText = await response.text();
      document.getElementById("result").innerText = "保存失敗：" + errText;
    }
  } catch (err) {
    console.error(err);
    document.getElementById("result").innerText = "トークン取得または通信に失敗しました。";
  }
}

async function loadProgress() {
  const mail = Office.context.mailbox.item;
  const mailId = mail.itemId;

  try {
    const token = await getAccessToken();
    const webUrl = `${siteUrl}/_api/web/lists/GetByTitle('${listName}')/items?$filter=MailID eq '${mailId}'`;
    const response = await fetch(webUrl, {
      headers: { Authorization: `Bearer ${token}` }
    });
    const data = await response.json();

    if (data.value.length > 0) {
      const item = data.value[0];
      document.getElementById("userName").value = item.担当者 || "";
      document.getElementById("status").value = item.状態 || "未対応";
      document.getElementById("result").innerText =
        `最終更新: ${item.更新日時}`;
    } else {
      document.getElementById("result").innerText = "進捗データはまだありません。";
    }
  } catch (err) {
    console.error(err);
    document.getElementById("result").innerText = "読み込みエラー。";
  }
}
