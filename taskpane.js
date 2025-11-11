Office.onReady(() => {
  document.getElementById("updateBtn").onclick = updateStatus;
});

// アクセストークンを localStorage から取得
function getAccessToken() {
  const token = localStorage.getItem("graph_access_token");
  if (!token) {
    alert("まだMicrosoftにサインインしていません。auth.html でサインインしてください。");
  }
  return token;
}

// Graph API でメール本文を更新
async function updateMailBody(progressText, accessToken) {
  const itemId = Office.context.mailbox.item.itemId;
  const endpoint = `https://graph.microsoft.com/v1.0/me/messages/${itemId}`;

  try {
    const response = await fetch(endpoint, {
      method: "PATCH",
      headers: {
        "Authorization": `Bearer ${accessToken}`,
        "Content-Type": "application/json"
      },
      body: JSON.stringify({
        body: {
          contentType: "HTML",
          content: `<p><strong>進捗状況:</strong> ${progressText}</p>`
        }
      })
    });

    if (response.ok) {
      document.getElementById("result").innerText = "メール本文に進捗情報を反映しました！";
    } else {
      const errText = await response.text();
      console.error("Graph API Error:", errText);
      document.getElementById("result").innerText = "Graph API呼び出しに失敗しました。コンソールを確認してください。";
    }
  } catch (err) {
    console.error(err);
    document.getElementById("result").innerText = "通信エラーです。ネットワークを確認してください。";
  }
}

// ボタン押下で進捗情報を取得して反映
function updateStatus() {
  const user = document.getElementById("userName").value || "未入力";
  const status = document.getElementById("status").value;
  const now = new Date().toLocaleString("ja-JP");
  const progressText = `担当者: ${user}, 状態: ${status}, 更新日時: ${now}`;

  const accessToken = getAccessToken();
  if (!accessToken) return;

  updateMailBody(progressText, accessToken);
}
