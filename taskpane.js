Office.onReady(() => {
  document.getElementById("updateBtn").onclick = updateStatus;
});

function updateStatus() {
  const user = document.getElementById("userName").value || "未入力";
  const status = document.getElementById("status").value;
  const now = new Date().toLocaleString("ja-JP");

  const footer = `
---
【進捗状況】
担当者：${user}
状態：${status}
更新日時：${now}
---`;

  Office.context.mailbox.item.body.getAsync("text", (result) => {
    if (result.status === Office.AsyncResultStatus.Succeeded) {
      let body = result.value;
      // 既存の進捗情報があれば置き換え
      const regex = /---[\s\S]*【進捗状況】[\s\S]*---/g;
      body = body.replace(regex, "");
      body += "\n" + footer;

      Office.context.mailbox.item.body.setAsync(
        body,
        { coercionType: Office.CoercionType.Text },
        (res) => {
          if (res.status === Office.AsyncResultStatus.Succeeded) {
            document.getElementById("result").innerText = "進捗情報を更新しました。";
          } else {
            document.getElementById("result").innerText = "更新に失敗しました。";
          }
        }
      );
    }
  });
}
