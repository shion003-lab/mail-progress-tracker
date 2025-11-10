// メイン処理：進捗反映ボタン押下時に呼ばれる
function updateStatus() {
  const user = document.getElementById("userName").value || "未入力";
  const status = document.getElementById("status").value || "未選択";
  const updated = new Date().toLocaleString("ja-JP");

  updateProgressBlock(user, status, updated);
}

// 本文の MailPM 区切り領域を探して置換または追加
function updateProgressBlock(user, status, updated) {
  const blockStart = "────────────────────────";
  const blockEnd = "────────────────────────";

  const newBlock = `
${blockStart}
【進捗状況】
担当者：${user}
状態：${status}
更新日時：${updated}
${blockEnd}
`;

  Office.context.mailbox.item.body.getAsync("text", function (res) {
    if (res.status !== Office.AsyncResultStatus.Succeeded) {
      alert("本文取得に失敗しました。");
      return;
    }

    let body = res.value;

    // 既存の MailPM ブロックがあるか？
    const regex = new RegExp(`${blockStart}[\\s\\S]*?${blockEnd}`, "g");

    if (regex.test(body)) {
      // 既存ブロックを置換
      body = body.replace(regex, newBlock);
    } else {
      // なければ末尾に追加
      body = body + "\n" + newBlock;
    }

    // 本文を更新
    Office.context.mailbox.item.body.setAsync(
      body,
      { coercionType: Office.CoercionType.Text },
      function (setRes) {
        if (setRes.status === Office.AsyncResultStatus.Succeeded) {
          alert("進捗情報を本文に反映しました！");
        } else {
          alert("本文の更新に失敗しました。");
          console.error(setRes.error);
        }
      }
    );
  });
}
