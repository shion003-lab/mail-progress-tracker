Office.onReady(() => {
  // 初期表示時に既存ステータス読み込み
  loadStatus();
  document.getElementById("updateBtn").onclick = updateStatus;
});

function loadStatus() {
  Office.context.mailbox.item.loadCustomPropertiesAsync((result) => {
    if (result.status === Office.AsyncResultStatus.Succeeded) {
      const props = result.value;
      const user = props.get("user") || "";
      const status = props.get("status") || "未対応";
      const updated = props.get("updated") || "";

      document.getElementById("userName").value = user;
      document.getElementById("status").value = status;
      document.getElementById("result").innerText = updated
        ? `最終更新：${updated}（${user}）`
        : "進捗未登録";
    } else {
      document.getElementById("result").innerText = "進捗情報を読み込めませんでした。";
    }
  });
}

function updateStatus() {
  const user = document.getElementById("userName").value || "未入力";
  const status = document.getElementById("status").value;
  const now = new Date().toLocaleString("ja-JP");

  Office.context.mailbox.item.loadCustomPropertiesAsync((result) => {
    if (result.status === Office.AsyncResultStatus.Succeeded) {
      const props = result.value;
      props.set("user", user);
      props.set("status", status);
      props.set("updated", now);
      props.saveAsync((res) => {
        if (res.status === Office.AsyncResultStatus.Succeeded) {
          document.getElementById("result").innerText = `保存しました：${now}`;
        } else {
          document.getElementById("result").innerText = "保存に失敗しました。";
        }
      });
    } else {
      document.getElementById("result").innerText = "プロパティの読み込みに失敗しました。";
    }
  });
}
