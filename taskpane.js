Office.onReady(() => {
  document.getElementById("updateBtn").onclick = saveProgress;
  loadProgress();
});

function saveProgress() {
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
          document.getElementById("result").innerText = "進捗情報を保存しました。";
        } else {
          document.getElementById("result").innerText = "保存に失敗しました。";
        }
      });
    } else {
      document.getElementById("result").innerText = "プロパティの読み込みに失敗しました。";
    }
  });
}

function loadProgress() {
  Office.context.mailbox.item.loadCustomPropertiesAsync((result) => {
    if (result.status === Office.AsyncResultStatus.Succeeded) {
      const props = result.value;
      const user = props.get("user") || "";
      const status = props.get("status") || "";
      const updated = props.get("updated") || "";

      if (user || status || updated) {
        document.getElementById("userName").value = user;
        document.getElementById("status").value = status || "未対応";
        document.getElementById("result").innerText = `最終更新：${updated}`;
      } else {
        document.getElementById("result").innerText = "このメールには進捗情報がありません。";
      }
    }
  });
}
