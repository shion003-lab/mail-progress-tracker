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
          showPreview(user, status, now);
        } else {
          document.getElementById("result").innerText = "保存に失敗しました。";
        }
      });
    }
  });
}

function loadProgress() {
  Office.context.mailbox.item.loadCustomPropertiesAsync((result) => {
    if (result.status === Office.AsyncResultStatus.Succeeded) {
      const props = result.value;
      const user = props.get("user") || "";
      const status = props.get("status") || "未対応";
      const updated = props.get("updated") || "";

      document.getElementById("userName").value = user;
      document.getElementById("status").value = status;
      showPreview(user, status, updated);
    }
  });
}

function showPreview(user, status, updated) {
  let existing = document.getElementById("progressPreview");
  if (!existing) {
    existing = document.createElement("div");
    existing.id = "progressPreview";
    existing.style.marginTop = "15px";
    existing.style.padding = "10px";
    existing.style.borderTop = "1px solid #999";
    existing.style.fontSize = "0.9em";
    document.body.appendChild(existing);
  }

  existing.innerHTML = `
  ────────────────────────────<br>
  【進捗状況】<br>
  担当者：${user || "未入力"}<br>
  状態：${status || "未対応"}<br>
  更新日時：${updated || "未設定"}<br>
  ────────────────────────────
  `;
}
