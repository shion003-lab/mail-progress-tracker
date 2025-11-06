Office.onReady(() => {
  checkRecipientCondition();
  document.getElementById("updateBtn").onclick = updateStatus;
});

function checkRecipientCondition() {
  const item = Office.context.mailbox.item;
  const recipients = item.to || [];

  const target = "aomori-olp@openloop.co.jp";
  const matched = recipients.some(r =>
    (r.emailAddress || "").toLowerCase() === target.toLowerCase()
  );

  if (!matched) {
    document.getElementById("updateBtn").disabled = true;
    document.getElementById("result").innerText =
      "このメールは対象外です（aomori-olp@openloop.co.jp宛でのみ有効）";
  }
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

