Office.onReady(async () => {
  if (Office.context.mailbox) {
    loadSavedData();
    document.getElementById("saveButton").onclick = saveData;
  }
});

// メールの CustomProperties を読み込む
function loadSavedData() {
  Office.context.mailbox.item.loadCustomPropertiesAsync((result) => {
    if (result.status !== Office.AsyncResultStatus.Succeeded) {
      console.error("CustomProperties 読み込み失敗");
      return;
    }

    const props = result.value;

    document.getElementById("assignedTo").value =
      props.get("AssignedTo") || "";

    document.getElementById("status").value =
      props.get("Status") || "未対応";

    document.getElementById("comment").value =
      props.get("Comment") || "";
  });
}

// 保存処理
function saveData() {
  Office.context.mailbox.item.loadCustomPropertiesAsync((result) => {
    if (result.status !== Office.AsyncResultStatus.Succeeded) {
      console.error("CustomProperties 読み込み失敗");
      return;
    }

    const props = result.value;

    // 値をセット
    props.set("AssignedTo", document.getElementById("assignedTo").value);
    props.set("Status", document.getElementById("status").value);
    props.set("Comment", document.getElementById("comment").value);

    // 保存 commit
    props.saveAsync((asyncResult) => {
      if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
        Office.context.mailbox.item.notificationMessages.replaceAsync(
          "saveSuccess",
          {
            type: "informationalMessage",
            message: "進捗を保存しました",
            icon: "icon16",
            persistent: false
          }
        );
      } else {
        alert("保存に失敗しました");
      }
    });
  });
}
