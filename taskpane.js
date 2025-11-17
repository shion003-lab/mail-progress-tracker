Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
      console.log("Office.js ready");

      // ← HTML の ID にあわせて変更
      document.getElementById("saveButton").onclick = saveProgress;
  }
});

async function saveProgress() {
  const progress = document.getElementById("progress").value;
  const comment = document.getElementById("comment").value;

  console.log("保存：", progress, comment);

  Office.context.mailbox.item.notificationMessages.replaceAsync(
    "saved",
    {
      type: "informationalMessage",
      message: "保存処理（仮）が動きました",
      icon: "icon16",
      persistent: false
    }
  );
}
