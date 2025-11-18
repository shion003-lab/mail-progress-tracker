Office.onReady((info) => {
    if (info.host === Office.HostType.Outlook) {
        console.log("Outlook Add-in Loaded");
        document.getElementById("saveButton").addEventListener("click", saveData);

        loadExistingData();
    }
});

/** 現在のメールの一意IDを返す */
function getMailKey() {
    try {
        const mailId = Office.context.mailbox.item.internetMessageId;
        if (!mailId) {
            console.error("メールIDが取得できませんでした");
            return null;
        }
        return "mail_" + mailId;
    } catch (e) {
        console.error("メールID取得エラー:", e);
        return null;
    }
}

/** 保存処理 */
function saveData() {
    const key = getMailKey();
    if (!key) return;

    const data = {
        assignedTo: document.getElementById("assignedTo").value,
        status: document.getElementById("status").value,
        comment: document.getElementById("comment").value,
        updatedAt: new Date().toISOString(),
    };

    try {
        localStorage.setItem(key, JSON.stringify(data));
        Office.context.ui.displayDialogAsync;

        Office.context.mailbox.item.notificationMessages.replaceAsync("saveInfo", {
            type: "informationalMessage",
            message: "保存しました。",
            icon: "icon16",
            persistent: false
        });
    } catch (e) {
        console.error("保存エラー:", e);

        Office.context.mailbox.item.notificationMessages.replaceAsync("saveError", {
            type: "errorMessage",
            message: "保存に失敗しました。"
        });
    }
}

/** 保存済みデータを読み込む */
function loadExistingData() {
    const key = getMailKey();
    if (!key) return;

    try {
        const raw = localStorage.getItem(key);
        if (!raw) return;

        const data = JSON.parse(raw);

        if (data.assignedTo) document.getElementById("assignedTo").value = data.assignedTo;
        if (data.status) document.getElementById("status").value = data.status;
        if (data.comment) document.getElementById("comment").value = data.comment;

    } catch (e) {
        console.error("読み込みエラー:", e);
    }
}
