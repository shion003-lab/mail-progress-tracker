Office.onReady((info) => {
    if (info.host === Office.HostType.Outlook) {
        console.log("Outlook Add-in Loaded");
        document.getElementById("saveButton").addEventListener("click", saveData);
        
        // 初期データを読み込む
        loadExistingData();
        
        // メール切り替え時に自動更新するイベントリスナーを設定
        setupItemChangeListener();
    }
});

/** メール切り替え時のイベントリスナーを設定 */
function setupItemChangeListener() {
    try {
        Office.context.mailbox.item.addHandlerAsync(Office.EventType.ItemChanged, onItemChanged);
        console.log("ItemChanged イベントリスナーを設定しました");
    } catch (e) {
        console.error("イベントリスナー設定エラー:", e);
    }
}

/** メール切り替え時に自動実行されるハンドラー */
function onItemChanged(eventArgs) {
    console.log("メールが切り替わりました");
    // フォームをクリアして新しいメールのデータを読み込む
    clearForm();
    loadExistingData();
}

/** フォームをリセット */
function clearForm() {
    document.getElementById("assignedTo").value = "";
    document.getElementById("status").value = "未対応";
    document.getElementById("comment").value = "";
}

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
        
        // 成功通知を表示
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
        if (!raw) {
            console.log("このメールにはまだ保存データがありません");
            return;
        }
        
        const data = JSON.parse(raw);
        if (data.assignedTo) document.getElementById("assignedTo").value = data.assignedTo;
        if (data.status) document.getElementById("status").value = data.status;
        if (data.comment) document.getElementById("comment").value = data.comment;
        
        console.log("データを読み込みました:", data);
    } catch (e) {
        console.error("読み込みエラー:", e);
    }
}
