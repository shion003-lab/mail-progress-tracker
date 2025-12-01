Office.onReady((info) => {
    if (info.host === Office.HostType.Outlook) {
        console.log("Outlook Add-in Loaded");
        document.getElementById("saveButton").addEventListener("click", saveData);
        document.getElementById("refreshButton").addEventListener("click", refreshData);
        
        // 初期データを読み込む
        loadExistingData();
        
        // メール切り替え検知をポーリングで実装
        setupMailChangeDetection();
    }
});

/** 前回チェックしたメールIDを保持 */
let lastMailId = null;

/** メール切り替え検知を設定（ポーリング方式） */
function setupMailChangeDetection() {
    try {
        // 初期メールIDを取得
        lastMailId = Office.context.mailbox.item.internetMessageId;
        
        // 1秒ごとにメールが切り替わったかチェック
        setInterval(checkMailChanged, 1000);
        console.log("メール切り替え検知を開始しました");
    } catch (e) {
        console.error("メール切り替え検知設定エラー:", e);
    }
}

/** メール切り替えをチェック */
function checkMailChanged() {
    try {
        const currentMailId = Office.context.mailbox.item.internetMessageId;
        
        // メールIDが変わった = メールが切り替わった
        if (currentMailId !== lastMailId) {
            console.log("メールが切り替わりました");
            console.log("前回:", lastMailId);
            console.log("現在:", currentMailId);
            
            // 前回のメールIDを更新
            lastMailId = currentMailId;
            
            // フォームをクリアして新しいメールのデータを読み込む
            clearForm();
            loadExistingData();
        }
    } catch (e) {
        console.error("メール切り替え検知エラー:", e);
    }
}

/** 更新ボタン押下時の処理 */
function refreshData() {
    console.log("=== 更新ボタンがクリックされました ===");
    console.log("現在の lastMailId:", lastMailId);
    console.log("localStorage の全キー:", Object.keys(localStorage));
    
    // Office.context を再確認
    if (!Office.context || !Office.context.mailbox || !Office.context.mailbox.item) {
        console.error("Office コンテキストが利用できません。ページをリロードします。");
        window.location.reload();
        return;
    }
    
    // 現在のメールIDを強制的に再取得
    const currentMailId = Office.context.mailbox.item.internetMessageId;
    console.log("更新時のメールID:", currentMailId);
    console.log("取得されるキー:", getMailKey());
    
    // 前回と異なる場合は lastMailId を更新
    if (currentMailId !== lastMailId) {
        console.log("✅ メールIDが変わっています:", lastMailId, "→", currentMailId);
        lastMailId = currentMailId;
    } else {
        console.log("❌ メールIDは変わっていません");
    }
    
    clearForm();
    loadExistingData();
    
    // フィードバック表示（Office コンテキストが有効な場合のみ）
    try {
        Office.context.mailbox.item.notificationMessages.replaceAsync("refreshInfo", {
            type: "informationalMessage",
            message: "データを更新しました。",
            icon: "icon16",
            persistent: false
        });
    } catch (e) {
        console.warn("通知表示に失敗しました:", e);
    }
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
