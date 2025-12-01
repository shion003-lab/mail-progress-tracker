// Firebase 設定
const firebaseConfig = {
    apiKey: "AIzaSyAWM90N8W3QM4HoX_SEU10Rt8UTgKj4nrU",
    authDomain: "mail-progress-tracker.firebaseapp.com",
    projectId: "mail-progress-tracker",
    storageBucket: "mail-progress-tracker.firebasestorage.app",
    messagingSenderId: "1047658950318",
    appId: "1:1047658950318:web:5cc67b7897bb1d590a073f"
};

// Firebase 初期化
let db;
let auth;
let currentUser = null;

Office.onReady((info) => {
    if (info.host === Office.HostType.Outlook) {
        console.log("Outlook Add-in Loaded");
        
        // Firebase を初期化
        firebase.initializeApp(firebaseConfig);
        db = firebase.firestore();
        auth = firebase.auth();
        
        // 認証状態の変更を監視
        auth.onAuthStateChanged((user) => {
            if (user) {
                currentUser = user;
                updateUIForSignedIn();
                loadExistingData();
            } else {
                currentUser = null;
                updateUIForSignedOut();
            }
        });
        
        // イベントリスナーを設定
        document.getElementById("loginButton").addEventListener("click", signIn);
        document.getElementById("saveButton").addEventListener("click", saveData);
        document.getElementById("refreshButton").addEventListener("click", refreshData);
        
        // Enterキーでサインイン
        document.getElementById("passwordInput").addEventListener("keypress", (e) => {
            if (e.key === 'Enter') signIn();
        });
    }
});

/** メール/パスワードでサインイン */
async function signIn() {
    const email = document.getElementById("emailInput").value.trim();
    const password = document.getElementById("passwordInput").value;
    
    if (!email || !password) {
        alert("メールアドレスとパスワードを入力してください");
        return;
    }
    
    try {
        await auth.signInWithEmailAndPassword(email, password);
        console.log("サインインに成功しました");
    } catch (error) {
        console.error("サインインエラー:", error);
        
        // エラーメッセージを日本語で表示
        let errorMessage = "サインインに失敗しました";
        if (error.code === 'auth/user-not-found') {
            errorMessage = "このメールアドレスは登録されていません";
        } else if (error.code === 'auth/wrong-password') {
            errorMessage = "パスワードが正しくありません";
        } else if (error.code === 'auth/invalid-email') {
            errorMessage = "メールアドレスの形式が正しくありません";
        }
        
        alert(errorMessage);
    }
}

/** UI を更新（サインイン済み） */
function updateUIForSignedIn() {
    const displayName = currentUser.displayName || currentUser.email;
    document.getElementById("authStatus").textContent = `サインイン中: ${displayName}`;
    document.getElementById("authSection").classList.add("hidden");
    document.getElementById("mainContent").classList.remove("hidden");
}

/** UI を更新（サインアウト） */
function updateUIForSignedOut() {
    document.getElementById("authStatus").textContent = "サインインしてください";
    document.getElementById("authSection").classList.remove("hidden");
    document.getElementById("mainContent").classList.add("hidden");
}

/** 現在のメールの一意IDを返す */
function getMailKey() {
    try {
        const mailId = Office.context.mailbox.item.internetMessageId;
        if (!mailId) {
            console.error("メールIDが取得できませんでした");
            return null;
        }
        // メールIDをFirestoreのドキュメントIDとして使えるように正規化
        return mailId.replace(/[<>]/g, '').replace(/[@.]/g, '_');
    } catch (e) {
        console.error("メールID取得エラー:", e);
        return null;
    }
}

/** Firestore から既存データを取得 */
async function loadExistingData() {
    const mailId = getMailKey();
    if (!mailId) return;
    
    showLoading(true);
    
    try {
        const docRef = db.collection('mailProgress').doc(mailId);
        const doc = await docRef.get();
        
        if (doc.exists) {
            const data = doc.data();
            
            if (data.assignedTo) document.getElementById("assignedTo").value = data.assignedTo;
            if (data.status) document.getElementById("status").value = data.status;
            if (data.comment) document.getElementById("comment").value = data.comment;
            
            // 最終更新情報を表示
            if (data.updatedBy && data.updatedAt) {
                const updatedDate = new Date(data.updatedAt.toDate()).toLocaleString('ja-JP');
                document.getElementById("lastUpdatedInfo").textContent = 
                    `最終更新: ${updatedDate} by ${data.updatedBy}`;
            }
            
            console.log("データを読み込みました:", data);
        } else {
            console.log("このメールにはまだ保存データがありません");
            clearForm();
            document.getElementById("lastUpdatedInfo").textContent = "";
        }
    } catch (error) {
        console.error("データ読み込みエラー:", error);
        alert("データの読み込みに失敗しました: " + error.message);
    } finally {
        showLoading(false);
    }
}

/** Firestore にデータを保存 */
async function saveData() {
    if (!currentUser) {
        alert("サインインが必要です");
        return;
    }
    
    const mailId = getMailKey();
    if (!mailId) return;
    
    showLoading(true);
    
    const data = {
        mailId: Office.context.mailbox.item.internetMessageId,
        assignedTo: document.getElementById("assignedTo").value,
        status: document.getElementById("status").value,
        comment: document.getElementById("comment").value,
        updatedBy: currentUser.displayName || currentUser.email,
        updatedAt: firebase.firestore.FieldValue.serverTimestamp()
    };
    
    try {
        const docRef = db.collection('mailProgress').doc(mailId);
        await docRef.set(data, { merge: true });
        
        alert("保存しました");
        console.log("データを保存しました:", data);
        
        // 最終更新情報を即座に更新
        const now = new Date().toLocaleString('ja-JP');
        document.getElementById("lastUpdatedInfo").textContent = 
            `最終更新: ${now} by ${data.updatedBy}`;
            
    } catch (error) {
        console.error("保存エラー:", error);
        alert("保存に失敗しました: " + error.message);
    } finally {
        showLoading(false);
    }
}

/** 更新ボタン押下時の処理 */
function refreshData() {
    console.log("=== 更新ボタンがクリックされました ===");
    window.location.reload();
}

/** フォームをリセット */
function clearForm() {
    document.getElementById("assignedTo").value = "";
    document.getElementById("status").value = "未対応";
    document.getElementById("comment").value = "";
}

/** ローディング表示を切り替え */
function showLoading(show) {
    if (show) {
        document.getElementById("loadingIndicator").classList.remove("hidden");
        document.getElementById("mainContent").style.opacity = "0.5";
    } else {
        document.getElementById("loadingIndicator").classList.add("hidden");
        document.getElementById("mainContent").style.opacity = "1";
    }
}
