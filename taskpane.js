// MSAL 設定
const msalConfig = {
    auth: {
        clientId: "911ced4f-3287-4de9-a6e5-b706b323b2f5",
        authority: "https://login.microsoftonline.com/c1521ad5-8d23-4d3c-a5c4-e90a535b3c2a",
        redirectUri: "https://shion003-lab.github.io/mail-progress-tracker/taskpane.html"
    },
    cache: {
        cacheLocation: "localStorage",
        storeAuthStateInCookie: false
    }
};

// SharePoint 設定
const sharepointConfig = {
    siteUrl: "https://openloopcojp.sharepoint.com/sites/msteams_ed64e5",
    listName: "MailProgress"
};

// アクセス許可スコープ
const loginRequest = {
    scopes: ["Sites.ReadWrite.All", "User.Read"]
};

let msalInstance;
let currentUser = null;

Office.onReady((info) => {
    if (info.host === Office.HostType.Outlook) {
        console.log("Outlook Add-in Loaded");
        
        // MSAL インスタンスを初期化
        msalInstance = new msal.PublicClientApplication(msalConfig);
        
        // イベントリスナーを設定
        document.getElementById("loginButton").addEventListener("click", signIn);
        document.getElementById("saveButton").addEventListener("click", saveData);
        document.getElementById("refreshButton").addEventListener("click", refreshData);
        
        // 既存の認証状態を確認
        checkAuthStatus();
    }
});

/** 認証状態を確認 */
async function checkAuthStatus() {
    const accounts = msalInstance.getAllAccounts();
    
    if (accounts.length > 0) {
        currentUser = accounts[0];
        updateUIForSignedIn();
        await loadExistingData();
    } else {
        updateUIForSignedOut();
    }
}

/** サインイン処理 */
async function signIn() {
    try {
        const loginResponse = await msalInstance.loginPopup(loginRequest);
        currentUser = loginResponse.account;
        updateUIForSignedIn();
        await loadExistingData();
    } catch (error) {
        console.error("サインインエラー:", error);
        alert("サインインに失敗しました: " + error.message);
    }
}

/** UI を更新（サインイン済み） */
function updateUIForSignedIn() {
    document.getElementById("authStatus").textContent = `サインイン中: ${currentUser.username}`;
    document.getElementById("loginButton").classList.add("hidden");
    document.getElementById("mainContent").classList.remove("hidden");
}

/** UI を更新（サインアウト） */
function updateUIForSignedOut() {
    document.getElementById("authStatus").textContent = "認証が必要です";
    document.getElementById("loginButton").classList.remove("hidden");
    document.getElementById("mainContent").classList.add("hidden");
}

/** アクセストークンを取得 */
async function getAccessToken() {
    const account = msalInstance.getAllAccounts()[0];
    
    const request = {
        scopes: loginRequest.scopes,
        account: account
    };
    
    try {
        const response = await msalInstance.acquireTokenSilent(request);
        return response.accessToken;
    } catch (error) {
        console.log("サイレント取得失敗、ポップアップで再試行:", error);
        const response = await msalInstance.acquireTokenPopup(request);
        return response.accessToken;
    }
}

/** 現在のメールの一意IDを返す */
function getMailKey() {
    try {
        const mailId = Office.context.mailbox.item.internetMessageId;
        if (!mailId) {
            console.error("メールIDが取得できませんでした");
            return null;
        }
        return mailId;
    } catch (e) {
        console.error("メールID取得エラー:", e);
        return null;
    }
}

/** SharePoint リストから既存データを取得 */
async function loadExistingData() {
    const mailId = getMailKey();
    if (!mailId) return;
    
    showLoading(true);
    
    try {
        const accessToken = await getAccessToken();
        
        // SharePoint リストをクエリ
        const filter = `fields/MailID eq '${mailId}'`;
        const endpoint = `https://graph.microsoft.com/v1.0/sites/${await getSiteId()}/lists/${await getListId()}/items?$expand=fields&$filter=${encodeURIComponent(filter)}`;
        
        const response = await fetch(endpoint, {
            headers: {
                'Authorization': `Bearer ${accessToken}`
            }
        });
        
        if (!response.ok) {
            throw new Error(`HTTP ${response.status}: ${response.statusText}`);
        }
        
        const data = await response.json();
        
        if (data.value && data.value.length > 0) {
            const item = data.value[0].fields;
            
            if (item.AssignedTo) document.getElementById("assignedTo").value = item.AssignedTo;
            if (item.Status) document.getElementById("status").value = item.Status;
            if (item.Comment) document.getElementById("comment").value = item.Comment;
            
            console.log("データを読み込みました:", item);
        } else {
            console.log("このメールにはまだ保存データがありません");
            clearForm();
        }
    } catch (error) {
        console.error("データ読み込みエラー:", error);
        alert("データの読み込みに失敗しました: " + error.message);
    } finally {
        showLoading(false);
    }
}

/** SharePoint リストにデータを保存 */
async function saveData() {
    const mailId = getMailKey();
    if (!mailId) return;
    
    showLoading(true);
    
    const itemData = {
        fields: {
            MailID: mailId,
            AssignedTo: document.getElementById("assignedTo").value,
            Status: document.getElementById("status").value,
            Comment: document.getElementById("comment").value,
            UpdatedBy: currentUser.username,
            UpdatedAt: new Date().toISOString()
        }
    };
    
    try {
        const accessToken = await getAccessToken();
        
        // 既存のアイテムを検索
        const filter = `fields/MailID eq '${mailId}'`;
        const searchEndpoint = `https://graph.microsoft.com/v1.0/sites/${await getSiteId()}/lists/${await getListId()}/items?$expand=fields&$filter=${encodeURIComponent(filter)}`;
        
        const searchResponse = await fetch(searchEndpoint, {
            headers: {
                'Authorization': `Bearer ${accessToken}`
            }
        });
        
        const searchData = await searchResponse.json();
        
        let endpoint;
        let method;
        
        if (searchData.value && searchData.value.length > 0) {
            // 更新
            const itemId = searchData.value[0].id;
            endpoint = `https://graph.microsoft.com/v1.0/sites/${await getSiteId()}/lists/${await getListId()}/items/${itemId}/fields`;
            method = 'PATCH';
        } else {
            // 新規作成
            endpoint = `https://graph.microsoft.com/v1.0/sites/${await getSiteId()}/lists/${await getListId()}/items`;
            method = 'POST';
        }
        
        const response = await fetch(endpoint, {
            method: method,
            headers: {
                'Authorization': `Bearer ${accessToken}`,
                'Content-Type': 'application/json'
            },
            body: JSON.stringify(method === 'POST' ? itemData : itemData.fields)
        });
        
        if (!response.ok) {
            throw new Error(`HTTP ${response.status}: ${response.statusText}`);
        }
        
        alert("保存しました");
        console.log("データを保存しました");
        
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

/** SharePoint サイトIDを取得（キャッシュ） */
let cachedSiteId = null;
async function getSiteId() {
    if (cachedSiteId) return cachedSiteId;
    
    const accessToken = await getAccessToken();
    const siteUrl = sharepointConfig.siteUrl.replace('https://', '').replace('/sites/', ':/sites/');
    const endpoint = `https://graph.microsoft.com/v1.0/sites/${siteUrl}`;
    
    const response = await fetch(endpoint, {
        headers: {
            'Authorization': `Bearer ${accessToken}`
        }
    });
    
    const data = await response.json();
    cachedSiteId = data.id;
    return cachedSiteId;
}

/** SharePoint リストIDを取得（キャッシュ） */
let cachedListId = null;
async function getListId() {
    if (cachedListId) return cachedListId;
    
    const accessToken = await getAccessToken();
    const siteId = await getSiteId();
    const endpoint = `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${sharepointConfig.listName}`;
    
    const response = await fetch(endpoint, {
        headers: {
            'Authorization': `Bearer ${accessToken}`
        }
    });
    
    const data = await response.json();
    cachedListId = data.id;
    return cachedListId;
}
