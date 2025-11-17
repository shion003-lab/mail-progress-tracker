<!DOCTYPE html>
<html lang="ja">
<head>
  <meta charset="UTF-8">
  <title>MailPM - 進捗管理</title>

  <!-- Office.js を必ず最上部で読み込む -->
  <script src="https://appsforoffice.microsoft.com/lib/1.1/hosted/office.js"></script>

  <style>
    body {
      font-family: "Segoe UI", sans-serif;
      background-color: #fafafa;
      margin: 0;
      padding: 0;
    }
    .container {
      border-top: 1px solid #ddd;
      padding: 16px;
      background-color: #fff;
    }
    h2 { font-size: 16px; margin-bottom: 10px; }
    select, textarea, button {
      width: 100%;
      margin-top: 6px;
      margin-bottom: 12px;
      padding: 6px;
    }
    button {
      background-color: #0078d4;
      color: white;
      border: none;
      border-radius: 4px;
    }
  </style>
</head>
<body>
  <div class="container">
    <h2>📊 メール進捗管理</h2>

    <label for="progress">進捗ステータスを変更:</label>
    <select id="progress">
      <option value="未着手">未着手</option>
      <option value="進行中">進行中</option>
      <option value="完了">完了</option>
      <option value="保留">保留</option>
    </select>

    <label for="comment">コメント（任意）:</label>
    <textarea id="comment"></textarea>

    <button id="saveButton">保存</button>
  </div>

  <!-- taskpane.js をここで読み込む -->
  <script src="taskpane.js"></script>
</body>
</html>
