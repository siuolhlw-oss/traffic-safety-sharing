# 交通安全訪視經驗分享

這是一套無後台的靜態網頁簡報與 PowerPoint，從學校三次交通安全訪視經驗出發，討論評鑑效度、Safe System、人本交通、教育工作者的合理角色，以及對教授與主管機關的制度倡議。

## 本機觀看網頁版

直接開啟 `index.html` 即可。操作方式：

- `→`、`PageDown`、空白鍵：下一頁
- `←`、`PageUp`：上一頁
- `Home`／`End`：第一頁／最後一頁
- `F`：全螢幕
- `P`：列印或輸出 PDF

## 部署到 GitHub Pages

本專案不需要後台、資料庫或建置服務。建立 GitHub repo 並推送後：

1. 到 repo 的 **Settings → Pages**。
2. 在 **Build and deployment** 選 **Deploy from a branch**。
3. Branch 選 `main`，資料夾選 `/ (root)`，按 **Save**。
4. 等候 GitHub 顯示公開網址。

若尚未建立 remote，可在 GitHub 建立空 repo 後執行：

```powershell
git remote add origin https://github.com/你的帳號/你的-repo.git
git add .
git commit -m "feat: add traffic safety keynote and web deck"
git push -u origin main
```

## 專案檔案

- `index.html`：網頁簡報內容
- `styles.css`：16:9 舞台、排版與列印樣式
- `script.js`：鍵盤、觸控與進度導覽
- `GOALS.md`：演講與專案目標
- `AGENTS.md`：後續修改、驗證、commit 與 push 規範
- `SOURCES.md`：資料與影片來源
- `outputs/`：PowerPoint 交付檔

