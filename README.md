# PaTakaRUSH (旧: PaTa_260406データ分析)

Google Apps Script製のダッシュボードアプリを GitHub で管理するために移植したプロジェクトです。

## 構成

- `appsscript.json` — GASのマニフェスト
- `コード.gs` — サーバーサイドロジック
- `index.html` / `dashboard*.html` / `dialog.html` — クライアント側テンプレート
- `.clasp.json` — 紐付け先のGASスクリプトID

## GAS本体との同期（GitHub Actions）

`main` ブランチへのpush時に、`.github/workflows/clasp-push.yml` が自動で
`clasp push` を実行し、このリポジトリの内容をGoogle Apps Scriptプロジェクトへ反映します。

### 初回セットアップ（ローカルで一度だけ）

1. `npm install -g @google/clasp`
2. `clasp login` （ブラウザでGoogleアカウント認証）
3. ログイン後に生成される `~/.clasprc.json` の中身をコピー
4. GitHubリポジトリの Settings > Secrets and variables > Actions で
   `CLASPRC_JSON` という名前のSecretを作成し、その内容を貼り付け

これでpush時に自動でGAS側へ反映されるようになります。

## ローカルでの手動push

```bash
npm install -g @google/clasp
clasp login
clasp push
```
