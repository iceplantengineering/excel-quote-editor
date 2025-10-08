# Netlifyデプロイ手順

## 1. GitHubリポジトリの作成

1. [GitHub](https://github.com)にログイン
2. 新しいリポジトリを作成
   - Repository name: `excel-quote-editor`
   - Description: `DeepSeek AIを使用してExcelファイルを自然言語で編集できるWebアプリケーション`
   - Public または Private を選択
   - README, .gitignore, license は追加しない（既に存在するため）

## 2. ローカルリポジトリをGitHubにプッシュ

```bash
# リモートリポジトリを追加（your-usernameを実際のGitHubユーザー名に変更）
git remote add origin https://github.com/your-username/excel-quote-editor.git

# メインブランチにプッシュ
git branch -M main
git push -u origin main
```

## 3. Netlifyでのデプロイ設定

1. [Netlify](https://netlify.com)にログイン
2. "New site from Git" をクリック
3. GitHubを選択し、認証
4. `excel-quote-editor` リポジトリを選択
5. デプロイ設定を確認：
   - **Branch to deploy**: `main`
   - **Build command**: `npm run build`
   - **Publish directory**: `dist`
   - **Node version**: `18` (Environment variables で `NODE_VERSION=18` を設定)

## 4. 環境変数の設定（オプション）

Netlifyの管理画面で以下を設定：
- Site settings > Environment variables
- 必要に応じて環境変数を追加

## 5. カスタムドメインの設定（オプション）

1. Site settings > Domain management
2. "Add custom domain" でドメインを追加
3. DNS設定を更新

## 6. 自動デプロイの確認

- GitHubにプッシュするたびに自動でデプロイされます
- Deploy log でビルド状況を確認できます

## トラブルシューティング

### ビルドエラーが発生した場合

1. **Node.jsバージョン**: Environment variables で `NODE_VERSION=18` を設定
2. **依存関係**: `package.json` の dependencies を確認
3. **ビルドログ**: Netlify の Deploy log を確認

### 404エラーが発生した場合

- `netlify.toml` でSPAのリダイレクト設定済み
- ファイルが正しく配置されているか確認

## 設定ファイルの説明

- **netlify.toml**: Netlify用のビルド設定とリダイレクト設定
- **.gitignore**: Git管理から除外するファイル
- **README.md**: プロジェクトの説明とドキュメント

## デプロイ後の確認事項

1. ✅ ファイルアップロード機能
2. ✅ DeepSeek API接続テスト
3. ✅ データプレビュー表示
4. ✅ ズーム機能
5. ✅ フルスクリーン表示
6. ✅ Excel編集機能
7. ✅ ファイルダウンロード機能