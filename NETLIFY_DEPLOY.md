# 🚀 Netlifyデプロイ完全ガイド

## ✅ 準備完了状況

- ✅ プロジェクトビルド成功（777KB）
- ✅ Netlify設定ファイル（netlify.toml）作成済み
- ✅ Git リポジトリ初期化済み
- ✅ 依存関係の問題解決済み
- ✅ 本番環境用設定完了

## 📋 デプロイ手順

### 1. GitHubリポジトリ作成

1. [GitHub](https://github.com)にログイン
2. 右上の「+」→「New repository」をクリック
3. 以下の設定で作成：
   ```
   Repository name: excel-quote-editor
   Description: DeepSeek AIを使用してExcelファイルを自然言語で編集できるWebアプリケーション
   Public または Private を選択
   ✅ Add a README file のチェックを外す
   ✅ Add .gitignore のチェックを外す
   ✅ Choose a license のチェックを外す
   ```

### 2. ローカルリポジトリをGitHubにプッシュ

```bash
# 現在のディレクトリ: /workspace/excel_quote_editor

# リモートリポジトリを追加（your-usernameを実際のGitHubユーザー名に変更）
git remote add origin https://github.com/your-username/excel-quote-editor.git

# メインブランチにプッシュ
git branch -M main
git push -u origin main
```

### 3. Netlifyでのデプロイ設定

1. [Netlify](https://netlify.com)にログイン
2. 「New site from Git」をクリック
3. 「GitHub」を選択し、認証を完了
4. `excel-quote-editor` リポジトリを選択
5. デプロイ設定を確認：
   ```
   Branch to deploy: main
   Build command: npm run build
   Publish directory: dist
   ```
6. 「Deploy site」をクリック

### 4. 環境変数設定（推奨）

Netlifyの管理画面で：
1. Site settings → Environment variables
2. 以下を追加（オプション）：
   ```
   NODE_VERSION=18
   NPM_VERSION=8
   ```

### 5. カスタムドメイン設定（オプション）

1. Site settings → Domain management
2. 「Add custom domain」でドメインを追加
3. DNS設定を更新

## 🔧 設定ファイルの説明

### netlify.toml
```toml
[build]
  command = "npm run build"
  publish = "dist"
  
[build.environment]
  NODE_VERSION = "18"

[[redirects]]
  from = "/*"
  to = "/index.html"
  status = 200
```

### package.json（主要設定）
```json
{
  "name": "excel-quote-editor",
  "version": "1.0.0",
  "scripts": {
    "build": "vite build",
    "preview": "vite preview"
  }
}
```

## 🎯 デプロイ後の確認事項

### 基本機能テスト
- [ ] ページが正常に読み込まれる
- [ ] DeepSeek API設定エリアが表示される
- [ ] ファイルアップロード機能が動作する
- [ ] データプレビューが表示される
- [ ] ズーム機能が動作する
- [ ] フルスクリーン表示が動作する

### API機能テスト（APIキー必要）
- [ ] DeepSeek API接続テストが成功する
- [ ] 自然言語編集指示が実行される
- [ ] 編集履歴が記録される
- [ ] Undo機能が動作する
- [ ] 書式保持でダウンロードできる

## 🚨 トラブルシューティング

### ビルドエラーが発生した場合

**問題**: `npm run build` が失敗する
**解決策**:
1. Node.jsバージョンを18に設定
2. Environment variables で `NODE_VERSION=18` を追加
3. Deploy log を確認してエラー詳細を確認

**問題**: 依存関係エラー
**解決策**:
```bash
# package-lock.json を削除して再インストール
rm package-lock.json
npm install
npm run build
```

### 404エラーが発生した場合

**問題**: ページが404エラーになる
**解決策**:
- `netlify.toml` のリダイレクト設定を確認
- Publish directory が `dist` に設定されているか確認

### API接続エラーが発生した場合

**問題**: DeepSeek API接続が失敗する
**解決策**:
- CORS設定を確認
- APIキーの有効性を確認
- ブラウザの開発者ツールでネットワークエラーを確認

## 📊 パフォーマンス最適化

### バンドルサイズ最適化
現在のバンドルサイズ: 777KB（gzip: 255KB）

改善案：
1. 動的インポートでコード分割
2. 未使用の依存関係を削除
3. Tree shaking の最適化

### CDN設定
静的アセットのキャッシュ設定済み：
```toml
[[headers]]
  for = "/assets/*"
  [headers.values]
    Cache-Control = "public, max-age=31536000, immutable"
```

## 🔒 セキュリティ設定

セキュリティヘッダー設定済み：
```toml
[[headers]]
  for = "/*"
  [headers.values]
    X-Frame-Options = "DENY"
    X-XSS-Protection = "1; mode=block"
    X-Content-Type-Options = "nosniff"
    Referrer-Policy = "strict-origin-when-cross-origin"
```

## 📞 サポート

問題が発生した場合：
1. GitHub Issues で報告
2. Netlify Deploy log を確認
3. ブラウザの開発者ツールでエラーを確認

---

**🎉 デプロイ成功後、あなたのExcel見積書エディターが世界中からアクセス可能になります！**