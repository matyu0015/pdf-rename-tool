# GitHubでURL化する手順（初心者向け）

## ステップ1: GitHubで新しいリポジトリを作成

1. **GitHubにアクセス**
   - ブラウザで [https://github.com](https://github.com) を開く
   - 既にサインインしているので、そのまま進めます

2. **新しいリポジトリを作成**
   - 右上の「+」ボタンをクリック
   - 「New repository」を選択

3. **リポジトリの設定**
   - **Repository name**（リポジトリ名）: `pdf-rename-tool` と入力
     - （好きな名前でOK。英数字とハイフンのみ使用可）
   - **Description**（説明）: `PDFファイルの名前を変更するWebツール` と入力
     - （任意ですが、書いておくと分かりやすい）
   - **Public**（公開）を選択
     - GitHub Pagesは無料版ではPublicのみ使用可能
   - **「Add a README file」のチェックは外す**
     - 既にREADME.mdがあるため
   - **「Create repository」ボタンをクリック**

## ステップ2: ローカルのコードをGitHubにアップロード

リポジトリを作成すると、コマンドが表示されます。以下のコマンドを実行してください：

### 2-1. リモートリポジトリを追加

ターミナルで以下のコマンドを実行（`YOUR_USERNAME`を自分のGitHubユーザー名に置き換えてください）：

```bash
cd /Users/hinako/Desktop/PDF_name_change
git remote add origin https://github.com/YOUR_USERNAME/pdf-rename-tool.git
```

**例**: ユーザー名が `hinako123` の場合
```bash
git remote add origin https://github.com/hinako123/pdf-rename-tool.git
```

### 2-2. ブランチ名を変更（必要に応じて）

```bash
git branch -M main
```

### 2-3. GitHubにプッシュ

```bash
git push -u origin main
```

初回プッシュ時、GitHubのユーザー名とパスワード（またはパーソナルアクセストークン）を求められることがあります。

## ステップ3: GitHub Pagesを有効化

1. **GitHubのリポジトリページに戻る**
   - ブラウザで `https://github.com/YOUR_USERNAME/pdf-rename-tool` にアクセス

2. **Settings（設定）を開く**
   - リポジトリページの上部にある「Settings」タブをクリック

3. **Pagesの設定**
   - 左側のメニューから「Pages」をクリック
   - 「Branch」セクションで：
     - ドロップダウンから「main」を選択
     - フォルダは「/ (root)」のまま
     - 「Save」ボタンをクリック

4. **URLが発行される**
   - 数分後（通常は1〜2分）、ページ上部に以下のようなメッセージが表示されます：
   ```
   Your site is live at https://YOUR_USERNAME.github.io/pdf-rename-tool/
   ```
   - これがあなたのWebアプリのURLです！

## ステップ4: URLを確認してアクセス

1. 発行されたURLをクリックまたはコピー
2. ブラウザで開く
3. PDFリネームツールが表示されればOK！

## URLを仲間と共有

発行されたURL（例: `https://YOUR_USERNAME.github.io/pdf-rename-tool/`）を仲間に送れば、誰でもアクセスできます。

## トラブルシューティング

### 「404 Not Found」と表示される
- GitHub Pagesの有効化から数分待ってください
- ブラウザのキャッシュをクリアして再読み込み

### プッシュ時にエラーが出る
- GitHubのユーザー名とパーソナルアクセストークンが必要な場合があります
- [https://github.com/settings/tokens](https://github.com/settings/tokens) でトークンを作成してください

### ページは表示されるが動作しない
- ブラウザの開発者ツール（F12）でエラーを確認してください
- すべてのファイルが正しくアップロードされているか確認してください

## 更新方法（今後コードを変更した場合）

```bash
cd /Users/hinako/Desktop/PDF_name_change
git add .
git commit -m "更新内容の説明"
git push
```

数分後、GitHub Pagesに反映されます。

---

何か問題があれば、気軽に質問してください！
