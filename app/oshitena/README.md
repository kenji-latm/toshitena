# オシテナ（捺印欄スタンプ）– Wordアドイン

押印欄の点線円をワンクリックで挿入するWordアドイン（試作版）。
司法書士が委任状・契約書・議事録などに手作業で描いている「捺印の丸」をなくすためのツール。

## 構成

| ファイル | 役割 |
|---|---|
| `taskpane.html` / `taskpane.js` | タスクペインUI（サイズ・文字・線・個数を選んで挿入） |
| `ooxml.js` | 点線円のOOXML（DrawingML＋VMLフォールバック）を生成 |
| `manifest.xml` | アドインのマニフェスト（Wordに読み込ませる定義ファイル） |
| `index.html` | 配布用ランディングページ |
| `assets/icon-*.png` | リボン・ストア用アイコン（`tools/make-icons.mjs` で再生成可能） |

## 仕組み

- Office.js（Word JavaScript API）の `insertOoxml` で、カーソル位置に浮動図形（楕円）を挿入する
- WordApi 1.1 の範囲だけを使っているため、**Word 2016以降・Mac・Word on the web すべてで動く**
- 挿入されるのはWord標準の図形なので、手描きの円と同様にドラッグ・削除でき、
  電子署名前の一括削除（denshi-shomei-shitagoshirae）でもそのまま消せる

## 動作要件

- Word 2016以降 / Microsoft 365 / Word on the web
- タスクペインはHTTPSでホストされている必要がある（GitHub Pagesで配信）

## URLについて

現在のマニフェストは `https://kenji-latm.github.io/toshitena/oshitena/` を参照している
（このリポジトリの Pages は `app/` を公開ルートとして自動デプロイされるため、mainにマージすれば配信される）。

`tools.ishimoto-legal.com` 配下（agetenaリポジトリ）へ移す場合：

1. `app/oshitena/` を agetena の `app/oshitena/` へコピー
2. `manifest.xml` 内の `kenji-latm.github.io/toshitena` を `tools.ishimoto-legal.com` に一括置換
3. `<Version>` を上げる（例：0.1.0.0 → 0.2.0.0）

## テスト（サイドロード）

### Word on the web（いちばん簡単）

1. ブラウザでWord文書を開く
2. リボンの「アドイン」→「その他のアドイン」→「マイ アドイン」→「マイ アドインのアップロード」
3. `manifest.xml` を選択 → ホームタブ右端に「捺印欄」ボタンが出る

### デスクトップ版Word（Windows）

個人でのサイドロードは共有フォルダーカタログを使う：

1. 適当なフォルダー（例 `C:\addins`）を作り、フォルダーの共有を有効化（自分だけでOK）
2. `manifest.xml` をそこに置く
3. Wordの「ファイル」→「オプション」→「セキュリティ センター」→「セキュリティ センターの設定」→「信頼できるアドイン カタログ」に `\\PC名\addins` を追加し「メニューに表示する」にチェック
4. Word再起動 →「アドイン」→「その他のアドイン」→「共有フォルダー」タブから追加

Microsoft 365テナントがあれば、管理センターの「統合アプリ」からの展開（集中展開）も使える。

## 一般公開（AppSource）

日本中の誰でも「アドインを入手」から追加できるようにするには、Microsoft AppSource に掲載する：

1. [Microsoft Partner Center](https://partner.microsoft.com/) でアカウント登録（Microsoft 365・Copilotプログラムへの参加は無料。事業者確認あり）
2. 「Marketplace offers」で新しい Office アドインのオファーを作成し、`manifest.xml` を提出
3. [検証ポリシー](https://learn.microsoft.com/ja-jp/legal/marketplace/certification-policies)に基づく審査（数営業日〜）
   - プライバシーポリシーページのURLが必須（`SupportUrl` とは別に用意する）
   - スクリーンショット、日本語の説明文などのストア掲載情報が必要
4. 承認後、Wordの「アドインを入手」検索とAppSourceサイトに掲載される。無料アドインなら掲載料もなし

審査前に `npx office-addin-manifest validate app/oshitena/manifest.xml` でマニフェストを検証しておくとよい。

## 既知の制限・今後の課題

- 行内（インライン）配置は未対応（現状は前面・アンカー基準の浮動配置のみ）
- 角印（正方形）・二重円・「契印」用の割円などは未実装
- AppSource提出にはプライバシーポリシーページの用意が必要
