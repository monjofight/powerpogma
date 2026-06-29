# powerpogma

PowerPointの図形を、Figmaの右側インスペクターのようなパネルから編集できるOfficeアドインです。

選択中の図形に対して、位置、サイズ、回転、整列、塗り、線、不透明度、複数選択時の間隔調整などを数値入力で操作できます。ローカルでのサイドロードと開発利用を前提にしています。

## 主な機能

- PowerPointのタスクペインとして動くFigma風インスペクターUI
- 選択中の図形へのリアルタイム反映
- 位置の編集: `X`, `Y`
- サイズの編集: `W`, `H`
- 回転と90度回転
- 左寄せ、中央寄せ、右寄せ、上寄せ、上下中央寄せ、下寄せ
- 塗り色、塗りの不透明度
- 線の色、線の不透明度
- 複数図形選択時のバウンディングボックス編集
- 複数図形選択時の混合値表示
- 横方向、縦方向の間隔調整
- ブラウザ単体でのデモ表示

## 必要なもの

- Microsoft PowerPoint
- Office Add-insを利用できる環境
- Node.js 18以降
- OpenSSL

## セットアップ

```bash
npm install
npm run start
```

ローカルHTTPSサーバーが起動します。

```text
https://localhost:3000/src/taskpane.html
```

初回起動時に `work/certs/` 以下へローカル開発用の自己署名証明書を生成します。ブラウザやPowerPointで証明書の警告が出る場合は、開発用URLとして許可してください。

ポートを変更したい場合は `PORT` を指定できます。

```bash
PORT=3001 npm run start
```

その場合は `manifest.xml` 内の `https://localhost:3000` も同じポートに変更してください。

## PowerPointで読み込む

開発サーバーを起動した状態で、`manifest.xml` をPowerPointへサイドロードします。

Office Add-in toolingを使う場合は次のコマンドを実行します。

```bash
npx office-addin-debugging start manifest.xml desktop --app powerpoint --no-debug --dev-server-port 3000
```

読み込み後、PowerPointのホームタブに表示される `powerpogma` ボタンからタスクペインを開きます。

## 開発

```bash
npm run validate
```

このコマンドでは次の検証を行います。

- `src/taskpane.js` の構文チェック
- `manifest.xml` のOffice Add-in manifest検証

GitHub Actionsでも同じ検証を実行します。

## ディレクトリ構成

```text
assets/          アドイン用アイコン
src/             タスクペインのHTML/CSS/JavaScript
manifest.xml     Office Add-in manifest
server.mjs       ローカルHTTPS静的サーバー
PUBLISHING.md    GitHub公開時のメモ
```

## 注意点

- 現在の `manifest.xml` はローカル開発用に `https://localhost:3000` を参照しています。
- GitHubで公開するだけならlocalhostのままでも問題ありません。
- Microsoft AppSourceなどへ公開する場合は、`manifest.xml` のURLを公開済みHTTPS URLへ差し替えてください。
- AppSource公開時は `ProviderName`、サポートURL、アイコン、説明文、スクリーンショットなども実運用向けに更新してください。
- PowerPoint JavaScript APIで扱える図形プロパティの範囲内で動作します。すべてのPowerPointオブジェクトや特殊な図形効果に対応しているわけではありません。

## ライセンス

MIT
