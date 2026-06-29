<p align="center">
  <img src="./assets/icon-80.png" alt="powerpogma icon" width="96" height="96">
</p>

<h1 align="center">powerpogma</h1>

<p align="center">
  PowerPointの図形をFigma風の右側パネルから編集できるOfficeアドインです。
</p>

<p align="center">
  <strong>日本語</strong>
  ·
  <a href="./README.en.md">English</a>
</p>

<p align="center">
  <a href="https://monjofight.github.io/powerpogma/src/taskpane.html">Taskpane</a>
  ·
  <a href="https://github.com/monjofight/powerpogma/raw/main/manifest.xml">manifest.xmlをダウンロード</a>
  ·
  <a href="#install-on-powerpoint-for-mac">Mac版PowerPointに追加</a>
</p>

PowerPointの図形を、Figmaの右側インスペクターのようなパネルから編集できるOfficeアドインです。

`powerpogma` は GitHub Pages 上の画面を読み込むため、通常利用ではローカルサーバーの起動は不要です。

## デモ

![powerpogma demo](./assets/demo/powerpogma-demo.gif)

[高画質版のmp4を開く](./assets/demo/powerpogma-demo.mp4)

## 主な機能

- PowerPointのタスクペインとして動くFigma風インスペクターUI
- 選択中の図形へのリアルタイム反映
- 位置の編集: `X`, `Y`
- サイズの編集: `W`, `H`
- 左寄せ、左右中央、右寄せ、上寄せ、上下中央寄せ、下寄せ
- 塗り色、塗りの不透明度
- 線の色、線の不透明度
- 複数図形選択時の混合値表示
- 複数図形選択時の横方向/縦方向の間隔調整
- 入力欄の全選択、上下キーでの数値増減

## Install On PowerPoint For Mac

この手順は **Mac版PowerPoint** 用です。

### 1. manifest.xml をダウンロード

次のファイルをダウンロードします。

[manifest.xml](https://github.com/monjofight/powerpogma/raw/main/manifest.xml)

### 2. PowerPointを終了

PowerPointを完全に終了します。

### 3. manifest.xml を wef フォルダへ配置

ターミナルで次を実行します。ブラウザでダウンロードしたファイル名が `manifest.xml` になっている場合の手順です。

```bash
mkdir -p ~/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef
cp ~/Downloads/manifest.xml ~/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef/7f4b4b6a-31c8-42da-b9f8-cb6d558c4f31.manifest.xml
```

または、ダウンロードせずに直接配置できます。

```bash
mkdir -p ~/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef
curl -L https://github.com/monjofight/powerpogma/raw/main/manifest.xml \
  -o ~/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef/7f4b4b6a-31c8-42da-b9f8-cb6d558c4f31.manifest.xml
```

### 4. PowerPointを起動

PowerPointを起動し、任意のプレゼンテーションを開きます。

### 5. powerpogmaを開く

PowerPointのリボンから次の順に開きます。

```text
ホーム > アドイン > その他のアドイン > 開発者向けアドイン > powerpogma
```

`powerpogma` を選ぶと、右側にタスクペインが表示されます。

## 重要

`ツール > PowerPoint アドイン` からは追加しません。

その画面は古いPowerPointアドイン形式用で、Office.jsの `manifest.xml` は選択できないことがあります。`powerpogma` は `ホーム > アドイン` から開きます。

## Uninstall On PowerPoint For Mac

PowerPointを終了してから、次を実行します。

```bash
rm -f ~/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef/7f4b4b6a-31c8-42da-b9f8-cb6d558c4f31.manifest.xml
```

その後PowerPointを再起動してください。

## 開発

開発や検証を行う場合のみ Node.js を使います。

```bash
npm install
npm run validate
```

ローカルでタスクペインを確認したい場合は、HTTPSサーバーを起動できます。

```bash
npm run start
```

ただし、通常利用の `manifest.xml` は GitHub Pages を参照しているため、PowerPointで使うだけなら `npm run start` は不要です。

## Office Add-in Toolingで追加する場合

Office Add-in toolingで直接サイドロードする場合は次を実行します。

```bash
npx --yes office-addin-debugging start manifest.xml desktop --app powerpoint --no-debug
```

この方法でもローカルサーバーは不要です。`manifest.xml` は GitHub Pages 上のタスクペインを読み込みます。

## ディレクトリ構成

```text
assets/          アドイン用アイコン
src/             タスクペインのHTML/CSS/JavaScript
manifest.xml     Office Add-in manifest
server.mjs       開発用ローカルHTTPS静的サーバー
PUBLISHING.md    公開用メモ
```

## メモ

- Taskpane URL: https://monjofight.github.io/powerpogma/src/taskpane.html
- Repository: https://github.com/monjofight/powerpogma
- PowerPoint JavaScript APIで扱える図形プロパティの範囲内で動作します。
- すべてのPowerPointオブジェクトや特殊な図形効果に対応しているわけではありません。

## ライセンス

MIT
