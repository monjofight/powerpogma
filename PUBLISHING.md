# GitHub公開メモ

このリポジトリはGitHubへ公開できる状態です。公開前にREADMEとmanifestの内容が用途に合っているか確認してください。

## GitHub CLIで新しい公開リポジトリを作る

```bash
gh repo create pptx-figma-inspector --public --source=. --remote=origin --push
```

## 既存リポジトリへ接続する

```bash
git remote add origin git@github.com:<OWNER>/pptx-figma-inspector.git
git push -u origin main
```

## 公開前チェック

- 開発・サイドロード用途なら、`manifest.xml` の `https://localhost:3000` はそのままで問題ありません。
- 本番ホストやAppSource公開を想定する場合は、`manifest.xml` 内のlocalhost URLを公開済みHTTPS URLへ差し替えてください。
- `ProviderName`、サポートURL、アドイン説明文、アイコン、スクリーンショットが公開用途に合っているか確認してください。
- `npm run validate` が通ることを確認してください。
- `work/`、`node_modules/`、`.DS_Store` などのローカル生成物がコミット対象になっていないことを確認してください。
