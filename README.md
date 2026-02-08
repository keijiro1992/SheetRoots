# SheetRoots 🌳

Google スプレッドシートの依存関係を可視化するツールです。  
`IMPORTRANGE` による外部参照や、シート間の内部参照を解析してグラフ表示します。

## 機能

- スプレッドシート ID または URL を入力して解析
- 内部シート参照の検出
- IMPORTRANGE による外部参照の検出
- 依存関係のグラフ可視化

## 技術スタック

- Google Apps Script (GAS)
- Vanilla JavaScript + SVG
- clasp (ローカル開発用)

## セットアップ

### 1. clasp のインストール

```bash
npm install -g @google/clasp
clasp login
```

### 2. プロジェクト作成

```bash
clasp create --type standalone --title "SheetRoots"
clasp push
```

### 3. デプロイ

1. [スクリプトエディタ](https://script.google.com/) でプロジェクトを開く
2. 「デプロイ」→「新しいデプロイ」→「ウェブアプリ」
3. アクセス権限を設定してデプロイ

## ファイル構成

```
SheetRoots/
├── appsscript.json  # GAS マニフェスト
├── Code.gs          # メインロジック + doGet
├── Parser.gs        # IMPORTRANGE パーサー
└── Index.html       # フロントエンド UI
```

## ライセンス

MIT
