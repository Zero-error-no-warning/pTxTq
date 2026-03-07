# PptxThemeDocument

`.pptx` から抽出した編集可能な PresentationML モデルを扱うための、`src/` 中心のリポジトリです。

English version: [README.md](./README.md)

## 公開スコープ

このリポジトリでは、主に `src/` 以下の実装を公開対象とします。

以下はローカル用途または検証用途として扱い、特別な理由がない限り公開履歴には含めない想定です。

- `scripts/`
- `docs/`
- サンプル `.pptx` ファイル
- ローカル履歴やワークスペース用メタデータ

## `src/` の構成

- `src/core/`: パッケージ読み込み、テーマ解決、PresentationML の関係グラフ走査、編集用モデル
- `src/render/`: メモリ上のモデルから SVG / Canvas を描画
- `src/write/`: 編集済みモデルを `.pptx` に書き戻し
- `src/utils/`: XML、単位変換、色、ジオメトリ、オブジェクト補助

## 現在できること

- `.pptx` を編集可能なモデルとして読み込む
- テーマとスライド間の参照関係を解決する
- スライドを SVG として描画する
- スライドを Canvas に描画する
- あるスライドを別スライドへ編集可能な要素として埋め込む
- 編集したモデルを `.pptx` に書き戻す

## 公開 API の入口

`src/index.js`

```js
import { PptxThemeDocument } from "./src/index.js";

const doc = await PptxThemeDocument.load("input.pptx");
const svg = await doc.renderSlide(0, { mode: "svg" });
await doc.saveAs("roundtrip.pptx");
```

## 補足

- `src/` の実装は `jszip` と `fast-xml-parser` に依存します。
- 検証補助スクリプトは追加で Node、Python、PowerShell などに依存する場合がありますが、それらはこのリポジトリの公開スコープ外です。
