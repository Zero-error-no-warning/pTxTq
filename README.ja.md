# JavaScript PPTX to Canvas / PPTX to PPTX

PowerPoint の `.pptx` を扱うための、モデルベースの JavaScript 実装です。

English version: [README.md](./README.md)

このリポジトリは、次の 2 つのワークフローを主眼にしています。

- `pptx -> canvas`
- `pptx -> pptx`

`.pptx` パッケージを読み込んで編集可能な PresentationML モデルへ変換し、そのモデルからスライドを描画し、編集後の内容を新しい `.pptx` として書き戻せます。

## できること

- ファイルパス、`Buffer`、`Uint8Array`、`ArrayBuffer` から `.pptx` を読み込む
- 編集可能なスライドモデルをメモリ上に構築する
- スライドを Canvas に描画する
- スライドを SVG に描画する
- あるスライドを別スライドへ編集可能な要素として埋め込む
- 変更したモデルを `.pptx` として保存する
- モデルを JSON として出力して確認やデバッグに使う

## このリポジトリの位置づけ

これはスクリーンショット貼り付け型の処理ではありません。

描画も書き戻しも、PowerPoint の内容を解析して得たモデルを基準に行うので、JavaScript から次のような処理ができます。

- PowerPoint ファイルを開く
- スライド内容を確認、編集する
- Canvas や SVG に描画する
- 編集可能な PowerPoint ファイルを再生成する

## 公開エントリポイント

公開 API の入口は [`src/index.js`](./src/index.js) です。

```js
import { PptxThemeDocument, loadPptx } from "./src/index.js";
```

## 必要なもの

このリポジトリは現状、npm パッケージというより `src/` を中心にした source-first な構成です。別プロジェクトから直接利用する場合は、利用側で依存関係を入れてください。

必須の runtime 依存:

- `jszip`
- `fast-xml-parser`

Node.js で Canvas 描画を行う場合の任意依存:

- `@napi-rs/canvas`
- または 2D context を提供する別の Canvas 実装

例:

```bash
npm install jszip fast-xml-parser
npm install @napi-rs/canvas
```

## CDN build

ブラウザ向け bundle は次で生成できます。

```bash
npm install
npm run build:cdn
```

生成物:

- `dist/pTxTq.js`
- `dist/pTxTq.min.js`
- `ptxtq.min.js`

bundle はグローバル変数 `pTxTq` として公開されます。

Jujutsu の運用メモ: [JUJUTSU.md](./JUJUTSU.md)

### script tag から使う

```html
<script src="./dist/pTxTq.min.js"></script>
<script>
  async function run(file) {
    const doc = await pTxTq.loadPptx(await file.arrayBuffer());
    console.log(doc.slideCount);

    const canvas = document.querySelector("canvas");
    await doc.renderSlide(0, { mode: "canvas", target: canvas });

    const outBytes = await doc.toPptxBuffer();
    console.log(outBytes.byteLength);
  }
</script>
```

ブラウザ bundle の補足:

- 入力は `Blob`、`ArrayBuffer`、`Uint8Array` を想定しています
- ファイルパスを直接渡す読み込みはブラウザ bundle では使えません
- `toPptxBuffer()` はブラウザ bundle では `Uint8Array` を返します
- `saveAs()` のようなディスク書き込み API は Node.js 向けです

## Usage

スライド index は 0 始まりです。

### `.pptx` を読み込む

```js
import { PptxThemeDocument } from "./src/index.js";

const doc = await PptxThemeDocument.load("input.pptx");

console.log(doc.slideCount);
console.log(doc.metadata);

const firstSlide = doc.getSlide(0);
console.log(firstSlide?.name);
```

バイト列から読み込むこともできます。

```js
import { readFile } from "node:fs/promises";
import { loadPptx } from "./src/index.js";

const bytes = await readFile("input.pptx");
const doc = await loadPptx(bytes);
```

### `pptx -> svg`

```js
import { writeFile } from "node:fs/promises";
import { PptxThemeDocument } from "./src/index.js";

const doc = await PptxThemeDocument.load("input.pptx");
const svg = await doc.renderSlide(0, { mode: "svg" });

await writeFile("slide-1.svg", svg, "utf8");
```

### Node.js で `pptx -> canvas`

```js
import { writeFile } from "node:fs/promises";
import { createCanvas } from "@napi-rs/canvas";
import { PptxThemeDocument } from "./src/index.js";

const doc = await PptxThemeDocument.load("input.pptx");
const canvas = createCanvas(1280, 720);

await doc.renderSlide(0, {
  mode: "canvas",
  target: canvas,
  widthPx: 1280,
  heightPx: 720
});

await writeFile("slide-1.png", canvas.toBuffer("image/png"));
```

`widthPx` と `heightPx` を省略した場合は、スライドサイズから出力サイズを決めます。

### ブラウザで `pptx -> canvas`

```js
import { PptxThemeDocument } from "./src/index.js";

const doc = await PptxThemeDocument.load(arrayBufferFromUpload);
const canvas = document.querySelector("canvas");

await doc.renderSlide(0, {
  mode: "canvas",
  target: canvas
});
```

### `pptx -> pptx` の round-trip

```js
import { PptxThemeDocument } from "./src/index.js";

const doc = await PptxThemeDocument.load("input.pptx");

await doc.saveAs("roundtrip.pptx");
```

生成したバイト列を直接受け取ることもできます。

```js
const buffer = await doc.toPptxBuffer();
const bytes = await doc.toPptxBuffer({ type: "uint8array" });
```

### スライドを別スライドへ埋め込む

`embedSlideIntoSlide()` は、あるスライドを別スライドへ編集可能な要素としてコピーします。

```js
import { PptxThemeDocument } from "./src/index.js";

const doc = await PptxThemeDocument.load("input.pptx");

doc.embedSlideIntoSlide(0, 1, {
  x: 0,
  y: 0,
  cx: doc.metadata.slideSizeEmu.cx,
  cy: doc.metadata.slideSizeEmu.cy,
  includeBackground: true
});

await doc.saveAs("composed.pptx");
```

主なオプション:

- `x`, `y`, `cx`, `cy`: 埋め込み先の配置とサイズ。単位は EMU
- `includeBackground`: 元スライドの背景も含める
- `useRenderElements`: 利用可能なら render 用に整形済みの要素を使う
- `includeUnsupported`: 未対応要素も除外せず含める

### 解析済みモデルを出力する

```js
import { PptxThemeDocument } from "./src/index.js";

const doc = await PptxThemeDocument.load("input.pptx");

await doc.saveModelJson("dist/model.json", {
  includeDataUri: false,
  includePrivate: false,
  includeRaw: false
});
```

オブジェクトや文字列として直接取得することもできます。

```js
const json = doc.toJsonObject();
const text = doc.toJsonString({ space: 2 });
```

## API overview

### `PptxThemeDocument`

- `PptxThemeDocument.load(source)`
- `PptxThemeDocument.loadFile(filePath)`
- `doc.slideCount`
- `doc.metadata`
- `doc.slides`
- `doc.getSlide(index)`
- `doc.renderSlide(index, options)`
- `doc.embedSlideIntoSlide(sourceIndex, targetIndex, options)`
- `doc.toPptxPackage(options)`
- `doc.toPptxBuffer(options)`
- `doc.saveAs(filePath, options)`
- `doc.toJsonObject(options)`
- `doc.toJsonString(options)`
- `doc.saveModelJson(filePath, options)`

### `loadPptx(source)`

次のショートカットです。

```js
await PptxThemeDocument.load(source);
```

## リポジトリ構成

- `src/core/`: Open XML 読み込み、テーマ解決、PresentationML 解析、編集可能モデル
- `src/render/`: Canvas / SVG レンダラ
- `src/write/`: `.pptx` への書き戻し
- `src/utils/`: XML、単位変換、geometry、color、object などの補助処理

公開の中心となる実装は `src/` 配下です。
`scripts/`、`docs/`、サンプル deck などはローカル検証用で、主たる公開 API ではありません。
