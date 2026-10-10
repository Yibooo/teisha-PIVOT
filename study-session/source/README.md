# スライドの作り直し方

| ファイル | 内容 |
|---|---|
| `prompts.js` | プロンプトの原本。スライドと `../prompts.md` の両方に使う |
| `build.js` | スライド生成（pptxgenjs） |
| `check.py` | 重なり・はみ出しチェック（PDF の文字位置と、build.js が記録した配置を照合） |

```bash
cd study-session/source
npm install pptxgenjs@3
node build.js                                   # study.pptx と study.layout.json ができる
soffice --headless --convert-to pdf study.pptx  # LibreOffice で PDF に
python3 check.py study.layout.json study.pdf    # 「指摘 0件」になることを確認
```

`BREAK=1 OUT=broken.pptx node build.js` で、わざと崩したスライドを作れます。チェッカーが崩れを検出できるかの確認に使います。
