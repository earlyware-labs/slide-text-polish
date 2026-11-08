# slide-text-polish

AIが作ったPowerPoint（.pptx）の**日本語の文章を、プレゼン向けに後から直す** Claude スキルです。

- 言語指定を `ja-JP` に直し、PowerPoint の日本語禁則処理を効かせる
- 欧文フォントだけが指定された日本語に和文フォントを設定する
- 「・」を文字で打った箇条書きを箇条書きの書式＋ぶら下げインデントに直し、2行目以降を本文の先頭に揃える（「共通の注意 本文」のようなラベル付きの段落も、ラベルの幅でぶら下げる）
- 長すぎるタイトル・箇条書き・表のセル・グラフの項目名を検出し、Claude が短く書き直す
- **書き直しで数字・用語・否定表現が消えていないかを、修正前後で機械的に照合する**
- 文節と用語辞書をもとに、泣き別れ・禁則違反を見つけて改行位置を直す
- Windows では PowerPoint 実機（COM）の改行位置で最終確認する

## 構成

```
skills/slide-text-polish/
├── SKILL.md                # Claude への指示（ルール・手順・報告の形）
└── scripts/
    ├── common.py           # グループ・表・グラフを含む全テキストの走査
    ├── fix_lang_font.py    # 言語・和文フォント・斜体・箇条書きの一括修正
    ├── bullets.py          # 箇条書き・ラベル付きの段落のぶら下げと続きの段落の判定・修正
    ├── lint.py             # 文章量・言語・フォント・箇条書き・改行のチェック
    ├── breaks.py           # 改行してよい位置（BudouX、無ければ簡易分割）
    ├── check_loss.py       # 書き直しによる情報の消失チェック
    ├── check_ppt.ps1       # Windows: PowerPoint 実機で JPEG＋実際の行を書き出す
    └── render_mac.sh       # Mac: PowerPoint で PDF→PNG
glossary.example.txt        # 用語辞書の例
```

## 使い方

### Claude アプリ
`skills/slide-text-polish` フォルダを zip にしてスキルとして追加し、.pptx を渡して「チェックして直して」と頼みます。

### Claude Code
```bash
cp -r skills/slide-text-polish ~/.claude/skills/
# Python の実行には uv を使う（依存パッケージは --with で自動インストール）
```

### スクリプトを直接使う
```bash
cd skills/slide-text-polish/scripts
uv run --with python-pptx --with budoux python fix_lang_font.py 提案書.pptx work.pptx
uv run --with python-pptx --with budoux python lint.py work.pptx --glossary glossary.txt
# （文章を書き直して short.pptx に保存した後）
uv run --with python-pptx --with budoux python check_loss.py work.pptx short.pptx --glossary glossary.txt
```

Windows で実機の改行を使う場合:
```powershell
powershell -ExecutionPolicy Bypass -File check_ppt.ps1 -Path C:\work\short.pptx -Out C:\work\check
uv run --with python-pptx --with budoux python lint.py short.pptx --lines C:\work\check\lines.tsv
```

## 制限
- 改行の推定（実機を使わない場合）は枠の幅と文字サイズからの目安です
- BudouX が使えない環境では文字種の切れ目による簡易な文節分割になり、精度が下がります
- `render_mac.sh` は実機での動作確認がまだです
- 箇条書きやラベル付きの段落の直後にある、記号もラベルも無い段落は、項目の続きとみなしてつなぎます。補足の文など別の段落のつもりだった場合は戻してください（つないだ箇所は修正時に一覧で表示します）
- 改悪チェックは数字・用語・否定表現の照合で、文の意味の変化までは判定しません（Claude が読み直して確認します）

## 参考にした先行事例
- [jp-business-slides](https://github.com/haihan19955/jp-business-slides)（PowerPoint 実機の改行位置による泣き別れ修正、`ja-JP` と禁則処理の検証）
- [pptx-ja](https://github.com/k-ibaraki/my-agent-skills/tree/main/skills/pptx-ja)（日本語の言語設定・フォント・スペルミス印の対策）
- [consulting-pptx-skill](https://github.com/carnot-tech/jinba-consulting-pptx-skill)（規約と機械チェックの組み合わせ）

## ライセンス

[MIT License](LICENSE) © 2026 Earlyware Labs

実行時に uv で取得する依存パッケージ（同梱はしていません）:

| パッケージ | ライセンス |
|---|---|
| [python-pptx](https://github.com/scanny/python-pptx) | MIT |
| [BudouX](https://github.com/google/budoux) | Apache-2.0 |
| [lxml](https://lxml.de/)（python-pptx の依存） | BSD-3-Clause |
