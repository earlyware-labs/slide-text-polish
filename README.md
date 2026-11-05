# slide-text-polish

AIが作ったPowerPoint（.pptx）の**日本語の文章を、プレゼン向けに後から直す** Claude スキルです。

- 言語指定を `ja-JP` に直し、PowerPoint の日本語禁則処理を効かせる
- 欧文フォントだけが指定された日本語に和文フォントを設定する
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
    ├── fix_lang_font.py    # 言語・和文フォント・斜体の一括修正
    ├── lint.py             # 文章量・文字サイズ・言語・フォント・改行のチェック
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
pip install python-pptx budoux
```

### スクリプトを直接使う
```bash
cd skills/slide-text-polish/scripts
python fix_lang_font.py 提案書.pptx work.pptx
python lint.py work.pptx --glossary glossary.txt
# （文章を書き直して short.pptx に保存した後）
python check_loss.py work.pptx short.pptx --glossary glossary.txt
```

Windows で実機の改行を使う場合:
```powershell
powershell -ExecutionPolicy Bypass -File check_ppt.ps1 -Path C:\work\short.pptx -Out C:\work\check
python lint.py short.pptx --lines C:\work\check\lines.tsv
```

## 制限
- 改行の推定（実機を使わない場合）は枠の幅と文字サイズからの目安です
- BudouX が無い環境では文字種の切れ目による簡易な文節分割になり、精度が下がります
- `check_ppt.ps1` と `render_mac.sh` は実機での動作確認がまだです
- 改悪チェックは数字・用語・否定表現の照合で、文の意味の変化までは判定しません（Claude が読み直して確認します）

## 参考にした先行事例
- [jp-business-slides](https://github.com/haihan19955/jp-business-slides)（PowerPoint 実機の改行位置による泣き別れ修正、`ja-JP` と禁則処理の検証）
- [pptx-ja](https://github.com/k-ibaraki/my-agent-skills/tree/main/skills/pptx-ja)（日本語の言語設定・フォント・スペルミス印の対策）
- [consulting-pptx-skill](https://github.com/carnot-tech/jinba-consulting-pptx-skill)（規約と機械チェックの組み合わせ）
