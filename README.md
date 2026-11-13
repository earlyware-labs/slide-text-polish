# slide-text-polish

AIが作ったPowerPoint（.pptx）の**日本語の文章を、プレゼン向けに後から直す** AIエージェント用のスキルです。
[Agent Skills](https://agentskills.io) の形式（`SKILL.md`＋スクリプト）で作っており、**Claude Code・Codex・GitHub Copilot CLI** で動作します。

- 言語指定を `ja-JP` に直し、PowerPoint の日本語禁則処理を効かせる
- 欧文フォントだけが指定された日本語に和文フォントを設定する
- 「・」を文字で打った箇条書きを箇条書きの書式＋ぶら下げインデントに直し、2行目以降を本文の先頭に揃える（「共通の注意 本文」のようなラベル付きの段落も、ラベルの幅でぶら下げる）
- 長すぎるタイトル・箇条書き・表のセル・グラフの項目名を検出し、AIエージェントが短く書き直す
- **書き直しで数字・用語・否定表現が消えていないかを、修正前後で機械的に照合する**
- 文節と用語辞書をもとに、泣き別れ・禁則違反を見つけて改行位置を直す
- PowerPoint 実機（COM）の画像と改行位置で最終確認する

## 構成

```
skills/slide-text-polish/
├── SKILL.md                # AIエージェントへの指示（ルール・手順・報告の形）
└── scripts/
    ├── common.py           # グループ・表・グラフを含む全テキストの走査
    ├── fix_lang_font.py    # 言語・和文フォント・斜体・箇条書きの一括修正
    ├── bullets.py          # 箇条書き・ラベル付きの段落のぶら下げと続きの段落の判定・修正
    ├── lint.py             # 文章量・言語・フォント・箇条書き・改行のチェック
    ├── breaks.py           # 改行してよい位置（BudouX、無ければ簡易分割）
    ├── check_loss.py       # 書き直しによる情報の消失チェック
    └── check_ppt.ps1       # PowerPoint 実機で JPEG＋実際の行を書き出す
glossary.example.txt        # 用語辞書の例
```

## 動作環境
- Windows
- PowerPoint（デスクトップ版）：最終確認（SKILL.md の手順6）で実機の表示を使います
- [uv](https://docs.astral.sh/uv/)：Python の実行に使います
- 次のいずれかのAIエージェント：Claude Code、Codex、GitHub Copilot CLI

Mac と、PowerPoint の無い環境（Claude アプリのクラウド上の作業環境など）には対応していません。

## インストール

`skills/slide-text-polish` フォルダを、使うAIエージェントのスキル用フォルダにコピーします。Python の依存パッケージは、実行時に uv が自動で用意します。

| AIエージェント | コピー先（ユーザー単位） |
|---|---|
| Claude Code | `~/.claude/skills/` |
| Codex | `~/.agents/skills/` |
| GitHub Copilot CLI | `~/.copilot/skills/` または `~/.agents/skills/` |

Codex と GitHub Copilot CLI はどちらも `~/.agents/skills/` を読むので、両方で使う場合はここに1つ置けば足ります。

```powershell
# Claude Code
New-Item -ItemType Directory -Force $HOME\.claude\skills | Out-Null
Copy-Item -Recurse skills\slide-text-polish $HOME\.claude\skills

# Codex・GitHub Copilot CLI（共通）
New-Item -ItemType Directory -Force $HOME\.agents\skills | Out-Null
Copy-Item -Recurse skills\slide-text-polish $HOME\.agents\skills
```

コピーした後、AIエージェントを起動し直します（GitHub Copilot CLI は起動中なら `/skills reload` でも読み込めます）。

### 読み込まれたかを確かめる
| AIエージェント | 確かめ方 |
|---|---|
| Claude Code | `/skills` を実行し、一覧に `slide-text-polish` があるか見る |
| GitHub Copilot CLI | `/skills list` を実行する（「What skills do you have?」と聞いてもよい） |
| Codex | 「使えるスキルを一覧にして」と頼み、`slide-text-polish` が含まれるか見る |

## プロンプトから呼び出す

スクリプトを自分で実行する必要はありません。AIエージェントに日本語で頼むだけで、スキルの手順（言語・フォント・箇条書きの修正 → チェック → 文章の短縮 → 改悪チェック → 改行の修正 → PowerPoint 実機での確認 → 報告）がすべて進みます。

### 呼び出し方は2通り

**1. 普通に頼む（自動で選ばれる）**

.pptx のパスと、何をしてほしいかを書きます。AIエージェントはスキルの説明（`SKILL.md` の `description`）を見て、このスキルを使うかどうかを自分で判断します。

```text
C:\work\提案書.pptx の日本語をチェックして、読みやすく直して
```

「チェックして」「読みやすくして」「レビューして」など、.pptx の文章を見直す依頼であれば選ばれます。AIエージェントにスライドを作らせた直後に「仕上げて」と頼んでも使われます。

**2. スキル名を指定して頼む（確実に使わせる）**

自動で選ばれなかったときや、確実にこのスキルを使わせたいときは、スキル名を書きます。

| AIエージェント | 書き方の例 |
|---|---|
| Claude Code | `/slide-text-polish C:\work\提案書.pptx をチェックして直して`（先頭に `/スキル名`） |
| GitHub Copilot CLI | `/slide-text-polish スキルを使って、C:\work\提案書.pptx をチェックして直して`（文中に `/スキル名`） |
| Codex | `slide-text-polish スキルを使って、C:\work\提案書.pptx をチェックして直して`（スキル名を文中に書く） |

### プロンプトに書いておくとよいこと

| 伝えたいこと | 書き方の例 |
|---|---|
| 改行で分けたり、短縮で消したり言い換えたりしてほしくない語（固有名詞・商品名・業界用語） | 「『品番マスタ』『自動照合』は用語なので、分けたり言い換えたりしないで」 |
| 用語辞書のファイルがある | 「用語辞書は C:\work\glossary.txt を使って」（1行1語。[glossary.example.txt](glossary.example.txt) を参照） |
| 和文フォントの指定（既定は游ゴシック） | 「和文フォントはメイリオにして」 |
| 保存先・ファイル名（既定は `元の名前_修正版.pptx`） | 「C:\work\out\提案書_最終.pptx に保存して」 |

例：

```text
/slide-text-polish C:\work\提案書.pptx をチェックして直して。
「品番マスタ」「自動照合」は社内の用語なので、分けたり言い換えたりしないで。
```

### 実行中に起きること
- **コマンドの実行の許可を求められます。** スキルは uv（Python のスクリプト）と PowerShell（`check_ppt.ps1`）を実行します。AIエージェントの設定によっては、そのたびに許可を求められます
- **初回だけネットワークが必要です。** uv が python-pptx などのパッケージをダウンロードします。2回目以降はキャッシュだけで動き、通信しません（[外部との通信](#外部との通信)）
- **PowerPoint が画面に出ないまま起動します。** 最終確認で、スライドの画像と実際に表示された行を書き出すためです。ファイルは読み取り専用で開くので、元のファイルが書き換わることはありません
- 元のファイルは上書きされず、修正版は別名で保存されます

### 返ってくるもの
- 修正済みのファイル（既定は `元の名前_修正版.pptx`）
- 報告
  - 最初の1行：PowerPoint 実機で確認したか、見たスライドの枚数、指摘件数の修正前→修正後
  - 言語・フォントを直した件数
  - 文章の変更一覧（スライド・場所・修正前・修正後）
  - **要確認**：数字・用語・否定表現が消えた可能性のある箇所、意味が変わりうる修正
  - 手動で改行を入れた箇所、用語辞書に追加する候補

要確認の箇所は、必ず自分の目で確かめてください。修正が気に入らない箇所は、「スライド3の2行目は元に戻して」のように続けて頼めば直します。

## 外部との通信
- スクリプトと依存パッケージ（python-pptx・BudouX・lxml）は外部と通信しません。BudouX の学習済みデータもパッケージに同梱されています
- uv は初回にパッケージをダウンロードし、その後も実行のたびに PyPI へ新しい版を問い合わせます（スライドの内容は送りません）。`uv run --offline ...` とするとキャッシュだけで動き、通信しません。スキル（SKILL.md）のコマンドには `--offline` を付けています
- 手順6で使う PowerPoint は、アプリ自体が通信することがあります（Office の設定で管理してください）

## 制限
- 改行の推定（実機を使わない場合）は枠の幅と文字サイズからの目安です
- BudouX が使えない環境では文字種の切れ目による簡易な文節分割になり、精度が下がります
- 箇条書きやラベル付きの段落の直後にある、記号もラベルも無い段落は、項目の続きとみなしてつなぎます。補足の文など別の段落のつもりだった場合は戻してください（つないだ箇所は修正時に一覧で表示します）
- 改悪チェックは数字・用語・否定表現の照合で、文の意味の変化までは判定しません（AIエージェントが読み直して確認します）

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
