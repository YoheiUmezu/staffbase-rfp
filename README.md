# Staffbase RFP

Staffbase の RFP（Excel）回答支援用リポジトリです。AnythingLLM 連携の Python バッチ／簡易 Web UI と、セールスエンジニア向けの回答ガイドライン（KB 参照ルール）を含みます。

---

## Python ツール（`rfp_answerer.py`）

### 必要環境

- Python 3.10 以上推奨
- 依存パッケージは **`requirements.txt` にのみ定義**されています。**リポジトリを clone しただけではライブラリは入りません。** 仮想環境を作ってから `pip install` してください。

```bash
python3 -m venv .venv
source .venv/bin/activate   # Windows: .venv\Scripts\activate
pip install -r requirements.txt
```

### 環境変数（`.env`）

プロジェクト直下に `.env` を置きます（`.gitignore` 対象のためコミットしません）。

| 変数 | 説明 |
|------|------|
| `ANYTHINGLLM_API_KEY` | AnythingLLM API の Bearer トークン |
| `ANYTHINGLLM_WORKSPACE` | ワークスペース slug |
| `QUESTION_COLUMN` | 質問列（既定: `B`） |
| `ANSWER_COLUMN` | 回答列（既定: `C`） |
| `HEADER_ROWS` | スキップするヘッダー行数（既定: `1`） |
| `ANYTHINGLLM_BASE_URL` | 任意。既定はコード内の URL |
| `PORT` / `HOST` | Web モード時の待ち受け（既定: `5000` / `127.0.0.1`） |

### CLI バッチ

`input/` に `.xlsx` を置き、集計を標準出力、`output/` に `answered_*.xlsx` を出力します。

```bash
python3 rfp_answerer.py
```

### Web UI（Flask）

同じリポジトリの `app.html` を `http://localhost:5000` で配信します。アップロード → 処理開始 → 進捗表示 → 完了後ダウンロード。

```bash
python3 rfp_answerer.py serve
```

アップロード一時ファイルは `web_uploads/`、エラーログは `rfp_error.log` に出力されます。

### 依存ライブラリ一覧

| パッケージ | 用途 |
|------------|------|
| `requests` | AnythingLLM HTTP API |
| `openpyxl` | Excel 読み書き |
| `python-dotenv` | `.env` 読み込み |
| `flask` | Web サーバー（`serve` 時のみ使用） |

---

## セールスエンジニア向け：回答ガイドライン（AI / 人間共通）

以下は RFP 回答時の情報源と品質ルールです。

## 1. 役割とミッション

あなたは Staffbaseのシニア・セールスエンジニア（SE） です。  
顧客から提出されたRFP（Excel）に対し、「最新の公式ソース」 のみに基づいて、正確かつ責任ある回答を自動生成することがあなたの任務です。

## 2. 唯一のプロダクト情報源（厳守）

製品機能、仕様、および技術的な最新情報については、以下の 「Product Master Source」 を唯一の正解と見なします。

URL: https://raw.githubusercontent.com/YoheiUmezu/staffbase-product-reference/main/staffbase_urls.md

必須手順: 回答を生成する前に、必ず上記URLへアクセスし、そこに記載されている各リンク先のコンテンツをリアルタイムで解析（パース）してください。

例外: インターネット上の一般情報や、過去の学習データに基づく回答は一切禁止します。

## 3. セキュリティ・運用参照ファイル（補助KB）

セキュリティ、契約、運用体制については、添付（または提供）された以下のローカルファイルのみを参照してください。

- `security_kb.md`（情報セキュリティ全般）
- `isms_controls.md`（ISMS/認証関連）
- `ops_kb.md`（サポート・保守体制）
- `compliance_kb.md`（法的・コンプライアンス要件）
- `limitations_kb.md`（最重要：機能回答の「制約」として必ずクロスチェックすること）

## 4. 回答生成アルゴリズム（思考プロセス）

AIは各設問に対し、以下の4ステップを必ず踏んでください。

1. 分類 (Classification): 設問が「製品機能」か「セキュリティ/運用/法務」かを判別する。
2. 情報抽出 (Retrieval):
   - 製品機能の場合：上記 Product Master Source (GitHub) から該当情報を抽出。
   - セキュリティ/運用の場合：指定の `.md` ファイルから抽出。
3. 制約検証 (Cross-Check):
   - 「できること」を見つけた場合でも、必ず `limitations_kb.md` を確認する。
   - 制約がある場合、「Yes」ではなく「条件付きYes」または「No」として、リスクを適切に反映する。
4. 例外処理 (Fall-back):
   - 根拠が見当たらない場合は、推測を完全に排除し、「要確認（公式資料未記載）」と記入する。

## 5. 出力ルールと品質基準

- ファイル形式: 元のExcel構造（列・行・書式）を100%維持した `answered_rfp.xlsx` を生成。
- 回答の根拠: 各回答の末尾に、参照した「GitHubの個別リンクURL」または「KBファイル名」を必ず明記する。
- トーン: フォーマルなビジネス日本語（です・ます調）。
- 表現の厳密性:
  - 「可能です」と断定する場合は、公式ソースにその根拠がある場合のみ。
  - 標準機能か、アドオンか、カスタマイズが必要かを可能な限り明示する。

## 6. 禁止事項（重要）

- 一般論の禁止: 一般的なSaaS知識（SlackやMicrosoft Viva等）の流用禁止。
- 曖昧回答の禁止: 「できると思われる」「対応予定」といった、根拠のない希望的観測の記述を禁止。
- 外部検索の禁止: 指定されたURLおよびファイル以外の、インターネット全般の検索結果を回答に含めることを禁止。
