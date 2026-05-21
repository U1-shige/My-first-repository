Whisper Notes（iPhoneアプリ）で文字起こしされたテキストから、構造化された議事録を生成してください。

## 引数

`$ARGUMENTS` の内容に応じて動作を切り替えます：

- **引数なし** → テキスト対話モード（会議情報をユーザーに質問して議事録を作成）
- **Box ファイルID**（数字のみの文字列、例: `1234567890`）→ Box から文字起こしテキストを取得
- **ローカルファイルパス**（`/` または `./` で始まる文字列）→ ローカルファイルを読み込む

---

## 手順

### ステップ1：文字起こしテキストの取得

**Box ファイルIDの場合**
Box MCP ツール（`mcp__8004ad3b-98e4-41c0-9b57-a954a42597dc__get_file_content`）でテキストを取得する。

**ローカルファイルパスの場合**
`Read` ツールでファイルを読み込む。

**引数なしの場合**
ユーザーに以下を質問して議事録データを収集する：
1. 会議名・テーマ
2. 開催日時（日付・開始〜終了時刻）
3. 参加者
4. 議題と各議題の内容
5. 決定事項・保留事項
6. アクションアイテム（担当者・期限）

---

### ステップ2：議事録データの抽出

文字起こしテキストを解析し、以下のJSONデータを構築してください：

```json
{
  "meeting_name": "会議名（文脈から推定、不明なら「要確認」）",
  "datetime": "日付 開始時刻〜終了時刻（不明なら今日の日付）",
  "location": "場所またはオンラインツール名（不明なら「要確認」）",
  "participants": "参加者一覧（不明なら「要確認」）",
  "agenda": ["議題1", "議題2"],
  "discussions": [
    {
      "topic": "議題1",
      "summary": "議論の概要",
      "decisions": "決定事項",
      "pending": "保留事項（なければ空文字）"
    }
  ],
  "flow_summary": "協議全体の流れの概要（2〜5文程度）"
}
```

抽出方針：

| 項目 | 抽出方針 |
|------|---------|
| 会議名 | 冒頭の発言や文脈から推定。不明なら `要確認` |
| 日時 | ファイル名・本文中の日時表現から取得。不明なら今日の日付 |
| 参加者 | 「〜さん」「〜です」等の発言者名から収集 |
| 議題 | 「次に」「続いて」等の話題転換で区切り分類 |
| 決定事項 | 「〜することになりました」「〜で決定」等の表現から抽出 |
| 保留事項 | 「〜は次回」「要検討」等の表現から抽出 |
| アクションアイテム | 「〜をお願いします」「〜までに」等から担当者・期限とセットで抽出 |

---

### ステップ3：.docx ファイルの生成

1. `python-docx` がインストールされているか確認する：
   ```bash
   pip show python-docx || pip install python-docx
   ```

2. JSONデータを `/tmp/minutes_data.json` に書き出す：
   ```bash
   cat > /tmp/minutes_data.json << 'JSONEOF'
   （ステップ2で構築したJSONをここに記述）
   JSONEOF
   ```

3. `create_minutes_docx.py` を実行して .docx を生成する：
   ```bash
   python3 /home/user/My-first-repository/create_minutes_docx.py \
     /tmp/minutes_data.json \
     /home/user/My-first-repository/minutes_YYYYMMDD.docx
   ```
   ※ `YYYYMMDD` は会議の開催日（例: `20260521`）

   同名ファイルが存在する場合は `minutes_YYYYMMDD_会議名.docx` を使用する。

---

### ステップ4：Box への保存（テキスト版）

Box MCP の `upload_file` はテキストファイルのみ対応のため、議事録の**テキスト版**を Box に自動保存します。

以下の形式でテキストを作成し、Box の `Claude/文字起こし` フォルダ（フォルダID: `383194890172`）にアップロードする：

```
議事録

会議名：（値）
日時：（値）
場所：（値）
参加者：（値）

議題
１　（議題1）
２　（議題2）

議事内容
１　（議題1）

概要：
（概要テキスト）

決定事項：
（決定事項テキスト）

保留事項：
（保留事項テキスト）

２　（議題2）

概要：
（概要テキスト）

決定事項：
（決定事項テキスト）

保留事項：
（保留事項テキスト）

協議の流れ（概要）
（flow_summaryの内容）
```

アップロード時のファイル名: `minutes_YYYYMMDD.txt`（同名が存在する場合は `minutes_YYYYMMDD_会議名.txt`）

`mcp__8004ad3b-98e4-41c0-9b57-a954a42597dc__upload_file` ツールを使用：
- `file_name`: `minutes_YYYYMMDD.txt`
- `file_content`: 上記テキスト
- `parent_folder_id`: `383194890172`

---

### ステップ5：.docx を Git にコミット・プッシュ

```bash
git -C /home/user/My-first-repository add minutes_YYYYMMDD.docx
git -C /home/user/My-first-repository commit -m "Add meeting minutes YYYYMMDD"
git -C /home/user/My-first-repository push
```

---

### ステップ6：完了報告

以下をユーザーに伝える：

1. **Box（テキスト版）**: `Box > Claude > 文字起こし > minutes_YYYYMMDD.txt` に保存しました
2. **Word版（.docx）**: Git リポジトリにコミット済み。GitHub から `minutes_YYYYMMDD.docx` をダウンロードできます
3. 議事録の概要（会議名・決定事項・アクションアイテムのサマリー）
