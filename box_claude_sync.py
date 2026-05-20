#!/usr/bin/env python3
"""
Box-Claude自動同期スクリプト

使い方:
  1. .envファイルに認証情報を設定
  2. python box_claude_sync.py
  3. Boxの「Claude/prompts」フォルダに .txt/.md ファイルを保存
  4. 自動的にClaude APIへ送信し、結果が「Claude/results」に保存される
"""

import os
import sys
import time
import json
from datetime import datetime
from io import BytesIO
from pathlib import Path

try:
    import anthropic
except ImportError:
    sys.exit("anthropicパッケージが必要です: pip install anthropic")

try:
    import boxsdk
    from boxsdk import OAuth2, Client
except ImportError:
    sys.exit("boxsdkパッケージが必要です: pip install boxsdk")

try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    pass  # python-dotenv なしでも動作する

# ---- 設定 ----
POLL_INTERVAL_SEC = int(os.getenv("POLL_INTERVAL", "60"))
PROMPT_FOLDER_PATH = os.getenv("BOX_PROMPT_FOLDER", "Claude/prompts")
RESULT_FOLDER_PATH = os.getenv("BOX_RESULT_FOLDER", "Claude/results")
PROCESSED_FOLDER_PATH = os.getenv("BOX_PROCESSED_FOLDER", "Claude/processed")
CLAUDE_MODEL = os.getenv("CLAUDE_MODEL", "claude-sonnet-4-6")
CLAUDE_MAX_TOKENS = int(os.getenv("CLAUDE_MAX_TOKENS", "8192"))


def _load_token_cache(cache_path: str) -> dict:
    try:
        with open(cache_path) as f:
            return json.load(f)
    except (FileNotFoundError, json.JSONDecodeError):
        return {}


def _save_token_cache(cache_path: str, data: dict):
    with open(cache_path, "w") as f:
        json.dump(data, f)


def setup_box_client() -> Client:
    """
    Box クライアントを構築する。
    環境変数の組み合わせに応じて3つの認証方式をサポート:

    【方式A: Developerトークン（一時的な動作確認用）】
      BOX_ACCESS_TOKEN のみ設定

    【方式B: OAuth2（個人利用・長期運用向け）】
      BOX_CLIENT_ID, BOX_CLIENT_SECRET, BOX_ACCESS_TOKEN,
      BOX_REFRESH_TOKEN を設定

    【方式C: CCG / JWT（エンタープライズ向け）】
      Box SDK の設定ファイルを使う場合は自前で拡張してください
    """
    access_token = os.getenv("BOX_ACCESS_TOKEN", "")
    refresh_token = os.getenv("BOX_REFRESH_TOKEN", "")
    client_id = os.getenv("BOX_CLIENT_ID", "")
    client_secret = os.getenv("BOX_CLIENT_SECRET", "")
    token_cache = os.getenv("BOX_TOKEN_CACHE", ".box_tokens.json")

    if client_id and client_secret and refresh_token:
        # 方式B: OAuth2 with refresh
        cache = _load_token_cache(token_cache)
        if cache:
            access_token = cache.get("access_token", access_token)
            refresh_token = cache.get("refresh_token", refresh_token)

        def _store_tokens(new_access, new_refresh):
            _save_token_cache(token_cache, {
                "access_token": new_access,
                "refresh_token": new_refresh,
            })

        auth = OAuth2(
            client_id=client_id,
            client_secret=client_secret,
            access_token=access_token,
            refresh_token=refresh_token,
            store_tokens=_store_tokens,
        )
    elif access_token:
        # 方式A: Developerトークン（60分で失効）
        auth = OAuth2(
            client_id="",
            client_secret="",
            access_token=access_token,
        )
    else:
        sys.exit(
            "Box認証情報が設定されていません。\n"
            ".envファイルに BOX_ACCESS_TOKEN を設定してください。\n"
            "詳細は .env.example を参照。"
        )

    return Client(auth)


def _get_or_create_folder(client: Client, path: str):
    """スラッシュ区切りのパスでフォルダを取得 / 作成する"""
    parts = [p for p in path.strip("/").split("/") if p]
    current = client.folder("0")  # ルートフォルダ

    for part in parts:
        found = None
        offset = 0
        while True:
            items = current.get_items(limit=100, offset=offset)
            batch = list(items)
            if not batch:
                break
            for item in batch:
                if item.type == "folder" and item.name == part:
                    found = item
                    break
            if found:
                break
            offset += 100

        if found:
            current = client.folder(found.id)
        else:
            current = current.create_subfolder(part)
            print(f"フォルダ作成: {part}")

    return current


def call_claude(prompt: str) -> str:
    api_key = os.getenv("ANTHROPIC_API_KEY")
    if not api_key:
        raise EnvironmentError(
            "ANTHROPIC_API_KEY が設定されていません。.envファイルを確認してください。"
        )
    client = anthropic.Anthropic(api_key=api_key)
    message = client.messages.create(
        model=CLAUDE_MODEL,
        max_tokens=CLAUDE_MAX_TOKENS,
        messages=[{"role": "user", "content": prompt}],
    )
    return message.content[0].text


def process_once(box: Client):
    """プロンプトフォルダを1回スキャンして処理する"""
    prompt_folder = _get_or_create_folder(box, PROMPT_FOLDER_PATH)
    result_folder = _get_or_create_folder(box, RESULT_FOLDER_PATH)
    processed_folder = _get_or_create_folder(box, PROCESSED_FOLDER_PATH)

    items = list(prompt_folder.get_items(limit=100))
    prompt_files = [
        i for i in items
        if i.type == "file" and Path(i.name).suffix in (".txt", ".md")
    ]

    if not prompt_files:
        return

    for item in prompt_files:
        print(f"\n[{datetime.now().strftime('%H:%M:%S')}] 処理開始: {item.name}")

        # プロンプト取得
        raw = box.file(item.id).content()
        prompt = raw.decode("utf-8", errors="replace").strip()
        if not prompt:
            print(f"  スキップ（空ファイル）: {item.name}")
            continue

        # Claude API 呼び出し
        print(f"  Claude ({CLAUDE_MODEL}) に送信中...")
        result = call_claude(prompt)
        print(f"  回答取得完了 ({len(result)}文字)")

        # 結果を Box に保存
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        stem = Path(item.name).stem
        result_filename = f"result_{stem}_{timestamp}.md"

        header = (
            f"# {stem}\n\n"
            f"**生成日時**: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}  \n"
            f"**モデル**: {CLAUDE_MODEL}  \n\n"
            f"---\n\n"
            f"## プロンプト\n\n{prompt}\n\n"
            f"---\n\n"
            f"## 回答\n\n{result}\n"
        )
        result_bytes = header.encode("utf-8")
        result_folder.upload_stream(BytesIO(result_bytes), result_filename)
        print(f"  結果保存: {RESULT_FOLDER_PATH}/{result_filename}")

        # 処理済みフォルダへ移動
        box.file(item.id).move(processed_folder)
        print(f"  移動完了: {PROMPT_FOLDER_PATH}/{item.name} → {PROCESSED_FOLDER_PATH}/")


def main():
    print("=" * 50)
    print("Box-Claude 自動同期スクリプト")
    print("=" * 50)
    print(f"プロンプトフォルダ : {PROMPT_FOLDER_PATH}")
    print(f"結果フォルダ       : {RESULT_FOLDER_PATH}")
    print(f"処理済みフォルダ   : {PROCESSED_FOLDER_PATH}")
    print(f"ポーリング間隔     : {POLL_INTERVAL_SEC}秒")
    print(f"Claudeモデル       : {CLAUDE_MODEL}")
    print("=" * 50)
    print("Ctrl+C で停止\n")

    try:
        box = setup_box_client()
        print("Box認証成功\n")
    except Exception as e:
        sys.exit(f"Box認証エラー: {e}")

    while True:
        try:
            process_once(box)
        except KeyboardInterrupt:
            print("\n停止しました。")
            break
        except anthropic.AuthenticationError:
            print("エラー: ANTHROPIC_API_KEY が無効です。")
        except Exception as e:
            print(f"エラー: {e}")

        try:
            time.sleep(POLL_INTERVAL_SEC)
        except KeyboardInterrupt:
            print("\n停止しました。")
            break


if __name__ == "__main__":
    main()
