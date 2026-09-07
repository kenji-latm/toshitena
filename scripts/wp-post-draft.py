#!/usr/bin/env python3
"""WordPress REST API に記事を「下書き」として投稿する。

必要な環境変数:
  WP_SITE_URL      例: https://example.com  （/wp-json は付けない）
  WP_USER          WordPress のユーザー名（ログインIDまたはメール）
  WP_APP_PASSWORD  アプリケーションパスワード（ユーザー → プロフィール で発行）

使い方:
  python3 scripts/wp-post-draft.py blog/wp-draft-chihou-startup.html
  python3 scripts/wp-post-draft.py blog/wp-draft-chihou-startup.html --update 123
  python3 scripts/wp-post-draft.py blog/wp-draft-chihou-startup.html --tags   # タグも付与
"""
import argparse
import base64
import json
import os
import sys
import urllib.error
import urllib.parse
import urllib.request

TITLE = "地方スタートアップの登記は、もう「近さ」で選ばなくていい"
SLUG = "chihou-startup-shogyo-toki"
EXCERPT = (
    "「地元の司法書士にJ-KISSの話をしたら初めてだと言われた」。地方でよく聞く話です。"
    "これは力量ではなく案件分布の問題で、だからこそ距離を越えて頼めば解決します。"
    "商業登記が完全オンラインで終わる実際の流れと、ラウンドでよく詰まる5つのポイントをまとめました。"
)
TAGS = [
    "スタートアップ", "商業登記", "J-KISS", "種類株式", "新株予約権",
    "資金調達", "オンライン申請", "電子署名", "資本政策", "地方創業",
]


def env(name):
    value = os.environ.get(name, "").strip()
    if not value:
        sys.exit(f"環境変数 {name} が設定されていません。")
    return value


def request(method, url, auth, payload=None):
    data = json.dumps(payload).encode("utf-8") if payload is not None else None
    req = urllib.request.Request(url, data=data, method=method)
    req.add_header("Authorization", f"Basic {auth}")
    req.add_header("Content-Type", "application/json; charset=utf-8")
    req.add_header("User-Agent", "wp-post-draft/1.0")
    try:
        with urllib.request.urlopen(req, timeout=60) as res:
            return json.loads(res.read().decode("utf-8"))
    except urllib.error.HTTPError as err:
        body = err.read().decode("utf-8", "replace")[:800]
        sys.exit(f"{method} {url}\nHTTP {err.code} {err.reason}\n{body}")
    except urllib.error.URLError as err:
        sys.exit(f"{method} {url}\n接続できませんでした: {err.reason}")


def resolve_tags(api, auth, names):
    ids = []
    for name in names:
        query = urllib.parse.urlencode({"search": name, "per_page": 100})
        found = request("GET", f"{api}/tags?{query}", auth)
        match = next((t for t in found if t.get("name") == name), None)
        if match is None:
            match = request("POST", f"{api}/tags", auth, {"name": name})
            print(f"  タグ作成: {name} (id={match['id']})")
        ids.append(match["id"])
    return ids


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("html", help="投稿本文のHTMLファイル（Gutenbergブロック形式）")
    parser.add_argument("--update", type=int, metavar="POST_ID",
                        help="既存の下書きを上書き更新する")
    parser.add_argument("--tags", action="store_true",
                        help="タグを解決・作成して付与する")
    args = parser.parse_args()

    site = env("WP_SITE_URL").rstrip("/")
    auth = base64.b64encode(
        f'{env("WP_USER")}:{env("WP_APP_PASSWORD")}'.encode("utf-8")
    ).decode("ascii")
    api = f"{site}/wp-json/wp/v2"

    with open(args.html, encoding="utf-8") as f:
        content = f.read()

    payload = {
        "title": TITLE,
        "slug": SLUG,
        "excerpt": EXCERPT,
        "content": content,
        "status": "draft",
    }
    if args.tags:
        print("タグを解決しています…")
        payload["tags"] = resolve_tags(api, auth, TAGS)

    if args.update:
        post = request("POST", f"{api}/posts/{args.update}", auth, payload)
        print(f"下書きを更新しました（ID {post['id']}）")
    else:
        post = request("POST", f"{api}/posts", auth, payload)
        print(f"下書きを作成しました（ID {post['id']}）")

    print(f"編集画面: {site}/wp-admin/post.php?post={post['id']}&action=edit")
    print(f"プレビュー: {post.get('link', '(なし)')}")
    print(f"ステータス: {post.get('status')}")


if __name__ == "__main__":
    main()
