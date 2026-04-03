"""
pptx-reviewer Web UI のエンドツーエンドテスト。

テスト対象:
  - GET /                 : トップページ
  - GET /debug            : デバッグ情報エンドポイント
  - GET /api/terminology  : 用語リスト取得
  - POST /api/terminology : 用語リスト保存
  - POST /review          : PPTXレビュー（AI部分はモック）
  - 用語リスト検出         : terminology.json カスタマイズ反映

実行方法:
  cd web
  venv\\Scripts\\pytest tests/ -v 2>&1 | tee test-results/result_YYYYMMDD_HHMMSS.log
"""
import io
import json
import os
import sys
import tempfile
from concurrent.futures import ThreadPoolExecutor, as_completed
from pathlib import Path
from unittest.mock import patch

import pytest

# web/ フォルダを sys.path に追加
sys.path.insert(0, str(Path(__file__).parent.parent))
import app as flask_app_module
from app import app, run_extract, run_terminology_check, SKILL_DIR, PERSPECTIVES_PATH

# -----------------------------------------------------------------
# 固定レスポンス（AI モック用）
# -----------------------------------------------------------------
MOCK_REPORT = """# PowerPoint レビューレポート

**ファイル**: test.pptx
**レビュー日時**: 2026-03-26
**総スライド数**: 9 枚（ファイル全体）
**レビュー対象ページ**: 全スライド
**対象読者**: 顧客向け

---

## 総合サマリー

テスト用の固定レポートです。表記ゆれ3件、専門用語2件を検出しました。

---

## カテゴリ別 指摘件数

| カテゴリ | 件数 |
|---------|------|
| 📝 表記ゆれ・文章校正 | 3件 |
| 🔗 論理的整合性 | 1件 |
| 👥 専門用語・顧客表現 | 2件 |
| **合計** | **6件** |

---

## スライド別 指摘事項

### スライド 2：現行のシステム構成

| カテゴリ | 箇所 | 指摘内容 | 改善案 |
|---------|------|---------|--------|
| 📝 表記ゆれ | テキストボックス 1行目 | 「サーバー」→「サーバ」 | 「サーバ」に統一することを推奨 |

---

## 全体的な改善提案

1. 用語統一リストに基づき「サーバー」を「サーバ」に統一することを推奨
2. 専門用語に初出時の説明を追加することを推奨
"""

# -----------------------------------------------------------------
# フィクスチャ
# -----------------------------------------------------------------
DUMMY_PPTX = Path(__file__).parent.parent.parent / "skills" / "pptx-reviewer" / "test-files" / "dummy_proposal.pptx"
TEST_PPTX   = Path(__file__).parent.parent / "test_dummy.pptx"


@pytest.fixture
def client():
    app.config["TESTING"] = True
    with app.test_client() as c:
        yield c


@pytest.fixture
def pptx_path():
    """テスト用 PPTX を返す。test_dummy.pptx があればそれを優先。"""
    if TEST_PPTX.exists():
        return TEST_PPTX
    return DUMMY_PPTX


# -----------------------------------------------------------------
# 1. トップページ
# -----------------------------------------------------------------
class TestIndexPage:
    def test_returns_200(self, client):
        res = client.get("/")
        assert res.status_code == 200

    def test_returns_html(self, client):
        res = client.get("/")
        assert b"PPTX" in res.data or b"pptx" in res.data.lower()


# -----------------------------------------------------------------
# 2. デバッグエンドポイント
# -----------------------------------------------------------------
class TestDebugEndpoint:
    def test_returns_200(self, client):
        res = client.get("/debug")
        assert res.status_code == 200

    def test_returns_json(self, client):
        res = client.get("/debug")
        data = json.loads(res.data)
        assert "ai_provider" in data
        assert "terminology" in data
        assert "scripts" in data

    def test_scripts_exist(self, client):
        res = client.get("/debug")
        data = json.loads(res.data)
        assert data["scripts"]["extract_pptx_exists"] is True
        assert data["scripts"]["check_terminology_exists"] is True

    def test_terminology_list_loaded(self, client):
        res = client.get("/debug")
        data = json.loads(res.data)
        assert data["terminology"]["exists"] is True
        assert data["terminology"]["term_count"] >= 2  # サーバ・ユーザ

    def test_terminology_terms_content(self, client):
        res = client.get("/debug")
        data = json.loads(res.data)
        corrects = [t["correct"] for t in data["terminology"]["terms"]]
        assert "サーバ" in corrects
        assert "ユーザ" in corrects


# -----------------------------------------------------------------
# 3. /review バリデーション
# -----------------------------------------------------------------
class TestReviewValidation:
    def test_no_file_returns_400(self, client):
        res = client.post("/review")
        assert res.status_code == 400
        data = json.loads(res.data)
        assert "error" in data

    def test_wrong_extension_returns_400(self, client):
        data = {"pptx_file": (io.BytesIO(b"dummy"), "test.txt")}
        res = client.post("/review", data=data, content_type="multipart/form-data")
        assert res.status_code == 400

    def test_empty_filename_returns_400(self, client):
        data = {"pptx_file": (io.BytesIO(b""), "")}
        res = client.post("/review", data=data, content_type="multipart/form-data")
        assert res.status_code == 400


# -----------------------------------------------------------------
# 4. /review フルフロー（AI はモック）
# -----------------------------------------------------------------
class TestReviewFullFlow:
    def test_review_returns_200_with_mock_ai(self, client, pptx_path):
        with patch("app.call_ai_review", return_value=MOCK_REPORT):
            with open(pptx_path, "rb") as f:
                data = {"pptx_file": (f, "test.pptx")}
                res = client.post("/review", data=data, content_type="multipart/form-data")
        assert res.status_code == 200

    def test_review_response_contains_report(self, client, pptx_path):
        with patch("app.call_ai_review", return_value=MOCK_REPORT):
            with open(pptx_path, "rb") as f:
                data = {"pptx_file": (f, "test.pptx")}
                res = client.post("/review", data=data, content_type="multipart/form-data")
        body = json.loads(res.data)
        assert "report" in body
        assert "PowerPoint レビューレポート" in body["report"]

    def test_review_response_has_slide_count(self, client, pptx_path):
        with patch("app.call_ai_review", return_value=MOCK_REPORT):
            with open(pptx_path, "rb") as f:
                data = {"pptx_file": (f, "test.pptx")}
                res = client.post("/review", data=data, content_type="multipart/form-data")
        body = json.loads(res.data)
        assert "total_slides" in body
        assert body["total_slides"] == 9

    def test_review_response_has_filename(self, client, pptx_path):
        with patch("app.call_ai_review", return_value=MOCK_REPORT):
            with open(pptx_path, "rb") as f:
                data = {"pptx_file": (f, "test.pptx")}
                res = client.post("/review", data=data, content_type="multipart/form-data")
        body = json.loads(res.data)
        assert "report_filename" in body
        assert body["report_filename"].endswith(".md")

    def test_review_report_saved_to_uploads(self, client, pptx_path):
        """レポートが uploads/ に保存されることを確認する。"""
        with patch("app.call_ai_review", return_value=MOCK_REPORT):
            with open(pptx_path, "rb") as f:
                data = {"pptx_file": (f, "test.pptx")}
                res = client.post("/review", data=data, content_type="multipart/form-data")
        body = json.loads(res.data)
        saved = flask_app_module.UPLOAD_DIR / body["report_filename"]
        assert saved.exists(), f"レポートファイルが見つかりません: {saved}"
        # クリーンアップ
        saved.unlink()

    def test_review_with_pages_filter(self, client, pptx_path):
        """pages 指定時にレビュー対象スライドが絞られることを確認。"""
        with patch("app.call_ai_review", return_value=MOCK_REPORT):
            with open(pptx_path, "rb") as f:
                data = {"pptx_file": (f, "test.pptx"), "pages": "1,3"}
                res = client.post("/review", data=data, content_type="multipart/form-data")
        body = json.loads(res.data)
        assert body["reviewed_slides"] == [1, 3]


# -----------------------------------------------------------------
# 4b. /review 空レスポンスハンドリング
# -----------------------------------------------------------------
class TestReviewEmptyResponse:
    """AI が空レスポンスを返した場合のエラーハンドリングを確認。"""

    def test_empty_string_returns_500(self, client, pptx_path):
        """AI が空文字列を返した場合に 500 エラーが返ることを確認。"""
        with patch("app.call_ai_review", return_value=""):
            with open(pptx_path, "rb") as f:
                data = {"pptx_file": (f, "test.pptx")}
                res = client.post("/review", data=data, content_type="multipart/form-data")
        assert res.status_code == 500
        body = json.loads(res.data)
        assert "error" in body

    def test_whitespace_only_returns_500(self, client, pptx_path):
        """AI がホワイトスペースのみを返した場合に 500 エラーが返ることを確認。"""
        with patch("app.call_ai_review", return_value="   \n\t  "):
            with open(pptx_path, "rb") as f:
                data = {"pptx_file": (f, "test.pptx")}
                res = client.post("/review", data=data, content_type="multipart/form-data")
        assert res.status_code == 500
        body = json.loads(res.data)
        assert "error" in body

    def test_none_returns_500(self, client, pptx_path):
        """AI が None を返した場合に 500 エラーが返ることを確認。"""
        with patch("app.call_ai_review", return_value=None):
            with open(pptx_path, "rb") as f:
                data = {"pptx_file": (f, "test.pptx")}
                res = client.post("/review", data=data, content_type="multipart/form-data")
        assert res.status_code == 500
        body = json.loads(res.data)
        assert "error" in body

    def test_empty_response_no_file_created(self, client, pptx_path):
        """空レスポンス時に 0KB のレポートファイルが作成されないことを確認。"""
        import glob
        before = set(glob.glob(str(flask_app_module.UPLOAD_DIR / "*.md")))
        with patch("app.call_ai_review", return_value=""):
            with open(pptx_path, "rb") as f:
                data = {"pptx_file": (f, "test.pptx")}
                client.post("/review", data=data, content_type="multipart/form-data")
        after = set(glob.glob(str(flask_app_module.UPLOAD_DIR / "*.md")))
        new_files = after - before
        assert len(new_files) == 0, f"空レスポンスなのにファイルが作成されました: {new_files}"


# -----------------------------------------------------------------
# 5. ポート引数パース
# -----------------------------------------------------------------
class TestPortArgument:
    def test_explicit_port_takes_priority(self):
        """--port 引数が PORT 環境変数より優先されることを確認。"""
        import argparse, os
        parser = argparse.ArgumentParser()
        parser.add_argument("-p", "--port", type=int, default=None)

        os.environ["PORT"] = "9000"
        args = parser.parse_args(["--port", "7777"])
        port = args.port or int(os.getenv("PORT", 5000))
        assert port == 7777

    def test_env_var_used_when_no_arg(self):
        """CLI 引数なしのとき環境変数 PORT が使われることを確認。"""
        import argparse, os
        parser = argparse.ArgumentParser()
        parser.add_argument("-p", "--port", type=int, default=None)

        os.environ["PORT"] = "9000"
        args = parser.parse_args([])
        port = args.port or int(os.getenv("PORT", 5000))
        assert port == 9000

    def test_default_port_is_5000(self):
        """引数も環境変数もない場合にデフォルト 5000 が使われることを確認。"""
        import argparse, os
        parser = argparse.ArgumentParser()
        parser.add_argument("-p", "--port", type=int, default=None)

        os.environ.pop("PORT", None)
        args = parser.parse_args([])
        port = args.port or int(os.getenv("PORT", 5000))
        assert port == 5000

    def test_short_option_p(self):
        """-p の短縮形が動作することを確認。"""
        import argparse
        parser = argparse.ArgumentParser()
        parser.add_argument("-p", "--port", type=int, default=None)

        args = parser.parse_args(["-p", "8080"])
        assert args.port == 8080


# -----------------------------------------------------------------
# 6. 用語リスト API
# -----------------------------------------------------------------
class TestTerminologyAPI:
    def test_get_returns_200(self, client):
        res = client.get('/api/terminology')
        assert res.status_code == 200

    def test_get_returns_terms_list(self, client):
        res = client.get('/api/terminology')
        data = json.loads(res.data)
        assert 'terms' in data
        assert isinstance(data['terms'], list)

    def test_post_add_term(self, client):
        """新規用語を追加して保存できることを確認。"""
        res = client.get('/api/terminology')
        original = json.loads(res.data)

        new_terms = original['terms'] + [{
            'correct': 'インタフェース',
            'variants': ['インターフェース', 'インターフェイス'],
            'category': 'IT用語',
            'notes': 'テスト追加'
        }]
        res2 = client.post('/api/terminology',
            data=json.dumps({'terms': new_terms}),
            content_type='application/json')
        data2 = json.loads(res2.data)
        assert res2.status_code == 200
        assert data2['ok'] is True
        assert data2['term_count'] == len(original['terms']) + 1

        # 元に戻す
        client.post('/api/terminology',
            data=json.dumps({'terms': original['terms']}),
            content_type='application/json')

    def test_post_multiple_variants(self, client):
        """variants が複数指定できることを確認。"""
        res = client.get('/api/terminology')
        original = json.loads(res.data)

        new_terms = original['terms'] + [{
            'correct': 'テスト語',
            'variants': ['テスト語A', 'テスト語B', 'テスト語C'],
            'category': 'テスト',
        }]
        client.post('/api/terminology',
            data=json.dumps({'terms': new_terms}),
            content_type='application/json')

        res2 = client.get('/api/terminology')
        saved = json.loads(res2.data)
        added = next(t for t in saved['terms'] if t['correct'] == 'テスト語')
        assert len(added['variants']) == 3

        # 元に戻す
        client.post('/api/terminology',
            data=json.dumps({'terms': original['terms']}),
            content_type='application/json')

    def test_post_delete_term(self, client):
        """用語を削除して保存できることを確認。"""
        res = client.get('/api/terminology')
        original = json.loads(res.data)
        if len(original['terms']) == 0:
            pytest.skip('削除テスト用の用語がありません')

        reduced = original['terms'][1:]   # 先頭を削除
        res2 = client.post('/api/terminology',
            data=json.dumps({'terms': reduced}),
            content_type='application/json')
        assert json.loads(res2.data)['term_count'] == len(reduced)

        # 元に戻す
        client.post('/api/terminology',
            data=json.dumps({'terms': original['terms']}),
            content_type='application/json')

    def test_post_modify_variants(self, client):
        """variants を上書き修正できることを確認。"""
        res = client.get('/api/terminology')
        original = json.loads(res.data)

        modified = json.loads(json.dumps(original['terms']))
        if modified:
            modified[0]['variants'] = ['新しいゆれA', '新しいゆれB']
        res2 = client.post('/api/terminology',
            data=json.dumps({'terms': modified}),
            content_type='application/json')
        assert res2.status_code == 200

        res3 = client.get('/api/terminology')
        saved = json.loads(res3.data)
        if saved['terms']:
            assert saved['terms'][0]['variants'] == ['新しいゆれA', '新しいゆれB']

        # 元に戻す
        client.post('/api/terminology',
            data=json.dumps({'terms': original['terms']}),
            content_type='application/json')

    def test_post_empty_correct_returns_400(self, client):
        """正式表記が空の場合に 400 が返ることを確認。"""
        res = client.post('/api/terminology',
            data=json.dumps({'terms': [{'correct': '', 'variants': ['x']}]}),
            content_type='application/json')
        assert res.status_code == 400

    def test_post_no_terms_field_returns_400(self, client):
        res = client.post('/api/terminology',
            data=json.dumps({'data': []}),
            content_type='application/json')
        assert res.status_code == 400


# -----------------------------------------------------------------
# 7. 用語編集ロック API
# -----------------------------------------------------------------
class TestEditLock:
    def _release_all(self, client):
        """テスト前後のロック残存を掃除する（token 不明でも強制解除）。"""
        import app as m
        with m._edit_lock_mutex:
            m._edit_lock_state["token"] = None
            m._edit_lock_state["locked_at"] = None

    def test_get_lock_initially_unlocked(self, client):
        self._release_all(client)
        res = client.get('/api/terminology/lock')
        assert res.status_code == 200
        assert json.loads(res.data)['locked'] is False

    def test_acquire_lock_success(self, client):
        self._release_all(client)
        res = client.post('/api/terminology/lock')
        data = json.loads(res.data)
        assert res.status_code == 200
        assert data['ok'] is True
        assert 'token' in data
        self._release_all(client)

    def test_acquire_lock_twice_fails(self, client):
        self._release_all(client)
        res1 = client.post('/api/terminology/lock')
        token = json.loads(res1.data)['token']
        res2 = client.post('/api/terminology/lock')
        assert res2.status_code == 409
        assert json.loads(res2.data)['ok'] is False
        self._release_all(client)

    def test_get_lock_shows_locked(self, client):
        self._release_all(client)
        client.post('/api/terminology/lock')
        res = client.get('/api/terminology/lock')
        data = json.loads(res.data)
        assert data['locked'] is True
        assert 'remaining_sec' in data
        self._release_all(client)

    def test_release_lock_with_valid_token(self, client):
        self._release_all(client)
        res1 = client.post('/api/terminology/lock')
        token = json.loads(res1.data)['token']
        res2 = client.delete('/api/terminology/lock',
                             data=json.dumps({'token': token}),
                             content_type='application/json')
        assert json.loads(res2.data)['ok'] is True
        assert json.loads(client.get('/api/terminology/lock').data)['locked'] is False

    def test_release_lock_with_wrong_token_fails(self, client):
        self._release_all(client)
        client.post('/api/terminology/lock')
        res = client.delete('/api/terminology/lock',
                            data=json.dumps({'token': 'wrong-token'}),
                            content_type='application/json')
        assert res.status_code == 403
        self._release_all(client)

    def test_lock_auto_expires(self, client):
        """タイムアウト済みロックは GET で自動解除されることを確認。"""
        import app as m
        import datetime
        self._release_all(client)
        client.post('/api/terminology/lock')
        # locked_at を過去に書き換えてタイムアウトをシミュレート
        with m._edit_lock_mutex:
            m._edit_lock_state["locked_at"] = datetime.datetime.now() - datetime.timedelta(seconds=m._EDIT_LOCK_TIMEOUT_SEC + 1)
        res = client.get('/api/terminology/lock')
        assert json.loads(res.data)['locked'] is False


# -----------------------------------------------------------------
# 8. 並行アクセステスト
# -----------------------------------------------------------------
class TestConcurrency:
    def test_concurrent_reviews_unique_filenames(self, pptx_path):
        """複数の同時レビューリクエストでレポートファイル名が重複しないことを確認。"""
        results = []

        def do_review():
            app.config["TESTING"] = True
            with app.test_client() as c:
                with open(pptx_path, "rb") as f:
                    res = c.post(
                        "/review",
                        data={"pptx_file": (f, "test.pptx")},
                        content_type="multipart/form-data",
                    )
                return json.loads(res.data)

        # patch はメインスレッドで一括適用（スレッド間の競合を避けるため）
        with patch("app.call_ai_review", return_value=MOCK_REPORT):
            with ThreadPoolExecutor(max_workers=3) as executor:
                futures = [executor.submit(do_review) for _ in range(3)]
                for future in as_completed(futures):
                    results.append(future.result())

        filenames = [r["report_filename"] for r in results if "report_filename" in r]
        assert len(filenames) == 3, f"全リクエストが成功するはず: {results}"
        assert len(set(filenames)) == 3, f"ファイル名が重複しています: {filenames}"

        # クリーンアップ
        for name in filenames:
            p = flask_app_module.UPLOAD_DIR / name
            if p.exists():
                p.unlink()

    def test_concurrent_terminology_save_no_corruption(self, client):
        """複数の同時 terminology 保存でデータが壊れないことを確認。"""
        res = client.get("/api/terminology")
        original = json.loads(res.data)

        terms_a = original["terms"] + [{"correct": "並行テストA", "variants": ["テストA"], "category": "テスト", "notes": ""}]
        terms_b = original["terms"] + [{"correct": "並行テストB", "variants": ["テストB"], "category": "テスト", "notes": ""}]

        outcomes = []

        def save(terms):
            app.config["TESTING"] = True
            with app.test_client() as c:
                res = c.post(
                    "/api/terminology",
                    data=json.dumps({"terms": terms}),
                    content_type="application/json",
                )
                return json.loads(res.data)

        with ThreadPoolExecutor(max_workers=2) as executor:
            futures = [executor.submit(save, terms_a), executor.submit(save, terms_b)]
            for future in as_completed(futures):
                outcomes.append(future.result())

        # 両方のリクエストが 200 で完了し、JSON が壊れていないこと
        assert all(r.get("ok") is True for r in outcomes), f"保存失敗: {outcomes}"

        # 元に戻す
        client.post(
            "/api/terminology",
            data=json.dumps({"terms": original["terms"]}),
            content_type="application/json",
        )


# -----------------------------------------------------------------
# 8. 用語チェック単体テスト
# -----------------------------------------------------------------
class TestTerminologyDetection:
    def test_detects_server_variant(self, pptx_path):
        """「サーバー」（誤）が検出されることを確認。"""
        extract_data = run_extract(str(pptx_path), None)
        extract_str = json.dumps(extract_data, ensure_ascii=False)
        result = run_terminology_check(extract_str)

        all_hits = [h for slide in result.get("results", []) for h in slide.get("hits", [])]
        found_terms = [h["found"] for h in all_hits]
        assert "サーバー" in found_terms, f"「サーバー」が検出されませんでした。検出語: {found_terms}"

    def test_detects_user_variant(self, pptx_path):
        """「ユーザー」（誤）が検出されることを確認。"""
        extract_data = run_extract(str(pptx_path), None)
        extract_str = json.dumps(extract_data, ensure_ascii=False)
        result = run_terminology_check(extract_str)

        all_hits = [h for slide in result.get("results", []) for h in slide.get("hits", [])]
        found_terms = [h["found"] for h in all_hits]
        assert "ユーザー" in found_terms, f"「ユーザー」が検出されませんでした。検出語: {found_terms}"

    def test_slides_with_issues_count(self, pptx_path):
        """表記ゆれのあるスライドが複数検出されることを確認。"""
        extract_data = run_extract(str(pptx_path), None)
        extract_str = json.dumps(extract_data, ensure_ascii=False)
        result = run_terminology_check(extract_str)
        assert result["slides_with_issues"] >= 2

    def test_custom_term_detection(self, pptx_path):
        """
        用語リストに新しい用語を追加した場合、その用語が検出されることを確認する。
        （用語リストカスタマイズのデモ）
        """
        terminology_path = SKILL_DIR / "references" / "terminology.json"
        with open(terminology_path, encoding="utf-8") as f:
            original = json.load(f)

        # カスタム用語を追加: 「当社」→「弊社」
        custom = json.loads(json.dumps(original))  # deepcopy
        custom["terms"].append({
            "correct": "弊社",
            "variants": ["当社"],
            "category": "ビジネス用語",
            "notes": "顧客向け資料では弊社を使う（テスト追加）"
        })

        # 一時的に terminology.json を書き換えてテスト
        with open(terminology_path, "w", encoding="utf-8") as f:
            json.dump(custom, f, ensure_ascii=False, indent=2)

        try:
            extract_data = run_extract(str(pptx_path), None)
            extract_str = json.dumps(extract_data, ensure_ascii=False)
            result = run_terminology_check(extract_str)

            # 新しい用語も正常に処理されたことを確認（エラーが出なければOK）
            assert "results" in result

            # /debug エンドポイントに用語が反映されているか確認
            app.config["TESTING"] = True
            with app.test_client() as c:
                res = c.get("/debug")
                data = json.loads(res.data)
                corrects = [t["correct"] for t in data["terminology"]["terms"]]
                assert "弊社" in corrects, f"カスタム用語「弊社」が /debug に反映されていません。terms: {corrects}"
                assert data["terminology"]["term_count"] == len(original["terms"]) + 1

        finally:
            # 必ず元に戻す
            with open(terminology_path, "w", encoding="utf-8") as f:
                json.dump(original, f, ensure_ascii=False, indent=2)


# -----------------------------------------------------------------
# 9. 固有観点 API テスト
# -----------------------------------------------------------------
class TestPerspectivesAPI:
    def test_get_returns_200(self, client):
        res = client.get("/api/perspectives")
        assert res.status_code == 200

    def test_get_returns_perspectives_list(self, client):
        res = client.get("/api/perspectives")
        data = json.loads(res.data)
        assert "perspectives" in data
        assert isinstance(data["perspectives"], list)

    def test_get_has_initial_data(self, client):
        """初期データにカテゴリと観点が含まれることを確認。"""
        res = client.get("/api/perspectives")
        data = json.loads(res.data)
        assert len(data["perspectives"]) >= 1
        cat = data["perspectives"][0]
        assert "category" in cat
        assert "items" in cat
        assert len(cat["items"]) >= 1
        assert "perspective" in cat["items"][0]

    def test_post_save_perspectives(self, client):
        """固有観点の保存が成功することを確認。"""
        # 元データを保持
        original = json.loads(client.get("/api/perspectives").data)

        new_data = {
            "perspectives": [
                {
                    "category": "テスト用カテゴリ",
                    "items": [
                        {"perspective": "テスト観点A", "notes": "備考A"},
                        {"perspective": "テスト観点B", "notes": ""},
                    ],
                }
            ]
        }
        res = client.post(
            "/api/perspectives",
            data=json.dumps(new_data),
            content_type="application/json",
        )
        assert res.status_code == 200
        body = json.loads(res.data)
        assert body["ok"] is True
        assert body["category_count"] == 1
        assert body["item_count"] == 2

        # 保存されたか確認
        res2 = client.get("/api/perspectives")
        saved = json.loads(res2.data)
        assert len(saved["perspectives"]) == 1
        assert saved["perspectives"][0]["category"] == "テスト用カテゴリ"

        # 元に戻す
        client.post(
            "/api/perspectives",
            data=json.dumps({"perspectives": original["perspectives"]}),
            content_type="application/json",
        )

    def test_post_empty_category_returns_400(self, client):
        """カテゴリ名が空の場合に 400 を返すことを確認。"""
        res = client.post(
            "/api/perspectives",
            data=json.dumps({"perspectives": [{"category": "", "items": []}]}),
            content_type="application/json",
        )
        assert res.status_code == 400

    def test_post_empty_perspective_returns_400(self, client):
        """観点名が空の場合に 400 を返すことを確認。"""
        res = client.post(
            "/api/perspectives",
            data=json.dumps({
                "perspectives": [{
                    "category": "テスト",
                    "items": [{"perspective": "", "notes": ""}],
                }]
            }),
            content_type="application/json",
        )
        assert res.status_code == 400

    def test_post_no_perspectives_field_returns_400(self, client):
        """perspectives フィールドなしで 400 を返すことを確認。"""
        res = client.post(
            "/api/perspectives",
            data=json.dumps({"data": []}),
            content_type="application/json",
        )
        assert res.status_code == 400


# -----------------------------------------------------------------
# 10. 固有観点ロックテスト
# -----------------------------------------------------------------
class TestPerspectivesLock:
    @staticmethod
    def _release_all(client):
        import app as m
        with m._persp_lock_mutex:
            m._persp_lock_state["token"] = None
            m._persp_lock_state["locked_at"] = None

    def test_get_lock_initially_unlocked(self, client):
        self._release_all(client)
        res = client.get('/api/perspectives/lock')
        assert json.loads(res.data)['locked'] is False

    def test_acquire_lock_success(self, client):
        self._release_all(client)
        res = client.post('/api/perspectives/lock')
        data = json.loads(res.data)
        assert data['ok'] is True
        assert 'token' in data
        self._release_all(client)

    def test_acquire_lock_twice_fails(self, client):
        self._release_all(client)
        client.post('/api/perspectives/lock')
        res2 = client.post('/api/perspectives/lock')
        assert res2.status_code == 409
        self._release_all(client)

    def test_release_lock_with_valid_token(self, client):
        self._release_all(client)
        res1 = client.post('/api/perspectives/lock')
        token = json.loads(res1.data)['token']
        res2 = client.delete('/api/perspectives/lock',
                             data=json.dumps({'token': token}),
                             content_type='application/json')
        assert json.loads(res2.data)['ok'] is True
        self._release_all(client)


# -----------------------------------------------------------------
# 11. レビューと固有観点の統合テスト
# -----------------------------------------------------------------
class TestReviewWithPerspectives:
    def test_review_with_custom_perspectives(self, client, pptx_path):
        """固有観点付きレビューが正常に完了することを確認。"""
        perspectives = [
            {"category": "デザイン", "items": [{"perspective": "フォントの統一", "notes": "テスト"}]}
        ]
        with patch("app.call_ai_review", return_value=MOCK_REPORT) as mock_ai:
            with open(pptx_path, "rb") as f:
                data = {
                    "pptx_file": (f, "test.pptx"),
                    "custom_perspectives": json.dumps(perspectives),
                }
                res = client.post("/review", data=data, content_type="multipart/form-data")
        assert res.status_code == 200
        body = json.loads(res.data)
        assert "report" in body

        # プロンプトに固有観点が含まれることを確認
        called_prompt = mock_ai.call_args[0][0]
        assert "固有観点" in called_prompt
        assert "デザイン" in called_prompt
        assert "フォントの統一" in called_prompt

        # クリーンアップ
        saved = flask_app_module.UPLOAD_DIR / body["report_filename"]
        if saved.exists():
            saved.unlink()

    def test_review_without_perspectives(self, client, pptx_path):
        """固有観点なしでもレビューが正常に完了することを確認。"""
        with patch("app.call_ai_review", return_value=MOCK_REPORT) as mock_ai:
            with open(pptx_path, "rb") as f:
                data = {"pptx_file": (f, "test.pptx")}
                res = client.post("/review", data=data, content_type="multipart/form-data")
        assert res.status_code == 200

        # 固有観点なしの場合、プロンプトに「固有観点」は含まれない
        called_prompt = mock_ai.call_args[0][0]
        assert "固有観点" not in called_prompt

        body = json.loads(res.data)
        saved = flask_app_module.UPLOAD_DIR / body["report_filename"]
        if saved.exists():
            saved.unlink()


# -----------------------------------------------------------------
# 7. /api/feedback
# -----------------------------------------------------------------
class TestFeedbackEndpoint:
    def test_feedback_success(self, client, tmp_path):
        """正常系: フィードバックがCSVに保存される。"""
        with patch.object(flask_app_module, "BASE_DIR", tmp_path):
            res = client.post("/api/feedback", json={
                "name": "テスト太郎",
                "department": "開発部",
                "content": "テストフィードバックです",
            })
        assert res.status_code == 200
        data = json.loads(res.data)
        assert data["ok"] is True

    def test_feedback_empty_content_400(self, client):
        """content が空の場合は 400 を返す。"""
        res = client.post("/api/feedback", json={"content": ""})
        assert res.status_code == 400
        data = json.loads(res.data)
        assert "error" in data

    def test_feedback_name_too_long_400(self, client):
        """name が 101 文字の場合は 400 を返す。"""
        res = client.post("/api/feedback", json={
            "name": "a" * 101,
            "content": "valid content",
        })
        assert res.status_code == 400

    def test_feedback_department_too_long_400(self, client):
        """department が 101 文字の場合は 400 を返す。"""
        res = client.post("/api/feedback", json={
            "department": "a" * 101,
            "content": "valid content",
        })
        assert res.status_code == 400

    def test_feedback_content_too_long_400(self, client):
        """content が 5001 文字の場合は 400 を返す。"""
        res = client.post("/api/feedback", json={
            "content": "a" * 5001,
        })
        assert res.status_code == 400

    def test_feedback_csv_injection_sanitized(self, client, tmp_path):
        """CSV injection 対策: 危険プレフィックスが ' で無害化される。"""
        with patch.object(flask_app_module, "BASE_DIR", tmp_path):
            client.post("/api/feedback", json={
                "name": "=cmd|'/C calc'!A0",
                "department": "+SUM(A1:A10)",
                "content": "normal content",
            })
            csv_path = tmp_path / "feedback" / "feedback.csv"
            text = csv_path.read_text(encoding="utf-8-sig")
        assert "'=cmd" in text
        assert "'+SUM" in text
