# -*- coding: utf-8 -*-
"""
Python: 3.12+
Module: core/core_cst
Created: 2025-12-05
Updated: 2026-03-09 (JST)
Version: 2.5.1
Purpose:
    アドインツール全体で使用される定数・UI設定を一括管理する。
    画面別の詳細設定は config/ui_*.json に移行済み。core_cst の古い画面設定（_CSV_MG / UI_SCREENS の実体）は削除し、UI_SCREENS は空 dict。

History (latest 3):
  - 2.5.1 (2026-04-16) `resolve_config_file_path`: `HC_INSTALL_ROOT` 配下の config を優先（配布の単一正本）。無ければ従来の core 親＝プロジェクト／EXE バンドルルートの config。
  - 2.5.0 (2026-03-09) 未使用定義を削除。APP_VERSION/BUILD_TAG, フォントサイズBOLD/SUB, Excel/VBA/Win32未使用定数, IPC, QT_PROGRESS_POLL_MS 等を整理。
  - 2.4.2 (2026-03-09) 古い画面設定を削除。_CSV_MG ブロックと UI_SCREENS 実体を廃止し、UI_SCREENS={} に。画面設定は config/ui_*.json に統一。
  - 2.4.1 (2026-03-09) オートフィット行数上限 AUTOFIT_MAX_ROWS を追加。全モジュールで行数基準に統一するため core に集約。
"""

# ==============================================================================
# システム共通（他モジュールから直接参照）
# ==============================================================================

APP_TITLE = "CSV Tool"

# フォント（Tk/svc 等での利用）
SUB_WINDOW_FONT_NAME = "Yu Gothic UI"
SUB_WINDOW_FONT_SIZE = 9

# エラー表示用セル背景色 (R, G, B) - hc_dupli 等で使用
ERR_BG_COLOR = (255, 255, 0)

# Undoキャッシュファイル名
CACHE_FILE_NAME = "header_converter_cache.pkl"

# svc 層
SVC_SKIP_FIND_SHEET_BY_GUID = False
SVC_EXCEL_WRITE_TARGET_CELLS = 50000
SVC_EXCEL_WRITE_MIN_ROWS = 100
SVC_EXCEL_WRITE_MAX_ROWS = 5000

# オートフィット実行の行数上限（超過時はスキップ）。全モジュールで行数基準に統一。1行でも列数が多くてもオートフィットする。
AUTOFIT_MAX_ROWS = 100000


# ==============================================================================
# 共有 dict（SYS）
# ==============================================================================

SYS = {
    "APP_TITLE": APP_TITLE,
    "TEMP_SUBDIR": "csv_tool",
    "LOG_MAX_BYTES": 1_048_576,
    "UI_SERVER_LOG_NAME": "ui_server.log",
}


# ==============================================================================
# UI 共通デフォルト（JSON とマージする際のベース。画面別の詳細は config/ui_*.json）
# ==============================================================================

UI_COMMON = {
    "WINDOW": {
        "RESIZABLE": True,
        "SHOW_MINIMIZE": False,
        "SHOW_MAXIMIZE": False,
        "STARTUP_POSITION": "center",
        "CENTER_ON_EXCEL": True,
    },
    "FONT": {
        "FAMILY": "",
        "SIZE_PT": 10,
        "BOLD_HEADERS": True,
    },
    "TABLE": {
        "ALTERNATE_ROW_COLORS": True,
    },
}

# 画面別設定は config/ui_<feature>.json に移行済み。get_ui_config2 は CSV_MG 等をファイルから取得するため空でよい。
UI_SCREENS = {}

# ==============================================================================
# 機能別画面設定ファイル（外部のみ・救済なし）
# ==============================================================================
# CSV_MG 等は config/ui_<feature_key小文字>.json のみを参照する。
# ファイルが無い・JSON が壊れている場合はエラー種別を表示して終了し、救済（core_cst フォールバック）は行わない。

import json
from pathlib import Path as _Path

_CONFIG_FILE_CACHE: dict[tuple[str, str], tuple[float, dict]] = {}


class UiConfigLoadError(RuntimeError):
    """設定ファイル読み込み失敗時に使用。エラー種別とメッセージを保持する。"""
    def __init__(self, kind: str, path: str, detail: str = ""):
        self.kind = kind
        self.path = path
        self.detail = detail
        msg = f"[{kind}] {path}"
        if detail:
            msg += f"\n{detail}"
        super().__init__(msg)


def resolve_config_file_path(filename: str) -> _Path:
    """
    `config` 配下の1ファイルの絶対パスを解決する。

    1. 環境変数 **`HC_INSTALL_ROOT`** が有効なディレクトリを指し、
       **`<インストールルート>\\config\\<filename>` が存在する** → そのパス（配布の単一正本）。
    2. それ以外 → **`<core_cst の親の親>\\config\\<filename>`**（開発時はリポジトリルートの `config\\`。
       Nuitka 単体 EXE では各バンドルに `config\\` を同梱しない方針のため、**配布運用では 1 を満たすこと**）。

    filename はベース名のみ（例: ``ui_csv_mg.json``, ``svc_warmup.json``）。
    """
    if not filename or not isinstance(filename, str):
        raise ValueError("filename must be a non-empty str")
    fn = filename.strip().replace("\\", "/").split("/")[-1]
    if not fn or fn != filename.strip():
        raise ValueError("filename must be a basename only")
    if ".." in fn or fn.startswith(("/", "\\")):
        raise ValueError("invalid filename")

    from core import runtime_layout

    ir = runtime_layout.install_root()
    if ir is not None:
        p = ir / "config" / fn
        if p.is_file():
            return p.resolve()
    return (_Path(__file__).resolve().parent.parent / "config" / fn).resolve()


def get_ui_config_from_file_required(feature_key: str) -> dict:
    """
    機能別設定ファイルを読み、辞書を返す。必須。失敗時は救済せず UiConfigLoadError を発生させる。
    ファイル: **resolve_config_file_path** により決まる `config/ui_<feature_key小文字>.json`
    """
    if not feature_key or not isinstance(feature_key, str):
        raise UiConfigLoadError("引数不正", "", "feature_key が空です")
    fk = str(feature_key).strip().lower()
    if not fk:
        raise UiConfigLoadError("引数不正", "", "feature_key が空です")
    path = resolve_config_file_path(f"ui_{fk}.json")
    if not path.is_file():
        raise UiConfigLoadError("設定ファイルが見つかりません", str(path))
    try:
        with path.open("r", encoding="utf-8") as f:
            data = json.load(f)
    except json.JSONDecodeError as e:
        raise UiConfigLoadError(
            "設定ファイルの形式が正しくありません（JSON エラー）",
            str(path),
            f"行{e.lineno} 列{e.colno}: {e.msg}",
        ) from e
    except OSError as e:
        raise UiConfigLoadError("設定ファイルの読み込みに失敗しました", str(path), str(e)) from e
    except Exception as e:
        raise UiConfigLoadError("設定ファイルの読み込みに失敗しました", str(path), f"{type(e).__name__}: {e}") from e
    if not isinstance(data, dict):
        raise UiConfigLoadError("設定の形式が正しくありません", str(path), "ルートは JSON オブジェクトである必要があります")
    # ヘッダ・区切りは設定として使わない
    data = {k: v for k, v in data.items() if k not in ("_header", "_separator")}
    # キャッシュ（パスごと・更新時のみ再読込）
    try:
        mtime = path.stat().st_mtime
        _CONFIG_FILE_CACHE[(fk, str(path))] = (mtime, data)
    except Exception:
        pass
    return data
