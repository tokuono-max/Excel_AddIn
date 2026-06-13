# -*- coding: utf-8 -*-
"""
Python: 3.12
Module: ui_qt/ui_dialog_progress.py
Created: 2026-03-10
Updated: 2026-06-06
Version: 0.1.35
Purpose:
  進捗表示用ダイアログ（ProgressDialog）および create_progress_dialog を提供する。
  ui_common の Progress 実装を本モジュールへ移し、画面種別ごとの責務分離を行う。

実行:
  - ProgressDialog クラス全体を ui_common から本モジュールへ移動。
  - create_progress_dialog は本モジュールで実装し、ダイアログ生成後に _keep_modeless で保護して返す。
  - ヘルパは ui_common から import。完了通知表示は ui_dialog_done.create_done_dialog を利用。
  - 呼び出し側は ui_common.create_progress_dialog / create_dialog 経由のため既存コード変更不要。

History (latest 3):
  - 0.1.35 (2026-06-13) 砂時計: Qt WaitCursor を進捗表示中に付与。Excel xlWait は遅延再武装。
  - 0.1.34 (2026-06-13) DONE 終了アニメ中は直前工程ではなく「仕上げ中…」を表示（オートフィット表示の残りを防止）。
  - 0.1.33 (2026-06-13) DONE: バー100%到達後に「完了」表示。終了アニメは加速 creep で短時間化。
  - 0.1.28 (2026-06-13) 完了通知: partner 無しは DoneDialog をモデルレス show（exec 廃止）で ui_server ブロック・WaitForm タイムアウトを防止。
  - 0.1.27 (2026-06-13) 工程3でも creep を適用し進捗バーの瞬間ジャンプを抑制。
  - 0.1.26 (2026-06-06) 工程3は creep 無効で pct 即反映。DONE クローズは UI 先・cursor OFF はタイムアウト付き。
  - 0.1.25 (2026-06-06) showEvent: WaitForm 解除・砂時計 ON を excel_lock より先に実行。
  - 0.1.24 (2026-06-06) 砂時計は showEvent で ON・teardown で OFF のみ（進捗 pickle 更新では制御しない）。
  - 0.1.23 (2026-06-04) RUN 時: pickle seq 更新のたびに砂時計を即再武装（svc 側 OFF 後も進捗表示中は維持）。
  - 0.1.21 (2026-06-04) RUN 時: pickle の pct が明示されていれば done/total より優先（CSV保存等で 50% フェーズが 99% に張り付くのを防止）。
  - 0.1.20 (2026-05-05) req.progress_closed_path 指定時、進捗クローズ時に ACK pickle を書き出す。svc 側がクローズ完了を待って結果画面を出せるようにし、重なり表示を抑制。
  - 0.1.19 (2026-05-03) _teardown_progress_shared_state: excel_unlock 未指定時は parent_hwnd があれば常に True（excel_lock false でも他経路の Win32 無効取り残しを解除）。CANCEL も同様に parent があるときは解除。
  - 0.1.18 (2026-05-02) teardown に stop_front_follow_match_widget=self。プレビューが先に start_front_follow した場合は進捗 teardown でフックを潰さない。
  - 0.1.17 (2026-05-02) 終了時 teardown_feature_ui_shared_state: stop_front_follow を destroy 待ちにせず実行。closeEvent / _close_after_done / OVER_LIMIT / ERROR / CANCEL。
  - 0.1.16 (2026-04-10) hide_cancel_button: キャンセル行を QWidget で包み行ごと非表示。中央 stretch を縮め、固定高さ時は内容に合わせて高さを再確定（余白低減）。
  - 0.1.15 (2026-04-10) DONE pickle の done_delay_ms で閉じるまでの待ちを上書き可（重複チェック・未着色キャンセル時の短い表示用）。
  - 0.1.14 (2026-04-10) req.cancel_request_path 指定時のみキャンセルボタン表示。押下で専用 pickle に cancel を書込。RUN.hide_cancel_button でボタン非表示。
  - 0.1.13 (2026-04-09) raise_csv_sp_partner_progress: 重複確認終了前に csv_sp 進捗を前面化（モーダル背後の Z-order／ゴースト枠対策）。
  - 0.1.12 (2026-04-09) CANCEL 時 partner の Shiboken.isValid と csv_sp IPC 再オープン。完了前に clear_sp_progress_partner_phase。
  - 0.1.11 (2026-04-08) CANCEL/ERROR は seq 古さで無視しない。CANCEL 経路の診断ログ。DONE に output_dir を渡すための保持。
  - 0.1.10 (2026-04-08) CANCEL 時、partner_widget_after_cancel がある場合は excel_lock に関係なく Excel を一度有効化してから分割を再表示。
  - 0.1.9 (2026-04-08) partner 再表示前に WA_DontShowOnScreen を解除（分割画面の進捗直前 deflake との整合）。
  - 0.1.8 (2026-04-08) CANCEL 時に partner_widget_after_cancel で分割画面を再表示（csv_sp 重複確認キャンセル）。
  - 0.1.7 (2026-04-08) req.partner_widget_after_done: 完了後に次イベントで分割画面を accept/close（csv_sp exec 解放・再入回避）。
  - 0.1.6 (2026-04-08) 進捗のちらつき・位置跳び抑制: show 前の二重センタ／オーナー廃止、showEvent でオーナー→中央→前面化を1回、150ms 再前面化削除、no_native の opacity0 廃止、DONE 再配置を単発タイマーに集約。
  - 0.1.5 (2026-04-05) _hc_show_taskbar 取得失敗時はタスクバー非表示扱い。owner 設定を遅延リトライ。showEvent で ensure_front を遅延再実行。
  - 0.1.4 (2026-04-05) DONE 受信時 INFO ログ。ERROR/CANCEL の Pickle 処理。_close_after_done の完了条件を show_done_dialog と (非空 items または detail) に明示。
"""

from __future__ import annotations

import os
import time
from pathlib import Path
from typing import Optional

from PySide6.QtCore import QTimer, Qt
from PySide6.QtWidgets import (
    QApplication,
    QDialog,
    QHBoxLayout,
    QLabel,
    QMessageBox,
    QProgressBar,
    QPushButton,
    QSizePolicy,
    QSpacerItem,
    QVBoxLayout,
    QWidget,
)

# 変数: 進捗データのやり取り用（svc が Pickle で status / pct / phase 等を書き、本ダイアログがポーリングで読む）
from ui_qt import ipc_file
# 変数: ui_common のヘルパ（中央配置・ウィンドウ設定・モデルレス管理・Excel操作有効化・トレース等）
from ui_qt.ui_common import (
    _close_all_modeless,
    _get_progress_config,
    _normalize_message_newlines,
    _progress_done_recenter,
    _w32,
    apply_common_window_style,
    apply_tooltip_if_set,
    apply_window_config,
    center_on_excel,
    center_on_parent_widget,
    enable_excel_window,
    ensure_front,
    _set_owner_hwnd,
    _keep_modeless,
    teardown_feature_ui_shared_state,
)
from ui_qt.ui_dialog_done import create_done_dialog

# 変数: 進捗フロー計測用ログ（CSV_LD_FLOW 等）。core 未利用環境では None にフォールバック
try:
    from core import core_log
    _log = core_log.get_logger(__name__)
except Exception:  # pragma: no cover
    core_log = None  # type: ignore
    _log = None  # type: ignore
try:
    from core.core_log import get_diag_logger

    _diag_ui = get_diag_logger("hc_csv_tool.diag.ui_progress")
except Exception:  # pragma: no cover
    _diag_ui = None  # type: ignore

__version__ = "0.1.35"

# DONE 受信後、バーが 100% に達するまでの暫定ラベル（直前工程名のまま残さない）
DONE_FINISH_INTERIM_LABEL = "仕上げ中…"


def _progress_label_min_height_for_lines(label: QLabel, lines: int = 2) -> int:
    """2 行表示が切れないようラベルの最小高さ（px）を返す。"""
    n = max(1, int(lines))
    try:
        fm = label.fontMetrics()
        lh = int(fm.lineSpacing() or fm.height() or 16)
        return max(lh * n + 2, 30)
    except Exception:
        return 40


def _format_progress_status_text(
    *,
    phase: str,
    msg: str,
    detail: str,
    current_file: str,
    window_title: bool,
) -> str:
    """進捗ラベル用: フェーズ＋詳細（最大 2 行）。detail があるときは ［ファイル名］ 行を付けない。"""
    phase_s = str(phase or "").strip()
    msg_s = str(msg or "").strip()
    detail_s = str(detail or "").strip()
    cf = str(current_file or "").strip()
    if detail_s:
        line1 = phase_s or "準備中"
        return "%s\n%s" % (line1, detail_s)
    parts: list[str] = []
    if phase_s:
        parts.append(phase_s)
    if msg_s and msg_s != phase_s:
        parts.append(msg_s)
    body = "\n".join(parts) if parts else "準備中"
    # detail 無しの旧形式のみ。ファイル名が本文に無いときだけ ［］ 行を足す。
    if cf and cf not in body:
        return "%s\n［%s］" % (body, cf)
    return body


def compute_run_progress_bar_pct(
    *,
    svc_pct: int,
    prev_bar: int,
    creep: int,
    phase_i: int,
    display_target: int,
    center_on_parent: bool = False,
) -> tuple[int, int]:
    """RUN 更新時の進捗バー pct と display_target。svc pct へは creep で段階的に追従する。"""
    pct = max(0, min(99, int(svc_pct)))
    if center_on_parent and pct < prev_bar:
        pct = prev_bar
    if creep <= 0:
        tgt = max(int(display_target), pct)
        return (max(prev_bar, pct), tgt)
    tgt = max(int(display_target), pct)
    cur_bar = int(prev_bar)
    if tgt > cur_bar:
        pct = min(tgt, cur_bar + int(creep))
    else:
        pct = tgt
    if pct < prev_bar:
        pct = prev_bar
    return (pct, tgt)


from core.csv_ld_progress_ack import (
    compute_bar_creep_next_value,
    compute_done_close_delay_ms,
    compute_done_finish_creep_pct,
)


class ProgressDialog(QDialog):
    """
    【概要】
        進捗表示用ダイアログ（モデルレス想定）。IPC ファイルをポーリングして進捗を更新する。
    【補足】
        progress_cfg 未指定時は _get_progress_config() で既定設定を取得。各機能は progress_cfg を渡して自機能の設定を使える。
    """

    def __init__(
        self,
        req: dict,
        parent_hwnd: int,
        parent: Optional[QWidget] = None,
        progress_cfg: Optional[dict] = None,
    ) -> None:
        super().__init__(parent if parent is not None else None)
        # 命令: 既定は親がいれば WindowModal。non_modal_progress 時は NonModal（デバッグ等で親を操作不能にしない）
        if parent is not None:
            try:
                if bool(req.get("non_modal_progress", False)):
                    self.setWindowModality(Qt.WindowModality.NonModal)
                else:
                    self.setWindowModality(Qt.WindowModality.WindowModal)
            except Exception:
                try:
                    self.setWindowModality(Qt.WindowModal)
                except Exception:
                    pass

        # 判定: no_native_window が True のときは WA_NativeWindow を付けない（csv_ld 等で枠だけになる問題対策）
        if not bool(req.get("no_native_window", False)):
            try:
                self.setAttribute(Qt.WA_NativeWindow, True)
                self.winId()
            except Exception:
                pass

        # 変数: 画面設定（progress_cfg 未指定時は _get_progress_config で CSV_MG 用既定を取得）
        _cfg = progress_cfg if isinstance(progress_cfg, dict) else _get_progress_config()
        self._done_cfg = (_cfg or {}).get("_done_cfg")
        # 変数: DONE 到達時に完了通知を出すかどうかと、その際の items / detail_text（_close_after_done で使用）
        self._pending_show_done_dialog = False
        self._pending_done_items: list = []
        self._pending_done_detail_text: Optional[str] = None
        self._pending_output_dir: Optional[str] = None
        # TITLE キーが明示されていて空ならベースタイトル無し（window_title のみ表示用）。
        if isinstance(_cfg, dict) and "TITLE" in _cfg and not str(_cfg.get("TITLE") or "").strip():
            title_raw = ""
        else:
            title_raw = str(_cfg.get("TITLE") or "ファイル結合 進捗").strip(" \t\r")
        self._progress_base_title = _normalize_message_newlines(title_raw)
        self.setWindowTitle(self._progress_base_title)
        apply_tooltip_if_set(self, _cfg, "TOOLTIP")
        # 変数: 進捗 Pickle ファイルパス（svc が書き、本ダイアログが 200ms 間隔でポーリング）
        self._progress_path = Path(str(req.get("progress_path", "") or ""))
        self._phase_total = int(req.get("phase_total", 0) or 0)
        self._parent_hwnd = int(parent_hwnd or 0)
        _closed_path_raw = str(req.get("progress_closed_path") or "").strip()
        self._progress_closed_path: Optional[Path] = Path(_closed_path_raw) if _closed_path_raw else None
        self._progress_closed_ack_written = False
        # 変数: 表示時点の Excel 矩形（svc が GetWindowRect で渡した場合、中央配置のずれを抑える）
        _er = req.get("excel_rect")
        self._excel_rect = None
        if _er and len(_er) >= 4:
            try:
                self._excel_rect = (int(_er[0]), int(_er[1]), int(_er[2]), int(_er[3]))
            except (TypeError, ValueError):
                pass
        # 変数: csv_ld 等が True を渡した場合のみ、表示中は Excel 操作を無効化
        self._excel_lock = bool(req.get("excel_lock", False))
        # 変数: True のときは show 後 1 フレームだけ中央・オーナー（HWND が遅延するため）
        self._no_native_window = bool(req.get("no_native_window", False))
        # データ集約デバッグ等: Excel ではなく親 QWidget の中央に進捗を置く
        self._center_on_parent_widget = bool(req.get("center_on_parent_widget", False))
        # 完了時に親 QWidget を close するか（データ集約デバッグ等は False）
        self._req: dict = dict(req) if isinstance(req, dict) else {}
        try:
            self._done_delay_ms = int(self._req.get("done_delay_ms", 1000))
        except (TypeError, ValueError):
            self._done_delay_ms = 1000
        self._done_delay_ms = max(0, min(30000, self._done_delay_ms))

        # 変数: UI 部品（上: 処理中ファイル名・工程、中: 進捗バー、下: done/total 右寄せ）
        self._label_file = QLabel("準備中...")
        self._label_file.setAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignTop)
        # 折り返しは使わない（phase + detail の 2 行固定。wrap すると疑似 3 段目が見切れる）
        try:
            self._label_file.setWordWrap(False)
        except Exception:
            pass
        try:
            _lh2 = _progress_label_min_height_for_lines(self._label_file, 2)
            self._label_file.setMinimumHeight(_lh2)
            self._label_file.setMaximumHeight(_lh2)
        except Exception:
            pass
        self._bar = QProgressBar()
        self._bar.setRange(0, 100)
        self._bar.setValue(0)
        try:
            self._bar.setMinimumHeight(22)
            self._bar.setMaximumHeight(22)
        except Exception:
            pass
        self._label_count = QLabel("0 / 0")
        self._label_count.setAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignTop)
        try:
            _lc1 = _progress_label_min_height_for_lines(self._label_count, 1)
            self._label_count.setMinimumHeight(_lc1)
            self._label_count.setMaximumHeight(_lc1)
        except Exception:
            pass

        # 変数: レイアウト（上: 2 行テキスト、中: バー、下: done/total、最下: キャンセル）
        lay = QVBoxLayout(self)
        try:
            lay.setContentsMargins(8, 6, 8, 6)
            lay.setSpacing(2)
        except Exception:
            pass
        lay.addWidget(self._label_file)
        lay.addWidget(self._bar)
        lay.addWidget(self._label_count)
        self._progress_mid_spacer = None

        self._cancel_request_path = str(req.get("cancel_request_path") or "").strip()
        self._btn_cancel: Optional[QPushButton] = None
        self._cancel_row_widget: Optional[QWidget] = None
        self._cancel_ui_compacted = False
        if self._cancel_request_path:
            cw = QWidget()
            row_c = QHBoxLayout(cw)
            row_c.setContentsMargins(0, 0, 0, 0)
            row_c.addStretch(1)
            self._btn_cancel = QPushButton(str(_cfg.get("BTN_PROGRESS_CANCEL") or "キャンセル"))
            try:
                self._btn_cancel.setFixedHeight(24)
            except Exception:
                pass
            try:
                tt_c = str(_cfg.get("PROGRESS_CANCEL_TOOLTIP") or "").strip()
                if tt_c:
                    self._btn_cancel.setToolTip(tt_c)
            except Exception:
                pass
            self._btn_cancel.clicked.connect(self._on_cancel_clicked)  # type: ignore[attr-defined]
            row_c.addWidget(self._btn_cancel)
            self._cancel_row_widget = cw
            lay.addWidget(cw)

        # 命令分離: 画面固有 WINDOW（TOPMOST / SHOW_MINIMIZE / タスクバー等）を JSON 設定で適用
        try:
            apply_window_config(self, _cfg, int(parent_hwnd or 0), "PROGRESS")
        except Exception:
            apply_common_window_style(self, int(parent_hwnd or 0))
        try:
            if isinstance(_cfg, dict):
                self._hc_prepare_window_cfg = dict(_cfg.get("WINDOW") or {})
            else:
                self._hc_prepare_window_cfg = {}
        except Exception:
            self._hc_prepare_window_cfg = {}

        if parent is not None and bool(req.get("non_modal_progress", False)):
            try:
                fl = self.windowFlags()
                fl |= Qt.WindowType.WindowStaysOnTopHint
                self.setWindowFlags(fl)
            except Exception:
                pass

        # 変数: ポーリング用タイマー（既定 200ms。req の progress_poll_ms で 50〜500 に調整可）
        try:
            _poll = int(self._req.get("progress_poll_ms", 200) or 200)
        except (TypeError, ValueError):
            _poll = 200
        _poll = max(50, min(500, _poll))
        self._timer = QTimer(self)
        self._timer.setInterval(_poll)
        self._timer.timeout.connect(self._tick)  # type: ignore[attr-defined]
        self._timer.start()
        self._bar_creep_timer = QTimer(self)
        self._bar_creep_timer.setInterval(_poll)
        self._bar_creep_timer.timeout.connect(self._bar_creep_tick)  # type: ignore[attr-defined]
        self._bar_creep_timer.start()
        # 初回 RUN / DONE を core_log に 1 回だけ出力するためのフラグ
        self._flow_logged_run = False
        self._flow_logged_done = False
        # 進捗の順序保証: RUN を一度でも表示したか。DONE は _run_seen が True のときだけ処理する。
        self._run_seen = False
        self._last_seen_seq = -1
        try:
            _creep = int(self._req.get("progress_bar_creep_pct", 0) or 0)
        except (TypeError, ValueError):
            _creep = 0
        self._progress_bar_creep_pct = max(0, min(10, _creep))
        self._progress_display_target = 0
        self._last_run_phase_i = 0
        self._bar_creep_done_phase = False
        self._pending_done_label = None
        self._progress_bar_done_creep_pct = 0
        self._qt_wait_cursor_armed = False
        # データ集約デバッグ進捗: phase_i / done / total 表示の単調化（戻り見え防止）
        self._nm_pi_disp = 0
        self._nm_done_disp = 0
        self._nm_tot_disp = 0

    def _apply_pending_done_label(self) -> None:
        lbl = getattr(self, "_pending_done_label", None)
        if not lbl:
            return
        try:
            self._label_file.setText(str(lbl))
        except Exception:
            pass
        self._pending_done_label = None

    def _bar_creep_tick(self) -> None:
        """pickle 更新とは独立に進捗バーを段階表示。target 到達後も RUN 中はソフト上限まで進める。"""
        try:
            creep_base = int(getattr(self, "_progress_bar_creep_pct", 0) or 0)
            if creep_base <= 0:
                return
            done_pending = bool(getattr(self, "_bar_creep_done_phase", False))
            creep = (
                int(getattr(self, "_progress_bar_done_creep_pct", 0) or 0)
                if done_pending
                else creep_base
            )
            if creep <= 0:
                creep = creep_base
            prev = int(self._bar.value())
            tgt = int(getattr(self, "_progress_display_target", 0) or 0)
            phase_i = int(getattr(self, "_last_run_phase_i", 0) or 0)
            run_active = bool(getattr(self, "_run_seen", False)) and not done_pending
            nxt = compute_bar_creep_next_value(
                prev_bar=prev,
                display_target=tgt,
                creep=creep,
                phase_i=phase_i,
                run_active=run_active,
                done_pending=done_pending,
            )
            if nxt > prev:
                self._bar.setValue(nxt)
            if done_pending and int(self._bar.value()) >= 100:
                self._apply_pending_done_label()
        except Exception:
            pass

    def _stop_progress_timers(self) -> None:
        for _tm in (getattr(self, "_timer", None), getattr(self, "_bar_creep_timer", None)):
            if _tm is None:
                continue
            try:
                _tm.stop()
            except Exception:
                pass

    def _on_cancel_clicked(self) -> None:
        p = str(getattr(self, "_cancel_request_path", "") or "").strip()
        if not p:
            return
        is_data_agg_batch = "cancel_req_data_agg_batch" in p
        if _log is not None:
            try:
                _log.info(
                    "[DATA_AGG] progress cancel clicked path=%s data_agg_batch=%s",
                    p,
                    is_data_agg_batch,
                )
            except Exception:
                pass
        if is_data_agg_batch:
            try:
                self._label_file.setText("中止しています…")
            except Exception:
                pass
        try:
            ipc_file.write_pickle(Path(p), {"cancel": True, "v": 1})
        except Exception:
            pass
        if is_data_agg_batch:
            try:
                from svc.data_agg_cancel import force_data_agg_batch_cancel_from_ui  # noqa: WPS433

                force_data_agg_batch_cancel_from_ui(
                    cancel_path=Path(p),
                    progress_path=getattr(self, "_progress_path", None),
                    notify_parent=bool(self._req.get("data_agg_batch_notify_parent")),
                    parent_hwnd=int(self._parent_hwnd or 0),
                    scenario_id=str(self._req.get("data_agg_batch_scenario_id") or ""),
                    scenario_path=str(self._req.get("data_agg_batch_scenario_path") or ""),
                )
            except Exception as _term_exc:
                if _log is not None:
                    try:
                        _log.warning(
                            "[DATA_AGG] batch force terminate failed: %s",
                            _term_exc,
                        )
                    except Exception:
                        pass
        try:
            b = getattr(self, "_btn_cancel", None)
            if b is not None:
                b.setEnabled(False)
        except Exception:
            pass

    def showEvent(self, event) -> None:  # type: ignore[override]
        """
        【概要】
            excel_lock なら Excel 操作をロック。表示直後 1 回だけオーナー→中央→前面化（ネイティブ／no_native 共通方針でちらつき抑制）。
        """
        try:
            self._progress_wait_cursor_on()
            if getattr(self, "_excel_lock", False) and self._parent_hwnd:
                try:
                    enable_excel_window(self._parent_hwnd, False)
                except Exception:
                    pass
            if _log is not None:
                import time as _t
                _log.debug("[CSV_LD_FLOW] ProgressDialog showEvent (window visible) t=%.3f", _t.time())

            def _after_show_stack() -> None:
                ph = int(self._parent_hwnd or 0)
                rect = getattr(self, "_excel_rect", None)
                try:
                    show_tb = False
                    try:
                        show_tb = bool(self.property("_hc_show_taskbar"))
                    except Exception:
                        show_tb = False

                    if getattr(self, "_no_native_window", False):
                        if getattr(self, "_center_on_parent_widget", False):
                            pw = self.parentWidget()
                            if pw is not None:
                                try:
                                    ph_eff = int(pw.winId())
                                except Exception:
                                    ph_eff = 0
                                center_on_parent_widget(self, pw)
                                if not show_tb and ph_eff:
                                    _set_owner_hwnd(self, ph_eff)
                                if ph_eff:
                                    ensure_front(self, ph_eff)
                                try:
                                    self.setWindowOpacity(1.0)
                                except Exception:
                                    pass
                                return
                        ph_eff = ph
                        if ph_eff == 0:
                            pw = self.parentWidget()
                            if pw is not None:
                                try:
                                    ph_eff = int(pw.winId())
                                except Exception:
                                    ph_eff = 0
                        if ph_eff or rect:
                            if not show_tb and ph_eff:
                                _set_owner_hwnd(self, ph_eff)
                            center_on_excel(self, ph_eff, rect)
                            if ph_eff:
                                ensure_front(self, ph_eff)
                        try:
                            self.setWindowOpacity(1.0)
                        except Exception:
                            pass
                    elif ph or rect:
                        if not show_tb and ph:
                            _set_owner_hwnd(self, ph)
                        center_on_excel(self, ph, rect)
                        if ph:
                            ensure_front(self, ph)
                except Exception:
                    pass

            QTimer.singleShot(0, _after_show_stack)
        except Exception:
            pass
        super().showEvent(event)

    def _compact_progress_after_cancel_hidden(self) -> None:
        """キャンセル行を隠したあと、stretch とウィンドウ高さを内容に寄せる（固定高さの余白低減）。"""
        if getattr(self, "_cancel_ui_compacted", False):
            return
        self._cancel_ui_compacted = True
        try:
            sp = getattr(self, "_progress_mid_spacer", None)
            if sp is not None:
                sp.changeSize(0, 0, QSizePolicy.Policy.Minimum, QSizePolicy.Policy.Minimum)
            lay = self.layout()
            if lay is not None:
                lay.invalidate()
                lay.activate()
        except Exception:
            pass
        try:
            w = max(int(self.width()), 1)
            fixed_h = self.minimumHeight() > 0 and self.minimumHeight() == self.maximumHeight()
            self.setMinimumHeight(0)
            self.setMaximumHeight(16777215)
            self.adjustSize()
            nh = max(int(self.sizeHint().height()), 1)
            if fixed_h:
                self.setFixedSize(w, nh)
            else:
                self.resize(w, nh)
        except Exception:
            pass

    def _progress_sheet_id(self) -> str:
        return str(self._req.get("sheet_id") or "progress")

    def _progress_qt_wait_cursor_on(self) -> None:
        if getattr(self, "_qt_wait_cursor_armed", False):
            return
        try:
            QApplication.setOverrideCursor(Qt.CursorShape.WaitCursor)
            self._qt_wait_cursor_armed = True
        except Exception:
            pass

    def _progress_qt_wait_cursor_off(self) -> None:
        if not getattr(self, "_qt_wait_cursor_armed", False):
            return
        try:
            QApplication.restoreOverrideCursor()
        except Exception:
            pass
        self._qt_wait_cursor_armed = False

    def _progress_excel_wait_cursor_on(self) -> None:
        try:
            from core.core_cursor import progress_dialog_wait_cursor_on

            progress_dialog_wait_cursor_on(self._progress_sheet_id())
        except Exception:
            pass

    def _schedule_progress_wait_cursor_retries(self) -> None:
        """Excel 側 xlWait はフォーカス/マウス移動まで反映されないことがあるため再武装する。"""
        for ms in (100, 300, 800):
            QTimer.singleShot(int(ms), self._progress_excel_wait_cursor_on)

    def _progress_wait_cursor_on(self) -> None:
        self._progress_qt_wait_cursor_on()
        self._progress_excel_wait_cursor_on()
        self._schedule_progress_wait_cursor_retries()

    def _progress_wait_cursor_off(self) -> None:
        self._progress_qt_wait_cursor_off()
        try:
            from core.core_cursor import progress_dialog_wait_cursor_off

            progress_dialog_wait_cursor_off(cancel_reason="progress_dialog_done")
        except Exception:
            pass

    def _teardown_progress_shared_state(
        self, excel_unlock: Optional[bool] = None, *, cursor_off: bool = True
    ) -> None:
        """共有 UI 状態を片付ける。

        excel_unlock 未指定時: parent_hwnd があれば True（進捗がロックしていなくても、他処理の
        enable_excel_window(False) 取り残しを解除する）。
        """
        if excel_unlock is None:
            excel_unlock = bool(int(self._parent_hwnd or 0))
        teardown_feature_ui_shared_state(
            parent_hwnd=int(self._parent_hwnd or 0),
            modeless_widget=self,
            excel_unlock=excel_unlock,
        )
        if cursor_off:
            self._progress_wait_cursor_off()

    def _tick(self) -> None:
        """
        【概要】
            IPC ファイル（Pickle）から進捗情報を読み取り、status に応じて画面を更新する。
        【分岐】
            DONE → タイマー停止・完了表示・1 秒後に _close_after_done。OVER_LIMIT → 警告表示・return_merge 投入・閉じる。それ以外 → 通常の進捗表示更新。
        """
        try:
            if not self._progress_path.exists() or self._progress_path.stat().st_size <= 0:
                return

            d = ipc_file.read_pickle(self._progress_path)
            if not isinstance(d, dict):
                return

            status = str(d.get("status", "") or "").strip()
            status_u = status.upper()
            # 順序保証: seq が無い（従来形式）は -1 として常に採用。seq が前回以下なら古い更新なので無視
            try:
                seq = int(d.get("seq", -1))
            except (TypeError, ValueError):
                seq = -1
            # CANCEL/ERROR は seq の古さで無視しない（重複キャンセル等で必ず処理する）
            if status_u not in ("CANCEL", "ERROR"):
                if seq >= 0 and seq <= getattr(self, "_last_seen_seq", -1):
                    return
                self._last_seen_seq = seq
            else:
                if seq >= 0:
                    self._last_seen_seq = max(seq, getattr(self, "_last_seen_seq", -1))
            # DONE は RUN を一度でも表示してから処理。いきなり DONE の場合は _run_seen を立ててそのまま完了表示（極端に速い処理用）
            if status_u == "DONE" and not getattr(self, "_run_seen", False):
                self._run_seen = True
            # 初回 RUN / DONE を hc_csv.log に 1 回だけ出力（フロー計測用）
            try:
                import time as _t
                if status_u == "DONE":
                    if not getattr(self, "_flow_logged_done", False):
                        self._flow_logged_done = True
                        if _log is not None:
                            _log.debug("[CSV_LD_FLOW] ProgressDialog _tick: first DONE seen t=%.3f", _t.time())
                elif status_u == "RUN" and not getattr(self, "_flow_logged_run", False):
                    self._flow_logged_run = True
                    if _log is not None:
                        _log.debug("[CSV_LD_FLOW] ProgressDialog _tick: first RUN seen t=%.3f", _t.time())
            except Exception:
                pass

            # 判定: 完了時は「完了」表示を短時間残してから _close_after_done で進捗を閉じ、必要なら完了通知を表示
            if status_u == "DONE":
                self._bar_creep_done_phase = True
                if self._timer is not None:
                    try:
                        self._timer.stop()
                    except Exception:
                        pass
                self._pending_show_done_dialog = bool(d.get("show_done_dialog", False))
                self._pending_done_items = d.get("done_items") if isinstance(d.get("done_items"), list) else []
                self._pending_done_detail_text = str(d.get("done_detail_text", "") or "").strip() or None
                self._pending_output_dir = str(d.get("output_dir") or "").strip() or None
                try:
                    _ddm = d.get("done_delay_ms")
                    if _ddm is not None:
                        _close_ms = int(_ddm)
                    else:
                        _close_ms = int(getattr(self, "_done_delay_ms", 1000))
                except (TypeError, ValueError):
                    _close_ms = int(getattr(self, "_done_delay_ms", 1000))
                _close_ms = max(0, min(30000, _close_ms))
                if _log is not None:
                    try:
                        _n = len(self._pending_done_items)
                    except Exception:
                        _n = 0
                    _log.info(
                        "[CSV_MG] ProgressDialog DONE seq=%s show_done_dialog=%s done_items=%s done_delay_ms=%s",
                        seq,
                        self._pending_show_done_dialog,
                        _n,
                        _close_ms,
                    )
                total = d.get("total")
                done = d.get("done", total)
                _done_lbl = str(d.get("phase", "") or "").strip() or "完了"
                prev_bar = int(self._bar.value())
                creep = int(getattr(self, "_progress_bar_creep_pct", 0) or 0)
                self._progress_display_target = 100
                try:
                    poll_iv = int(self._bar_creep_timer.interval() or 40)
                except Exception:
                    poll_iv = 40
                done_creep = compute_done_finish_creep_pct(prev_bar, creep, poll_iv)
                self._progress_bar_done_creep_pct = done_creep
                _close_ms = compute_done_close_delay_ms(
                    prev_bar, creep, poll_iv, _close_ms, done_creep=done_creep
                )
                if creep <= 0 or prev_bar >= 100:
                    self._bar.setValue(100)
                    self._label_file.setText(_done_lbl)
                    self._pending_done_label = None
                else:
                    self._pending_done_label = _done_lbl
                    try:
                        self._label_file.setText(DONE_FINISH_INTERIM_LABEL)
                    except Exception:
                        pass
                if done is not None and total is not None:
                    self._label_count.setText(f"{done} / {total}")
                # ラベル更新後のレイアウト確定を待ち、前面化＋中央を 1 回だけ（直後の二重センタ廃止）
                if (
                    getattr(self, "_center_on_parent_widget", False)
                    and self.parentWidget() is not None
                ) or self._parent_hwnd or getattr(self, "_excel_rect", None):
                    QTimer.singleShot(16, lambda: _progress_done_recenter(self))
                # 完了表示を短時間維持してから _close_after_done（req の done_delay_ms、または DONE pickle の done_delay_ms）
                QTimer.singleShot(_close_ms, self._close_after_done)
                return

            if status_u == "RUN" and _log is not None:
                try:
                    import time as _t
                    _now = _t.time()
                    _last_t = float(getattr(self, "_run_info_last_t", 0.0) or 0.0)
                    _last_seq = int(getattr(self, "_run_info_last_seq", -1) or -1)
                    if seq != _last_seq and (_now - _last_t) >= 1.0:
                        self._run_info_last_t = _now
                        self._run_info_last_seq = seq
                        _log.info(
                            "[CSV_MG] ProgressDialog RUN seq=%s pct=%s phase=%s detail=%s done=%s total=%s",
                            seq,
                            d.get("pct"),
                            d.get("phase"),
                            d.get("detail"),
                            d.get("done"),
                            d.get("total"),
                        )
                except Exception:
                    pass

            # 判定: 行数超過時は Excel 操作を有効化し、警告メッセージ表示。OK で閉じたあと return_merge を投入し結合画面を再表示
            if status == "OVER_LIMIT":
                self._stop_progress_timers()
                self._teardown_progress_shared_state(
                    excel_unlock=bool(self._parent_hwnd),
                )
                msg = _normalize_message_newlines(str(d.get("msg", "") or "Excelの最大行数を超えるため結合を中止しました。").strip())
                try:
                    from ui_qt.ui_notification_sound import play_notification_sound

                    play_notification_sound("info")
                except Exception:
                    pass
                mb = QMessageBox(self)
                mb.setIcon(QMessageBox.Icon.Warning)
                mb.setWindowTitle("行数超過")
                mb.setText(msg)
                mb.setStandardButtons(QMessageBox.StandardButton.Ok)
                try:
                    fl = mb.windowFlags()
                    fl &= ~Qt.WindowType.WindowMinMaxButtonsHint
                    mb.setWindowFlags(fl)
                except Exception:
                    pass
                mb.exec()
                # 結合画面を再表示するため、return_merge リクエストを Pickle で投入
                try:
                    import time as _time
                    req_dir = d.get("request_dir") or ""
                    payload = d.get("return_merge_payload")
                    if req_dir and isinstance(payload, dict):
                        req_path = Path(req_dir) / f"req_return_merge_{int(_time.time()*1000)}_{os.getpid()}.pkl"
                        ipc_file.write_pickle(req_path, payload)
                except Exception:
                    pass
                self.accept()
                try:
                    _close_all_modeless()
                except Exception:
                    pass
                return

            if status_u == "ERROR":
                self._stop_progress_timers()
                detail = str(d.get("detail", "") or d.get("msg", "") or "").strip()
                self._teardown_progress_shared_state(
                    excel_unlock=bool(self._parent_hwnd),
                )
                try:
                    from ui_qt.ui_notification_sound import play_notification_sound

                    play_notification_sound("error")
                except Exception:
                    pass
                mb = QMessageBox(self)
                mb.setIcon(QMessageBox.Icon.Critical)
                mb.setWindowTitle("エラー")
                mb.setText(
                    _normalize_message_newlines(
                        detail or "結合処理中にエラーが発生しました。"
                    )
                )
                mb.setStandardButtons(QMessageBox.StandardButton.Ok)
                try:
                    fl = mb.windowFlags()
                    fl &= ~Qt.WindowType.WindowMinMaxButtonsHint
                    mb.setWindowFlags(fl)
                except Exception:
                    pass
                mb.exec()
                try:
                    self.accept()
                except Exception:
                    pass
                try:
                    _close_all_modeless()
                except Exception:
                    pass
                return

            if status_u == "CANCEL":
                self._stop_progress_timers()
                pw_cancel = self._req.get("partner_widget_after_cancel")
                self._teardown_progress_shared_state(
                    excel_unlock=bool(int(self._parent_hwnd or 0)),
                )
                try:
                    _tm_active = self._timer is not None and self._timer.isActive()
                except Exception:
                    _tm_active = False
                try:
                    _sz = self._progress_path.stat().st_size if self._progress_path.exists() else 0
                except Exception:
                    _sz = -1
                _msg = (
                    "[UI_PROGRESS_DIAG] CANCEL pickle seq=%s last_seen_seq=%s path=%s size=%s "
                    "parent_hwnd=%s excel_lock=%s partner_after_cancel=%s timer_was_active=%s"
                )
                _args = (
                    seq,
                    getattr(self, "_last_seen_seq", -1),
                    str(self._progress_path),
                    _sz,
                    int(self._parent_hwnd or 0),
                    bool(getattr(self, "_excel_lock", False)),
                    pw_cancel is not None,
                    _tm_active,
                )
                if _log is not None:
                    try:
                        _log.info(_msg, *_args)
                    except Exception:
                        pass
                if _diag_ui is not None:
                    try:
                        _diag_ui.info(_msg, *_args)
                    except Exception:
                        pass
                try:
                    self.accept()
                except Exception as _a_exc:
                    if _log is not None:
                        try:
                            _log.warning("[UI_PROGRESS_DIAG] progress accept() on CANCEL: %s", _a_exc)
                        except Exception:
                            pass
                try:
                    _close_all_modeless()
                except Exception:
                    pass
                # csv_sp: 重複確認でキャンセルしたあと分割画面を再表示（破棄済みなら IPC で開き直し）
                try:
                    pw = pw_cancel
                    tmpl = self._req.get("partner_csv_sp_reopen_template")

                    def _show_partner_cancel() -> None:
                        try:
                            from shiboken6 import Shiboken as _Shiboken  # type: ignore
                        except Exception:
                            _Shiboken = None  # type: ignore
                        _alive = pw is not None
                        if _alive and _Shiboken is not None:
                            try:
                                _alive = bool(_Shiboken.isValid(pw))
                            except Exception:
                                _alive = False
                        if _alive:
                            try:
                                _attr = getattr(Qt.WidgetAttribute, "WA_DontShowOnScreen", None) or getattr(
                                    Qt, "WA_DontShowOnScreen", None
                                )
                                if _attr is not None:
                                    pw.setAttribute(_attr, False)
                            except Exception as _e1:
                                if _log is not None:
                                    try:
                                        _log.warning("[UI_PROGRESS_DIAG] partner WA_DontShowOff: %s", _e1)
                                    except Exception:
                                        pass
                            try:
                                pw.show()
                                pw.raise_()
                                pw.activateWindow()
                                try:
                                    _clr = getattr(pw, "clear_sp_progress_partner_phase", None)
                                    if callable(_clr):
                                        _clr()
                                except Exception:
                                    pass
                                if _diag_ui is not None:
                                    try:
                                        _diag_ui.info(
                                            "[UI_PROGRESS_DIAG] partner show() ok type=%s",
                                            type(pw).__name__,
                                        )
                                    except Exception:
                                        pass
                                return
                            except Exception as _e2:
                                if _log is not None:
                                    try:
                                        _log.warning("[UI_PROGRESS_DIAG] partner show() failed: %s", _e2)
                                    except Exception:
                                        pass
                                if _diag_ui is not None:
                                    try:
                                        _diag_ui.warning(
                                            "[UI_PROGRESS_DIAG] partner show() failed: %s", _e2
                                        )
                                    except Exception:
                                        pass
                        if isinstance(tmpl, dict) and int(self._parent_hwnd or 0):
                            try:
                                from ui_qt.ui_csv_sp import queue_csv_sp_split_reopen_request

                                queue_csv_sp_split_reopen_request(tmpl)
                                if _diag_ui is not None:
                                    try:
                                        _diag_ui.info(
                                            "[UI_PROGRESS_DIAG] partner reopen queued (ipc) sheet_id=%s",
                                            str(tmpl.get("sheet_id") or ""),
                                        )
                                    except Exception:
                                        pass
                            except Exception as _e4:
                                if _log is not None:
                                    try:
                                        _log.warning("[UI_PROGRESS_DIAG] partner reopen ipc failed: %s", _e4)
                                    except Exception:
                                        pass
                                if _diag_ui is not None:
                                    try:
                                        _diag_ui.warning(
                                            "[UI_PROGRESS_DIAG] partner reopen ipc failed: %s", _e4
                                        )
                                    except Exception:
                                        pass

                    if pw is not None or isinstance(tmpl, dict):
                        QTimer.singleShot(0, _show_partner_cancel)
                    else:
                        if _diag_ui is not None:
                            try:
                                _diag_ui.warning(
                                    "[UI_PROGRESS_DIAG] CANCEL but partner_widget_after_cancel is None"
                                )
                            except Exception:
                                pass
                except Exception as _e3:
                    if _log is not None:
                        try:
                            _log.warning("[UI_PROGRESS_DIAG] CANCEL partner branch: %s", _e3)
                        except Exception:
                            pass
                return

            # 通常の進捗更新: phase_i / phase / current_file / pct / done / total をラベル・バーに反映
            self._run_seen = True  # RUN を表示したので DONE を処理してよい
            try:
                if bool(d.get("hide_cancel_button")):
                    cr = getattr(self, "_cancel_row_widget", None)
                    if cr is not None:
                        cr.setVisible(False)
                    elif getattr(self, "_btn_cancel", None) is not None:
                        self._btn_cancel.setVisible(False)
                    # キャンセル行が無い進捗（hlclr 等）でも stretch ＋固定高さの余白を詰める
                    self._compact_progress_after_cancel_hidden()
            except Exception:
                pass
            _pt = d.get("phase_total")
            if _pt is not None:
                try:
                    _pti = int(_pt)
                    if _pti > 0:
                        self._phase_total = _pti
                except (TypeError, ValueError):
                    pass
            total = d.get("total")
            done = d.get("done")
            prev_bar = int(self._bar.value())
            if status_u == "RUN":
                raw_pct = d.get("pct")
                svc_pct_explicit = raw_pct is not None
                if svc_pct_explicit:
                    try:
                        pct = max(0, min(99, int(raw_pct)))
                    except (TypeError, ValueError):
                        svc_pct_explicit = False
                if not svc_pct_explicit:
                    try:
                        to_i = int(total) if total is not None else 0
                        dn_i = int(done) if done is not None else None
                        if to_i > 0 and dn_i is not None:
                            pct = max(0, min(99, int(dn_i * 100 / to_i)))
                    except (TypeError, ValueError):
                        pct = 0
                self._progress_display_target = max(
                    int(getattr(self, "_progress_display_target", 0) or 0),
                    int(pct),
                )
            else:
                pct = int(d.get("pct", 0) or 0)
                pct = 0 if pct < 0 else 100 if pct > 100 else pct
            phase_i = int(d.get("phase_i", 0) or 0)
            self._last_run_phase_i = phase_i
            if getattr(self, "_center_on_parent_widget", False):
                dn_i = int(done) if done is not None else None
                to_i = int(total) if total is not None else None
                if phase_i <= 1 and (dn_i is None or dn_i <= 1):
                    self._nm_pi_disp = 0
                    self._nm_done_disp = 0
                    self._nm_tot_disp = 0
                pt_cur = int(getattr(self, "_phase_total", 0) or 0)
                self._nm_tot_disp = max(self._nm_tot_disp, to_i or 0, pt_cur)
                self._nm_pi_disp = max(self._nm_pi_disp, phase_i)
                phase_i = self._nm_pi_disp
                if dn_i is not None:
                    self._nm_done_disp = max(self._nm_done_disp, dn_i)
                    done = self._nm_done_disp
                if to_i is not None:
                    td = max(to_i, self._nm_tot_disp)
                    total = td
                    self._nm_tot_disp = td
                self._phase_total = max(int(getattr(self, "_phase_total", 0) or 0), self._nm_tot_disp)
            phase = str(d.get("phase", "") or "").strip()
            current_file = str(d.get("current_file", "") or "").strip()
            msg = str(d.get("msg", "") or "").strip()
            detail = str(d.get("detail", "") or "").strip()
            window_title = str(d.get("window_title", "") or "").strip()

            try:
                base = str(getattr(self, "_progress_base_title", "") or "").strip()
                if window_title:
                    self.setWindowTitle(
                        "%s — %s" % (base, window_title) if base else window_title
                    )
                elif base:
                    self.setWindowTitle(base)
            except Exception:
                pass

            head = _format_progress_status_text(
                phase=phase,
                msg=msg,
                detail=detail,
                current_file=current_file,
                window_title=bool(window_title),
            )
            self._label_file.setText(head)

            creep = int(getattr(self, "_progress_bar_creep_pct", 0) or 0)
            if status_u == "RUN":
                if creep <= 0:
                    self._bar.setValue(int(self._progress_display_target))
            else:
                self._bar.setValue(pct)

            # 進捗バー下: done / total を右寄せ（スラッシュ両側にスペース）
            if done is not None and total is not None:
                self._label_count.setText(f"{done} / {total}")
            elif total is not None:
                self._label_count.setText(f"0 / {total}")
            else:
                self._label_count.setText("0 / 0")
        except Exception:
            return

    def _write_progress_closed_ack(self) -> None:
        """進捗ダイアログのクローズ完了ACKを1回だけ書く。"""
        try:
            if bool(getattr(self, "_defer_closed_ack", False)):
                return
            if bool(getattr(self, "_progress_closed_ack_written", False)):
                return
            p = getattr(self, "_progress_closed_path", None)
            if p is None:
                return
            try:
                p.parent.mkdir(parents=True, exist_ok=True)
            except Exception:
                pass
            ipc_file.write_pickle(p, {"status": "CLOSED", "ts_ms": int(time.time() * 1000)})
            self._progress_closed_ack_written = True
        except Exception:
            pass

    def closeEvent(self, event) -> None:  # type: ignore[override]
        """
        【概要】
            × ボタン等で閉じた場合も、タイマー停止・Excel 操作有効化・モデルレス一覧からの削除・deleteLater を確実に行う。
        """
        try:
            event.accept()
        except Exception:
            pass
        try:
            self._write_progress_closed_ack()
        except Exception:
            pass
        try:
            self._stop_progress_timers()
            self._teardown_progress_shared_state()
        except Exception:
            pass
        try:
            self.deleteLater()
        except Exception:
            pass
        super().closeEvent(event)

    def _close_after_done(self) -> None:
        """
        【概要】
            DONE 表示後、1 秒遅延で呼ばれる。進捗画面を先に閉じ、show_done_dialog が True かつ items/detail_text があれば完了通知（DoneDialog）を表示する。
        【補足】
            req に partner_widget_after_done（QWidget）があれば、DoneDialog の exec 終了後（または通知なしの完了時）に次イベントで accept/close して呼び出し側の exec を解放する（csv_sp の分割画面など）。
        """
        try:
            self._apply_pending_done_label()
            try:
                if int(self._bar.value()) < 100:
                    self._bar.setValue(100)
            except Exception:
                pass
            if _log is not None:
                import time as _t
                _log.info(
                    "[CSV_MG] ProgressDialog _close_after_done (closing progress) t=%.3f",
                    _t.time(),
                )
            items = self._pending_done_items if isinstance(getattr(self, "_pending_done_items", None), list) else []
            detail_text = (getattr(self, "_pending_done_detail_text", None) or "").strip()
            show_done = bool(getattr(self, "_pending_show_done_dialog", False)) and (
                bool(items) or bool(detail_text)
            )
            done_cfg = getattr(self, "_done_cfg", None) or {}
            ph = int(self._parent_hwnd or 0)
            excel_rect = getattr(self, "_excel_rect", None)
            self._defer_closed_ack = True

            # 1) 進捗画面を先に閉じる（Excel COM より先に UI を解放）
            self._stop_progress_timers()
            try:
                self.hide()
            except Exception:
                pass
            try:
                if bool(self._req.get("close_parent_when_done", True)):
                    p = self.parent()
                    if p is not None and hasattr(p, "close"):
                        p.close()
            except Exception:
                pass
            try:
                self.close()
            except Exception:
                pass
            try:
                self.deleteLater()
            except Exception:
                pass
            self._teardown_progress_shared_state(cursor_off=False)
            self._progress_wait_cursor_off()

            # 2) 完了通知を表示（前面化は DoneDialog.showEvent に一本化）
            if show_done and (items or detail_text):
                try:
                    req = {"items": items}
                    if detail_text:
                        req["detail_text"] = detail_text
                    _od = getattr(self, "_pending_output_dir", None)
                    if isinstance(_od, str) and _od.strip():
                        req["output_dir"] = _od.strip()
                    if excel_rect is not None:
                        req["excel_rect"] = list(excel_rect)
                    dlg = create_done_dialog(req, ph, None, done_cfg)
                    pw_after = self._req.get("partner_widget_after_done")
                    if pw_after is not None:
                        dlg.show()
                        dlg.exec()
                    else:
                        try:
                            from ui_qt.ui_common import _keep_modeless, _remove_from_modeless

                            def _on_done_dialog_finished(_rc: int = 0) -> None:
                                _remove_from_modeless(dlg)

                            try:
                                dlg.finished.connect(_on_done_dialog_finished)
                            except Exception:
                                pass
                            _keep_modeless(dlg)
                            dlg.show()
                        except Exception:
                            try:
                                dlg.show()
                                dlg.exec()
                            except Exception:
                                pass
                except Exception:
                    pass
            try:
                self._defer_closed_ack = False
                self._write_progress_closed_ack()
            except Exception:
                pass
            try:
                self._defer_closed_ack = False
            except Exception:
                pass
            # 3) csv_sp 等: 分割画面の exec を終了。次イベントで accept/close し、_close_after_done 内からの exec 再入を避ける
            try:
                pw = self._req.get("partner_widget_after_done")
                if pw is not None:

                    def _finish_partner() -> None:
                        try:
                            _clr = getattr(pw, "clear_sp_progress_partner_phase", None)
                            if callable(_clr):
                                _clr()
                        except Exception:
                            pass
                        try:
                            if hasattr(pw, "accept"):
                                pw.accept()
                        except Exception:
                            pass
                        try:
                            pw.close()
                        except Exception:
                            pass

                    QTimer.singleShot(0, _finish_partner)
            except Exception:
                pass
        except Exception:
            pass


def raise_csv_sp_partner_progress(parent_hwnd: int) -> None:
    """
    csv_sp のモデルレス進捗（分割パートナー付き）を前面化する。
    重複確認モーダルを閉じる直前に呼び、背後に残った進捗と枠の重なり・ゴースト感を抑える。
    """
    app = QApplication.instance()
    if app is None:
        return
    ph = int(parent_hwnd or 0)
    for w in app.topLevelWidgets():
        if not isinstance(w, ProgressDialog):
            continue
        if not w.isVisible():
            continue
        req = getattr(w, "_req", None)
        if not isinstance(req, dict):
            continue
        if not (req.get("partner_widget_after_done") or req.get("partner_widget_after_cancel")):
            continue
        try:
            w.raise_()
            w.activateWindow()
            if ph:
                from ui_qt.ui_common import ensure_dialog_front_of_excel

                ensure_dialog_front_of_excel(w, ph, None)
        except Exception:
            pass


def create_progress_dialog(
    req_dict: dict,
    parent_hwnd: int,
    parent_widget: Optional[QWidget] = None,
    progress_cfg: Optional[dict] = None,
) -> ProgressDialog:
    """
    【概要】
        進捗ダイアログを生成し、サイズ調整後にモデルレス保護へ登録して返す。
    【補足】
        中央・オーナー・前面化は showEvent 直後 1 回に集約（show 前の二重配置とちらつきを避ける）。
        no_native_window 時も opacity は 1 のまま（透明化による点滅を避ける）。
    """
    dlg = ProgressDialog(req_dict, int(parent_hwnd or 0), parent_widget, progress_cfg=progress_cfg)
    _cfg = progress_cfg if isinstance(progress_cfg, dict) else _get_progress_config()
    win_cfg = (_cfg or {}).get("WINDOW") or {}
    if not bool(req_dict.get("no_native_window", False)):
        try:
            dlg.adjustSize()
        except Exception:
            pass
        if not win_cfg.get("SHOW_MINIMIZE", False) and not win_cfg.get("SHOW_MAXIMIZE", False):
            try:
                hwnd = int(dlg.winId()) if hasattr(dlg, "winId") else 0
                if hwnd and _w32 is not None and hasattr(_w32, "set_window_style_remove_min_max"):
                    _w32.set_window_style_remove_min_max(hwnd)
            except Exception:
                pass
        try:
            dlg.setWindowOpacity(1.0)
        except Exception:
            pass
    else:
        try:
            dlg.setWindowOpacity(1.0)
        except Exception:
            pass
    _keep_modeless(dlg)
    return dlg


# 公開シンボル（他モジュールから from ui_dialog_progress import ProgressDialog 等で参照可能）
__all__ = ["ProgressDialog", "create_progress_dialog", "raise_csv_sp_partner_progress"]
