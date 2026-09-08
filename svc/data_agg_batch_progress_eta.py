# -*- coding: utf-8 -*-
"""本番一括進捗: 時間寄りバー％（A）と右下作業件数（B）の算出。

2 行フェーズ表示は変更しない。
- バー％: ファイル処理の時間寄り推定（マウント件数には連動しない）
- 右下 N/M: N=u+f, M=U+F（UNC マウント件数 + 参照ファイル総数）
"""
from __future__ import annotations

import time
from dataclasses import dataclass, field
from typing import Sequence


FILE_BAND_MAX = 90  # ファイル処理帯の上限。書込みは 93〜100 を既存経路が担当。


def count_network_paths(paths: Sequence[object]) -> int:
    """UNC / マップドドライブなど、マウント対象になり得るパス数。"""
    try:
        from svc.data_agg_path_network import path_is_network
    except Exception:
        return 0
    n = 0
    for p in paths:
        try:
            if path_is_network(p):
                n += 1
        except Exception:
            continue
    return n


@dataclass
class BatchProgressWorkUnits:
    """
    右下 N/M 用。マウント完了 u/U とファイル処理完了 f/F を別管理し、
    N=u+f / M=U+F だけを公開する（片方で上書きしない）。
    """

    mount_done: int = 0
    mount_total: int = 0
    files_done: int = 0
    files_total: int = 0

    def reset_counts(self) -> None:
        self.mount_done = 0
        self.mount_total = 0
        self.files_done = 0
        self.files_total = 0

    def set_files_total(self, n: int, *, reset_done: bool = False) -> None:
        self.files_total = max(0, int(n))
        if reset_done:
            self.files_done = 0
        elif self.files_total > 0:
            self.files_done = min(self.files_done, self.files_total)

    def set_mount_total(self, n: int, *, reset_done: bool = False) -> None:
        self.mount_total = max(0, int(n))
        if reset_done:
            self.mount_done = 0
        elif self.mount_total > 0:
            self.mount_done = min(self.mount_done, self.mount_total)

    def note_mount_progress(self, done_i: int, total_n: int) -> None:
        # ステージ側の実ジョブ数を正とする（事前推定より優先）
        tn = max(0, int(total_n))
        if tn > 0:
            self.mount_total = tn
        u = max(0, int(done_i))
        if self.mount_total > 0:
            u = min(u, self.mount_total)
        self.mount_done = max(self.mount_done, u)

    def note_files_progress(self, done_f: int, total_f: int) -> None:
        ft = max(0, int(total_f))
        if ft > self.files_total:
            self.files_total = ft
        d = max(0, int(done_f))
        if self.files_total > 0:
            d = min(d, self.files_total)
        self.files_done = max(self.files_done, d)

    def nm(self) -> tuple[int, int]:
        """(N, M) = (u+f, max(1, U+F))."""
        u = max(0, int(self.mount_done))
        U = max(0, int(self.mount_total))
        f = max(0, int(self.files_done))
        F = max(0, int(self.files_total))
        if U > 0:
            u = min(u, U)
        if F > 0:
            f = min(f, F)
        M = max(1, U + F)
        N = min(M, u + f)
        return N, M


@dataclass
class BatchProgressEta:
    """実行中の平均ファイル時間からバー％を推定する（単調・ある程度時間均等）。"""

    file_band_max: int = FILE_BAND_MAX
    _t0: float = field(default_factory=time.perf_counter)
    _last_pct: int = 0
    _max_done: int = 0

    def reset(self) -> None:
        self._t0 = time.perf_counter()
        self._last_pct = 0
        self._max_done = 0

    def elapsed(self) -> float:
        return max(0.0, time.perf_counter() - self._t0)

    def update_pct(
        self,
        *,
        n_files: int,
        files_done: int,
        sub: int | None = None,
    ) -> int:
        """
        ファイル処理帯の pct（0〜file_band_max）を返す。単調増加。

        files_done: 完了とみなす件数（0..N）。並列時は done_n、逐次は処理中番号など。
        マウント件数は含めない（バーは時間寄り・ファイル処理基準）。
        """
        n = max(1, int(n_files or 1))
        d = max(0, min(n, int(files_done or 0)))
        if d > self._max_done:
            self._max_done = d
        d = self._max_done
        elapsed = self.elapsed()
        band = max(1, int(self.file_band_max))

        si = int(sub) if sub is not None else 0
        if si >= 7 or d >= n:
            pct = band
        elif d <= 0:
            # 1 件目完了前: 最初の枠の半分までゆっくり
            first = band / float(n)
            soft = min(first * 0.55, (elapsed / 8.0) * first)
            pct = int(max(0.0, soft))
        else:
            avg = elapsed / float(d)
            est_total = max(avg * float(n), 1e-6)
            time_frac = min(1.0, elapsed / est_total)
            file_frac = float(d) / float(n)
            # サンプルが増えるほど時間寄り
            w_time = min(0.85, 0.40 + 0.08 * float(d))
            frac = w_time * time_frac + (1.0 - w_time) * file_frac
            pct = int(frac * band)
            # 完了件数からの床（遅れすぎ防止）
            floor = int(file_frac * band * 0.80)
            pct = max(pct, floor)

        pct = max(int(self._last_pct), min(band, int(pct)))
        self._last_pct = pct
        return pct


def resolve_batch_progress_file_counts(
    *,
    sub: int,
    n_files: int,
    file_index: int | None,
    done_n: int | None,
) -> tuple[int, int]:
    """
    ファイル処理側の完了件数 f / 総数 F（マウントは含まない）。

    2 行表示用の phase_nm とは独立。sub>=7 は全ファイル完了扱い。
    """
    n = max(0, int(n_files or 0))
    if n <= 0:
        n = 1
    si = int(sub)
    if si >= 7:
        return n, n
    if done_n is not None:
        try:
            d = max(0, min(n, int(done_n)))
            return d, n
        except (TypeError, ValueError):
            pass
    if file_index is not None:
        try:
            d = max(0, min(n, int(file_index)))
            return d, n
        except (TypeError, ValueError):
            pass
    return 0, n


def apply_batch_hook_progress_metrics(
    eta: BatchProgressEta,
    *,
    sub: int,
    n_files: int,
    file_index: int | None,
    done_n: int | None,
    prev_pct: int,
    work: BatchProgressWorkUnits | None = None,
) -> tuple[int, int, int]:
    """
    hook 1 回分の (pct, done, total) を返す。
    pct はファイル処理の時間寄り（単調）。
    done/total は work があれば N=u+f / M=U+F、なければ f/F。
    """
    f_done, f_total = resolve_batch_progress_file_counts(
        sub=sub,
        n_files=n_files,
        file_index=file_index,
        done_n=done_n,
    )
    if work is not None:
        work.note_files_progress(f_done, f_total)
        pct = eta.update_pct(
            n_files=max(1, work.files_total or f_total),
            files_done=work.files_done,
            sub=sub,
        )
        done, total = work.nm()
    else:
        pct = eta.update_pct(n_files=f_total, files_done=f_done, sub=sub)
        done, total = f_done, f_total
    pct = max(int(prev_pct), min(int(eta.file_band_max), int(pct)))
    return pct, done, total
