import {
  Callout,
  Divider,
  H1,
  H2,
  Stack,
  Table,
  Text,
} from "cursor/canvas";

export default function Phase1822Report() {
  return (
    <Stack gap={20}>
      <H1>#18 / #22 フェーズ報告</H1>
      <Text tone="secondary">
        2026-09-12 · ステップバイステップ実施 · polars は速度経路のため維持
      </Text>

      <Callout tone="success" title="総合">
        #18（依存宣言整理）と #22（衛星テスト統合）を完了。本番ホットパス・Nuitka
        同梱ロジックは未変更。関連テスト 60 passed。公開 API ゲート PASS。
      </Callout>

      <H2>#18 依存の肥大（lock と配布）</H2>
      <Table
        striped
        headers={["段", "修正要点", "テスト結果"]}
        rows={[
          [
            "18-1 棚卸し",
            "lock のみ宣言。runtime / optional(polars) / dev / build を分類。配布は Nuitka import グラフ。ZIP 主因は PySide6+pandas（+同梱時 polars）。debugpy/ruff/PyInstaller はアプリ未参照。",
            "調査のみ（変更なし）",
          ],
          [
            "18-2 方針",
            "開発専用を runtime から外す。polars は optional で KEEP。bsdiff4 を runtime に正式記載。数値スタック削除・svc nofollow 強化はしない。",
            "方針確定（実装は 18-3）",
          ],
          [
            "18-3 反映",
            "requirements-runtime / optional / dev / build 新設。lock は -r 参照。Exe化ドキュメント追記。test_requirements_layout.py 追加。",
            "test_requirements_layout.py + polars helpers → PASS",
          ],
        ]}
      />

      <H2>#22 テスト肥大・重複整理</H2>
      <Table
        striped
        headers={["段", "修正要点", "テスト結果"]}
        rows={[
          [
            "22-1 棚卸し",
            "安全候補: host_restore 双子 / read_cap / join_seed / master_debug_cancel / extract_repeat_limit。巨大 join_priority・debug ゲートは触らない。",
            "調査のみ",
          ],
          [
            "22-2a",
            "excel_host_restore_after_operation → test_excel_host_restore.py に統合し衛星削除",
            "test_excel_host_restore.py PASS",
          ],
          [
            "22-2b/c",
            "read_cap・join_seed を join_priority 末尾へ移し衛星削除（断言は全維持）",
            "移設 4 本 PASS",
          ],
          [
            "22-2d",
            "master_debug_cancel → batch_cancel 末尾へ移し衛星削除",
            "移設 3 本 + batch_cancel 全体 PASS",
          ],
          [
            "22-2e",
            "extract_repeat_limit → extract_limit へ統合し衛星削除",
            "test_data_agg_extract_limit.py PASS",
          ],
        ]}
      />

      <H2>ゲート・関連スイート</H2>
      <Table
        striped
        headers={["スイート", "結果", "備考"]}
        rows={[
          [
            "今回変更クラスタ一式（60）",
            "PASS",
            "requirements / host_restore / extract_limit / batch_cancel / 移設4本 / public_api / polars",
          ],
          [
            "test_ui_data_agg_debug_public_api",
            "PASS",
            "フェーズ完了ゲート",
          ],
          [
            "既知・本変更外の失敗",
            "対象外",
            "join_priority::table_row_file_paths_groups_by_device_id / debug_step_snapshots 進捗文言・file_progress（既存乖離の可能性）",
          ],
        ]}
      />

      <Divider />
      <Text size="small" tone="secondary">
        速度波及: 実行コード未変更（宣言とテスト配置のみ）。polars 経路は維持。CI
        時間短縮はファイル統合中心で小規模。
      </Text>
    </Stack>
  );
}
