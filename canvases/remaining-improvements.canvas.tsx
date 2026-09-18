import {
  Callout,
  Divider,
  Grid,
  H1,
  H2,
  Stack,
  Stat,
  Table,
  Text,
} from "cursor/canvas";

export default function RemainingImprovements() {
  return (
    <Stack gap={20}>
      <H1>残件</H1>
      <Text tone="secondary">
        2026-09-12 この段はコードとテストまで。実機は未確認。ビルドは不要。
      </Text>

      <Grid columns={3} gap={12}>
        <Stat value="実機待ち" label="この段（#9 連携先集約 / ファイル通過）" tone="warning" />
        <Stat value="見送り" label="#10 追加ストリーム化" />
        <Stat value="終了後" label="進捗文言テスト 3件" tone="info" />
      </Grid>

      <H2>この段で動かしたもの</H2>
      <Table
        striped
        headers={["対象", "中身", "触っていないもの"]}
        rows={[
          [
            "#9",
            "結合の有無、連携先の集約、ファイル通過判定を source_ui へ移した。判定は同じ",
            "xlsx キャッシュ、抽出の照合、COM、Nuitka",
          ],
        ]}
      />

      <H2>まだやる</H2>
      <Table
        striped
        headers={["残", "中身", "注意"]}
        rows={[
          [
            "#9 追加分割",
            "モノリス本体は残",
            "xlsx キャッシュは動かさない",
          ],
          [
            "#12 COM",
            "Excel 本体の広域 except",
            "触ると、今まで完走していた一括が途中で止まり得る",
          ],
        ]}
      />

      <H2>見送り</H2>
      <Table
        striped
        headers={["対象", "理由"]}
        rows={[
          [
            "#10 追加ストリーム化",
            "大シートのメモリ削減。多ファイル小容量では効果が薄い",
          ],
        ]}
      />

      <H2>本キャンペーン終了後</H2>
      <Table
        striped
        headers={["対象", "直し方"]}
        rows={[
          [
            "test_data_agg_debug_step_snapshots.py の3件",
            "画面を古い文言に戻さない。テストを今の進捗文言に合わせる",
          ],
        ]}
      />

      <Callout tone="info" title="実機の見方（ビルド不要）">
        いつもの2シナリオが前回と同じ行数・列数で完走すること。バージョン確認は今回開かなくてよい。
      </Callout>

      <Divider />
      <Text size="small" tone="secondary">
        Source: codebase improvement review · テスト 2026-09-12。この段の実機ログは未取得。
      </Text>
    </Stack>
  );
}
