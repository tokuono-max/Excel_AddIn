# -*- coding: utf-8 -*-
"""Export codebase improvement review to Markdown / HTML for printing."""
from __future__ import annotations

import html
from pathlib import Path

FINDINGS = [
    {
        "id": 1,
        "priority": "高",
        "score": 7,
        "items": "正確性",
        "feature": "データ集約・前段行マージ（結合代入の前処理）",
        "title": "行マージと結合比較の正規化の細かい不一致",
        "impact": "大半は正確。特定条件でのみ行の横結合がずれる",
        "risk": "同一反復内で '001 と 001（や数値と文字列）が混在すると、前段マージだけ別扱いになることあり",
        "freq": "低〜中（キー表記ゆれが混在したときのみ）",
        "tradeoff": "マージも join_compare に揃えると該当シナリオの行のまとまり方が変わる。反復またぎは引き続き統合しない",
        "difficulty": "低",
    },
    {
        "id": 2,
        "priority": "すぐに",
        "score": 10,
        "items": "正確性",
        "feature": "データ集約・抽出（Excel読取）",
        "title": "抽出失敗の except→None/pass が空セルに見える",
        "impact": "原因不明の空欄・結合不一致",
        "risk": "失敗が空欄に化け、結合失敗や誤った完成表を気付かず確定",
        "freq": "中（破損・書式異常・巨大シート時）",
        "tradeoff": "失敗をエラー停止にすると一括が途中終了。継続＋明示エラー印の両立が必要",
        "difficulty": "中",
    },
    {
        "id": 3,
        "priority": "すぐに",
        "score": 9,
        "items": "正確性",
        "feature": "データ集約・シート指定読取",
        "title": "シート名不一致時の active フォールバック",
        "impact": "別シートを黙って読む",
        "risk": "指定外シートのデータを集約し、正しそうに見える誤結果を出力",
        "freq": "中（シート名変更・typo時は確実）",
        "tradeoff": "エラー化すると旧シナリオ（シート名曖昧運用）が止まる。警告付きフォールバック移行も可",
        "difficulty": "低",
    },
    {
        "id": 4,
        "priority": "すぐに",
        "score": 9,
        "items": "ファイル容量",
        "feature": "開発環境（Nuitkaビルドログ）",
        "title": "logs/nuitka 約1.5GB の蓄積",
        "impact": "約1.5GB即時回収",
        "risk": "ディスク逼迫・バックアップ遅延・ビルド失敗",
        "freq": "高（ビルドするたび増加）",
        "tradeoff": "削除すると過去ビルドの原因調査が難しくなる。直近N世代保持が現実的",
        "difficulty": "低",
    },
    {
        "id": 5,
        "priority": "すぐに",
        "score": 9,
        "items": "ファイル容量",
        "feature": "開発環境（配布ビルド成果物）",
        "title": "dist/ 旧リリース＋重複ビルド約3.1GB",
        "impact": "2GB超削減見込み",
        "risk": "ディスク逼迫・誤配布・旧版混同",
        "freq": "高（リリース積み上げのたび）",
        "tradeoff": "旧版ローカル比較ができなくなる。必要版はアーカイブ置き場へ退避",
        "difficulty": "低",
    },
    {
        "id": 6,
        "priority": "すぐに",
        "score": 8,
        "items": "ファイル容量 / スリム化",
        "feature": "リポジトリ衛生（一時ファイル・重複テスト）",
        "title": "ルート未追跡ゴミ (_tmp_* / .coverage / 重複テスト)",
        "impact": "誤コミット防止＋pytestノイズ排除",
        "risk": "ゴミの誤コミット・テスト収集失敗・レビューノイズ",
        "freq": "中（作業後に残存しやすい）",
        "tradeoff": "調査用スクリプトも消える。再現手順が必要なら docs/tools へ正規化",
        "difficulty": "低",
    },
    {
        "id": 7,
        "priority": "高",
        "score": 8,
        "items": "正確性",
        "feature": "データ集約・結合前加工（チェックラベル）",
        "title": "加工チェックの部分文字列マッチ",
        "impact": "ラベル変更で誤適用",
        "risk": "意図しないトリム／日付変換が走り、キー不一致や値破壊",
        "freq": "中（ラベル追加・文言変更時）",
        "tradeoff": "厳密一致は表記ゆれに弱くなる。ID化すると設定移行コストが発生",
        "difficulty": "低",
    },
    {
        "id": 8,
        "priority": "高",
        "score": 8,
        "items": "正確性 / スリム化",
        "feature": "基盤層（core / svc 境界）",
        "title": "core → svc 循環依存",
        "impact": "初期化順・Nuitka・単体テストの脆さ",
        "risk": "起動失敗・バンドル抜け・テストが壊れて回帰見逃し",
        "freq": "低〜中（リファクタ・配布ビルド時）",
        "tradeoff": "層分割は広い import 差し替えが必要。短期は回帰リスクが上がる",
        "difficulty": "中",
    },
    {
        "id": 9,
        "priority": "高",
        "score": 8,
        "items": "速度 / スリム化",
        "feature": "データ集約コア／デバッグUI全体",
        "title": "巨大モノリス分割 (集約コア＋デバッグUI)",
        "impact": "変更リスク局所化",
        "risk": "修正の副作用が広がり、回帰バグ・レビュー不能が常態化",
        "freq": "高（機能追加のたび）",
        "tradeoff": "分割作業自体が大規模・長期。途中はマージ衝突と二重メンテが増える",
        "difficulty": "高",
    },
    {
        "id": 10,
        "priority": "高",
        "score": 7,
        "items": "速度",
        "feature": "データ集約・Excel抽出（大シート）",
        "title": "シート全行 materialize のメモリ/I/O",
        "impact": "大シートで数百MB〜数GB・数秒〜数十秒",
        "risk": "メモリ逼迫・ハング・一括処理のタイムアウト",
        "freq": "高（大容量Excel利用時は毎回）",
        "tradeoff": "範囲限定・ストリーム化は実装複雑。小シートでは全読込の方が単純で速い場合あり",
        "difficulty": "高",
    },
    {
        "id": 11,
        "priority": "高",
        "score": 7,
        "items": "速度",
        "feature": "データ集約・非表示行除外",
        "title": "skip_hidden_rows 時の非 read_only 再オープン",
        "impact": "読込約2倍",
        "risk": "一括処理が大幅遅延・体感タイムアウト",
        "freq": "高（非表示除外ON時は毎回）",
        "tradeoff": "read_only維持は非表示判定が不完全になり得る。正確性優先なら再オープンが必要",
        "difficulty": "中",
    },
    {
        "id": 12,
        "priority": "高",
        "score": 7,
        "items": "正確性 / スリム化",
        "feature": "データ集約／UI／更新（横断）",
        "title": "広域 except Exception の縮減",
        "impact": "障害可視化・再発防止",
        "risk": "本番障害の原因特定不能・誤った空結果の継続",
        "freq": "高（例外経路は日常的）",
        "tradeoff": "例外を厳格化するとCOM/Excelの一時失敗で処理が止まりやすくなる",
        "difficulty": "中",
    },
    {
        "id": 13,
        "priority": "高",
        "score": 7,
        "items": "スリム化 / ファイル容量",
        "feature": "旧実装アーカイブ (svc/Old)",
        "title": "svc/Old/ 死コード (git追跡)",
        "impact": "認知負荷削減",
        "risk": "誤って旧APIを参照・検索ノイズ・レビュー誤誘導",
        "freq": "低（触ったとき）",
        "tradeoff": "git履歴・タグに残るが、ローカル参照が不便になる",
        "difficulty": "低",
    },
    {
        "id": 14,
        "priority": "高",
        "score": 7,
        "items": "スリム化 / 速度",
        "feature": "処理サーバ起動 (svc_server)",
        "title": "svc_server の DEBUG Popen モンキーパッチ常設",
        "impact": "本番ログノイズ除去",
        "risk": "ログ肥大・診断ノイズ・微小オーバーヘッド常時発生",
        "freq": "高（子プロセス起動のたび）",
        "tradeoff": "常時OFFだとスポーン調査が手間。環境変数でONにする形が無難",
        "difficulty": "低",
    },
    {
        "id": 15,
        "priority": "中",
        "score": 5,
        "items": "スリム化",
        "feature": "データ集約・polars任意読込",
        "title": "_get_polars 三重定義",
        "impact": "保守ずれ防止",
        "risk": "片方だけ修正され挙動分岐・再現困難なバグ",
        "freq": "低（修正時のみ）",
        "tradeoff": "共通化で循環importや遅延読込の設計が必要になる場合あり",
        "difficulty": "低",
    },
    {
        "id": 16,
        "priority": "中",
        "score": 5,
        "items": "スリム化",
        "feature": "日付変換 (YMD / HM)",
        "title": "svc_dt_ymd / svc_dt_hm のほぼ複製",
        "impact": "数百行削減",
        "risk": "片方だけ直して日付変換結果が機能間で不一致",
        "freq": "低〜中（日付機能改修時）",
        "tradeoff": "共通化で片方固有の仕様差を見失うリスク。差分を明示した共通化が必要",
        "difficulty": "中",
    },
    {
        "id": 17,
        "priority": "中",
        "score": 5,
        "items": "速度 / スリム化",
        "feature": "データ集約・シナリオ編集UI",
        "title": "UIシナリオ編集の過剰 deepcopy",
        "impact": "項目多いとUI鈍化",
        "risk": "編集操作のカクつき・入力遅延・誤操作誘発",
        "freq": "中（大規模シナリオ編集時）",
        "tradeoff": "浅いコピーは副作用で編集内容が汚染される危険。参照管理の設計が要る",
        "difficulty": "中",
    },
    {
        "id": 18,
        "priority": "中",
        "score": 5,
        "items": "ファイル容量 / スリム化",
        "feature": "配布パッケージ／依存管理",
        "title": "依存の肥大 (lockと配布の乖離)",
        "impact": "数十〜数百MB級の可能性",
        "risk": "配布ZIP肥大・導入時間増・不要DLL混入",
        "freq": "中（配布ビルドごと）",
        "tradeoff": "polars等を外すと大ファイル高速経路が弱まる。任意依存のままが妥当な場合あり",
        "difficulty": "中",
    },
    {
        "id": 19,
        "priority": "中",
        "score": 4,
        "items": "ファイル容量",
        "feature": "ドキュメント管理",
        "title": "ドキュメント版二重・docx/PDF重複",
        "impact": "数MB〜10MB",
        "risk": "版の取り違え・説明齟齬・容量浪費",
        "freq": "低（ドキュメント参照時）",
        "tradeoff": "docx削除は社内提出物ワークフローと衝突し得る。生成物置き場分離が安全",
        "difficulty": "低",
    },
    {
        "id": 20,
        "priority": "中",
        "score": 4,
        "items": "速度 / スリム化",
        "feature": "データ集約・結合書込ホットパス",
        "title": "結合ダンプ／診断コードの本番常駐",
        "impact": "ホットパス可読性向上",
        "risk": "保守困難・微小性能ロス・診断フラグ誤ON時の大量出力",
        "freq": "低（通常OFF）／高（誤ON時）",
        "tradeoff": "外出しすると現場診断の即応性が下がる。フラグ付き別モジュールが妥協点",
        "difficulty": "中",
    },
    {
        "id": 21,
        "priority": "中",
        "score": 4,
        "items": "速度",
        "feature": "データ集約UI・前面化／ロック再試行",
        "title": "ui_data_agg の processEvents / 多重 QTimer",
        "impact": "再入・ちらつき低減",
        "risk": "ダイアログちらつき・再入バグ・稀なフリーズ感",
        "freq": "中（前面化・Excel競合時）",
        "tradeoff": "状態マシン化は実装・検証コスト高。現状の簡易再試行は実装が短い",
        "difficulty": "中",
    },
    {
        "id": 22,
        "priority": "低",
        "score": 3,
        "items": "スリム化 / ファイル容量",
        "feature": "自動テスト／CI",
        "title": "テスト肥大・重複シナリオの整理",
        "impact": "CI時間短縮",
        "risk": "CI遅延・重複テストのメンテ漏れで偽緑",
        "freq": "中（CI実行のたび）",
        "tradeoff": "統合しすぎるとカバレッジ穴。回帰の重要ケースは残す必要がある",
        "difficulty": "中",
    },
]


def priority_label(f: dict) -> str:
    if f["priority"] == "すぐに":
        return f"すぐに {f['score']}"
    return f"{f['priority']}{f['score']}"


def main() -> None:
    out = Path(__file__).resolve().parent
    findings = sorted(FINDINGS, key=lambda x: -x["score"])

    md_lines = [
        "# コードモジュール見直し — 改善一覧",
        "",
        "- 作成日: 2026-09-11",
        "- 対象: Excel_AddIn（CSV Tool）全体",
        "- 注記: 頻度・トレードオフはコード経路と利用想定からの試算（実測テレメトリではない）",
        "",
        "## 最優先の結論",
        "",
        "正確性は結合キー正規化の不一致・抽出失敗の黙殺・シート名フォールバックが先。"
        "容量は logs/dist/_tmp 掃除で約4.5GB回収可能。構造は巨大モノリス分割が中期の本丸。",
        "",
        "## 改善優先度一覧（スコア降順）",
        "",
        "| # | 優先度 | 改善アイテム | 対象機能名 | 改善内容 | 改善度合い | "
        "未改善時に起こり得る事象 | 頻度 | トレードオフ | 難易度 |",
        "|---|--------|--------------|------------|----------|------------|"
        "--------------------------|------|--------------|--------|",
    ]
    for f in findings:
        cells = [
            str(f["id"]),
            priority_label(f),
            f["items"],
            f["feature"],
            f["title"],
            f["impact"],
            f["risk"],
            f["freq"],
            f["tradeoff"],
            f["difficulty"],
        ]
        md_lines.append("| " + " | ".join(cells) + " |")

    md_lines.extend(
        [
            "",
            "## 推奨着手順",
            "",
            "### すぐ（容量＋正確性の止血）",
            "1. logs/nuitka・dist旧版・_tmp_* / .coverage / 重複テスト削除＋gitignore",
            "2. シート不存在をエラー化・抽出失敗をログ＋区別可能値に",
            "3. 結合キー正規化方針を仕様文書と突合",
            "",
            "### 高〜中（構造と性能）",
            "4. svc/Old削除・DEBUGパッチ除去・_get_polars一本化",
            "5. materialize範囲制限・hidden二重オープン削減・except縮減",
            "6. svc_data_agg / ui_data_agg_debug の分割",
            "",
            "## 頻度の読み方",
            "",
            "| 表記 | 意味 |",
            "|------|------|",
            "| 高 | 該当機能利用時にほぼ毎回、または日常運用で頻発 |",
            "| 中 | 特定条件（大ファイル・設定ON・改修時など）で発生 |",
            "| 低 | 稀、または触ったとき／誤設定時のみ |",
            "",
        ]
    )
    md_path = out / "コードモジュール見直し_改善一覧.md"
    md_path.write_text("\n".join(md_lines), encoding="utf-8")

    esc = html.escape
    rows_html = []
    for f in findings:
        cells = [
            f["id"],
            priority_label(f),
            f["items"],
            f["feature"],
            f["title"],
            f["impact"],
            f["risk"],
            f["freq"],
            f["tradeoff"],
            f["difficulty"],
        ]
        rows_html.append(
            "<tr>" + "".join(f"<td>{esc(str(c))}</td>" for c in cells) + "</tr>"
        )

    html_doc = f"""<!DOCTYPE html>
<html lang="ja">
<head>
<meta charset="utf-8" />
<meta name="viewport" content="width=device-width, initial-scale=1" />
<title>コードモジュール見直し — 改善一覧</title>
<style>
  body {{ font-family: "Segoe UI", "Yu Gothic UI", Meiryo, sans-serif; margin: 24px; color: #111; line-height: 1.45; }}
  h1 {{ font-size: 22px; margin: 0 0 8px; }}
  h2 {{ font-size: 16px; margin: 28px 0 10px; border-bottom: 1px solid #ccc; padding-bottom: 4px; }}
  h3 {{ font-size: 13px; margin: 14px 0 6px; }}
  .meta {{ color: #555; font-size: 12px; margin-bottom: 16px; }}
  .callout {{ background: #fff8e6; border: 1px solid #e6c200; padding: 10px 12px; margin: 12px 0 20px; }}
  table {{ border-collapse: collapse; width: 100%; font-size: 11px; }}
  th, td {{ border: 1px solid #bbb; padding: 6px 7px; vertical-align: top; text-align: left; }}
  th {{ background: #f3f3f3; }}
  tr:nth-child(even) td {{ background: #fafafa; }}
  ol {{ padding-left: 1.3em; }}
  .toolbar {{ margin: 0 0 16px; }}
  .toolbar button {{ padding: 8px 14px; font-size: 13px; cursor: pointer; }}
  @media print {{
    body {{ margin: 10mm; }}
    .toolbar {{ display: none; }}
    h2 {{ break-after: avoid; }}
    tr {{ break-inside: avoid; }}
    @page {{ size: A4 landscape; margin: 10mm; }}
  }}
</style>
</head>
<body>
  <div class="toolbar"><button type="button" onclick="window.print()">印刷 / PDF保存</button></div>
  <h1>コードモジュール見直し — 改善一覧</h1>
  <div class="meta">作成日: 2026-09-11 / 対象: Excel_AddIn（CSV Tool）全体 / 頻度・トレードオフは試算</div>
  <div class="callout"><strong>最優先の結論:</strong> 正確性は結合キー正規化の不一致・抽出失敗の黙殺・シート名フォールバックが先。容量は logs/dist/_tmp 掃除で約4.5GB回収可能。構造は巨大モノリス分割が中期の本丸。</div>
  <h2>改善優先度一覧（スコア降順）</h2>
  <table>
    <thead><tr>
      <th>#</th><th>優先度</th><th>改善アイテム</th><th>対象機能名</th><th>改善内容</th>
      <th>改善度合い</th><th>未改善時に起こり得る事象</th><th>頻度</th><th>トレードオフ</th><th>難易度</th>
    </tr></thead>
    <tbody>
{chr(10).join(rows_html)}
    </tbody>
  </table>
  <h2>推奨着手順</h2>
  <h3>すぐ（容量＋正確性の止血）</h3>
  <ol>
    <li>logs/nuitka・dist旧版・_tmp_* / .coverage / 重複テスト削除＋gitignore</li>
    <li>シート不存在をエラー化・抽出失敗をログ＋区別可能値に</li>
    <li>結合キー正規化方針を仕様文書と突合</li>
  </ol>
  <h3>高〜中（構造と性能）</h3>
  <ol start="4">
    <li>svc/Old削除・DEBUGパッチ除去・_get_polars一本化</li>
    <li>materialize範囲制限・hidden二重オープン削減・except縮減</li>
    <li>svc_data_agg / ui_data_agg_debug の分割</li>
  </ol>
  <h2>頻度の読み方</h2>
  <table>
    <thead><tr><th>表記</th><th>意味</th></tr></thead>
    <tbody>
      <tr><td>高</td><td>該当機能利用時にほぼ毎回、または日常運用で頻発</td></tr>
      <tr><td>中</td><td>特定条件（大ファイル・設定ON・改修時など）で発生</td></tr>
      <tr><td>低</td><td>稀、または触ったとき／誤設定時のみ</td></tr>
    </tbody>
  </table>
</body>
</html>
"""
    html_path = out / "コードモジュール見直し_改善一覧.html"
    html_path.write_text(html_doc, encoding="utf-8")
    print(f"wrote {md_path}")
    print(f"wrote {html_path}")


if __name__ == "__main__":
    main()
