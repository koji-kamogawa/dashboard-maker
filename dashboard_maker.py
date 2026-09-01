#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
SharePoint「サイトの使用状況データ」統合ダッシュボード生成 Streamlit アプリ
================================================================================

build_dashboard3.py のロジック（xlsx読み込み・集計・グラフ生成・HTML組み立て）を
そのまま再利用し、ブラウザ上で完結する UI を提供する。

使い方
------
    streamlit run streamlit_app.py

処理の流れ
----------
1. ユーザーが同形式の xlsx（毎週エクスポート分）を複数アップロードする。
2. 「ダッシュボードを生成」ボタンを押すと、build_dashboard3.py の
   load_all_from_sources() / build_html_string() を呼び出して
   自己完結型HTML（グラフ画像・カスタム期間用データを埋め込んだ1ファイル）を組み立てる。
3. 生成結果のサマリー・警告を画面に表示し、st.download_button で
   dashboard.html としてダウンロードできるようにする（同一画面内でのプレビューも可能）。
"""

import datetime
from io import BytesIO

import streamlit as st

from build_dashboard3 import load_all_from_sources, build_html_string


st.set_page_config(
    page_title="サイト使用状況ダッシュボード生成",
    page_icon="📊",
    layout="wide",
)

st.title("📊 サイト使用状況ダッシュボード生成ツール")
st.caption(
    "SharePoint からダウンロードした「サイトの使用状況データ」の xlsx を複数アップロードすると、"
    "1つの自己完結型 HTML ダッシュボード（dashboard.html）を作成してダウンロードできます。"
)

with st.expander("ℹ️ 対象ファイルの形式・前提について（クリックで開く）", expanded=False):
    st.markdown(
        """
- 各 xlsx は SharePoint の「サイトの使用状況データ」エクスポートで、以下のシート構成を想定しています。
  - `全体的なトラフィック` / `人気のあるコンテンツ` / `デバイス別の使用状況` / `時間別の使用状況`
- 同じサイトから **毎週エクスポート** したファイルを想定しており、ファイル名は問いません（中身のシート構成のみを見ます）。
- `全体的なトラフィック` と `デバイス別の使用状況` は直近90日分が毎回含まれるため、複数週のファイルを重ねると日付が重複します。
  → 同じ日付が複数ファイルにある場合は、**シート内の「ダウンロード日」記載が最も新しいファイルの値**を採用します。
- シート内に「ダウンロード日」の記載が見つからない場合は、**本日の日付**をフォールバックとして使用します
  （フォルダ内ファイルを直接読み込む CLI版ではファイル更新日時を使用しますが、アップロード版ではその情報を取得できないためです）。
- 「重複しない閲覧者数(UU)」は日をまたいで重複しうるため、サマリー表は SharePoint 公式の集計値（最新ファイル記載分）をそのまま表示します。
- 生成された HTML 内では、固定の5期間（1週間/1ヶ月/3ヶ月/6ヶ月/全期間）に加えて、
  **「カスタム期間」タブでユーザーが任意の開始日・終了日を指定**してグラフを再描画できます（ブラウザ内JavaScriptで完結、外部通信なし）。
        """
    )

st.divider()

uploaded_files = st.file_uploader(
    "xlsx ファイルをアップロード（複数選択可）",
    type=["xlsx"],
    accept_multiple_files=True,
    help="毎週エクスポートした同形式の xlsx ファイルをまとめて選択してください。",
)

if uploaded_files:
    st.write(f"**{len(uploaded_files)} 件** のファイルが選択されました。")
    with st.expander("選択されたファイル一覧を表示"):
        for f in uploaded_files:
            st.write(f"- {f.name}（{f.size:,} bytes）")

col_a, col_b = st.columns([1, 3])
with col_a:
    generate = st.button("🚀 ダッシュボードを生成", type="primary", disabled=not uploaded_files)

if generate and uploaded_files:
    with st.spinner("読み込み・集計・グラフ生成を実行中です…（ファイル数や期間が多いと数十秒かかることがあります）"):
        # --- アップロードファイルを load_all_from_sources() 用のソースリストに変換 ---
        # openpyxl.load_workbook() はファイルパスだけでなく file-like オブジェクトも受け取れるため、
        # Streamlit の UploadedFile（BytesIOと同様に扱える）をそのまま渡す。
        # 「ダウンロード日」がシート内に見つからなかった場合のフォールバックには本日の日付を用いる。
        today = datetime.date.today()
        sources = []
        for f in uploaded_files:
            buf = BytesIO(f.getvalue())
            buf.name = f.name
            sources.append((f.name, buf, today))

        try:
            data = load_all_from_sources(sources)
        except Exception as e:
            st.error(f"読み込み中にエラーが発生しました: {e}")
            st.stop()

        if data["traffic"].empty:
            st.error(
                "有効なデータが1件も読み込めませんでした。アップロードした xlsx の"
                "シート構成をご確認ください（「全体的なトラフィック」等の想定シートが必要です）。"
            )
            st.stop()

        # --- 読み込み結果のサマリー・警告表示 ---
        c1, c2, c3 = st.columns(3)
        c1.metric("読み込みファイル数", data["n_files"])
        c1.metric("統合後の日次データ件数", len(data["traffic"]))
        c2.metric("最新データ取得日", str(data["latest_dl_date"]))
        n_days = (data["traffic"]["日付"].max() - data["traffic"]["日付"].min()).days + 1
        c2.metric("データ期間の日数", f"{n_days} 日")
        c3.metric("欠測期間の件数", len(data["gaps"]))
        c3.metric("ファイル警告件数", len(data["file_warnings"]))

        if data["file_warnings"]:
            with st.expander(f"⚠️ ファイルに関する警告（{len(data['file_warnings'])}件）", expanded=True):
                for name, msg in data["file_warnings"]:
                    st.warning(f"**{name}**: {msg}")

        if data["gaps"]:
            with st.expander(f"⚠️ データの欠測期間（{len(data['gaps'])}件）", expanded=False):
                st.write(
                    "90日を超えて次のファイルをダウンロードしなかった期間がある可能性があります。"
                )
                for prev, cur, g in data["gaps"][:20]:
                    st.write(f"- {prev} 〜 {cur}（{g - 1}日分欠測）")

        # --- HTML組み立て ---
        html_str = build_html_string(data)
        html_bytes = html_str.encode("utf-8")

    st.success("✅ ダッシュボードの生成が完了しました。")

    st.download_button(
        label="⬇️ dashboard.html をダウンロード",
        data=html_bytes,
        file_name="dashboard.html",
        mime="text/html",
        type="primary",
    )

    st.caption(
        f"ファイルサイズ: 約 {len(html_bytes) / 1024:.0f} KB "
        "（グラフ画像・全期間データを1ファイルに埋め込んでいるため、他の環境でもこのファイル単体で閲覧できます）"
    )

    with st.expander("🔍 生成されたダッシュボードをこの画面でプレビュー", expanded=False):
        st.components.v1.html(html_str, height=1400, scrolling=True)

elif not uploaded_files:
    st.info("👆 まずは xlsx ファイルをアップロードしてください。")
