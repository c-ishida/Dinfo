#!/home/tochimoto/miniconda3/envs/medicine-app/bin/python
# -*- coding: utf-8 -*-

import cgi
import cgitb
import pandas as pd
import os
import re
import html
from datetime import datetime

# デバッグ用
cgitb.enable()

# 設定
EXCEL_FILE = "処方の説明.xlsx"

def normalize_text(text):
    if not isinstance(text, str):
        text = str(text)
    text = text.replace('−', '-').replace('ー', '-').replace('ｰ', '-').replace('—', '-').replace('–', '-').replace('‐', '-')
    full_to_half = str.maketrans(
        '０１２３４５６７８９ＡＢＣＤＥＦＧＨＩＪＫＬＭＮＯＰＱＲＳＴＵＶＷＸＹＺａｂｃｄｅｆｇｈｉｊｋｌｍｎｏｐｑｒｓｔｕｖｗｘｙｚ',
        '0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz'
    )
    return text.translate(full_to_half).lower()

def print_html(content):
    print("Content-type: text/html; charset=utf-8\n")
    print(content)

def main():
    form = cgi.FieldStorage()
    search_query = form.getfirst("q", "")
    
    # 2回目以降かどうかの厳密な判定：
    # s=1 が含まれているか、またはチェックボックス(exact)のパラメータが送られてきているか
    is_subsequent = ("s" in form) or ("exact" in form)
    
    if search_query:
        if not is_subsequent:
            # 【初回】URLにパラメータがない真っさらな状態からの検索は、強制的に完全一致
            exact_match = True
        else:
            # 【2回目以降】チェックが入っている時だけ完全一致
            exact_match = (form.getfirst("exact") == "on")
    else:
        # 初期表示
        exact_match = True

    # 画面に表示するチェックボックスの状態
    # はじめて結果が出た直後は、次のために「OFF」で表示する。
    # 2回目以降は、今の検索設定（exact_match）をそのまま表示に反映させる。
    display_checked = exact_match if is_subsequent else False

    # 更新完了メッセージ
    msg_param = form.getfirst("m", "")
    updated_msg = ""
    if msg_param == "updated":
        updated_msg = "<div class='no-print' style='background:#e8f5e9; color:#2e7d32; padding:15px; border-radius:8px; margin-bottom:20px; border:1px solid #c8e6c9; font-weight:bold; text-align:center;'>✅ Excelデータが正常に更新されました！</div>"

    html_out = f"""
    <!DOCTYPE html>
    <html lang="ja">
    <head>
        <meta charset="UTF-8">
        <title>お薬の説明</title>
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <style>
            @media print {{
                @page {{ size: A5; margin: 10mm; }}
                .no-print {{ display: none !important; }}
                body {{ padding: 0; background: #ffffff !important; }}
                .print-container {{ border: none !important; padding: 0 !important; max-width: 100% !important; }}
            }}
            body {{
                font-family: 'Inter', 'Segoe UI', 'Meiryo', sans-serif;
                margin: 0;
                padding: 20px;
                background-color: #ffffff;
                color: #31333F;
                line-height: 1.6;
            }}
            .main-wrapper {{
                max-width: 700px;
                margin: 0 auto;
            }}
            h1 {{ font-size: 2.5rem; font-weight: 700; margin-bottom: 1.5rem; }}
            .instruction-header {{ margin-bottom: 10px; }}
            .instruction-main {{ margin: 0; font-size: 18px; font-weight: 400; }}
            .instruction-sub {{ margin: 3px 0 0 0; font-size: 13px; color: #666; }}
            .search-form {{
                display: flex;
                flex-wrap: wrap;
                gap: 12px;
                align-items: flex-start;
                margin-bottom: 1rem;
            }}
            .input-wrapper {{ flex: 3; min-width: 200px; }}
            .button-wrapper {{ flex: 1; min-width: 120px; }}
            .checkbox-wrapper {{ flex: 1.2; min-width: 120px; display: flex; align-items: center; height: 46px; }}
            input[type="text"] {{
                width: 100%;
                padding: 10px 12px;
                font-size: 1rem;
                border: 1px solid rgba(49, 51, 63, 0.2);
                border-radius: 0.5rem;
                box-sizing: border-box;
                background-color: #ffffff;
            }}
            input[type="submit"] {{
                width: 100%;
                padding: 10px 12px;
                font-size: 1rem;
                background-color: #ffffff;
                color: #31333F;
                border: 1px solid rgba(49, 51, 63, 0.2);
                border-radius: 0.5rem;
                cursor: pointer;
            }}
            input[type="submit"]:hover {{
                border-color: #ff4b4b;
                color: #ff4b4b;
            }}
            .checkbox-label {{
                font-size: 14px;
                cursor: pointer;
                display: flex;
                align-items: center;
                gap: 8px;
            }}
            .result-count {{ font-size: 14px; margin-bottom: 20px; }}
            .result-item {{ margin-bottom: 30px; page-break-inside: avoid; }}
            .first-line {{ display: flex; justify-content: space-between; align-items: center; border-bottom: 1px solid #000; padding-bottom: 5px; margin-bottom: 10px; }}
            .prescription-name {{ font-weight: bold; font-size: 12pt; }}
            .search-number {{ font-size: 10pt; text-align: right; margin-left: 20px; }}
            .description-content {{ white-space: pre-wrap; word-wrap: break-word; line-height: 1.6; font-size: 10pt; }}
            .print-button-container {{ margin-bottom: 24px; }}
            .print-btn {{ padding: 0.5rem 1rem; font-size: 1rem; background-color: #4CAF50; color: white; border: none; border-radius: 0.5rem; cursor: pointer; }}
            .print-container {{ background: #ffffff; padding: 0; }}
            .print-header {{ font-weight: bold; font-size: 14pt; margin-bottom: 15px; border-bottom: 2px solid #000; padding-bottom: 5px; }}
        </style>
    </head>
    <body>
        <div class="main-wrapper">
            {updated_msg}
            <div class="no-print">
                <h1>🔍 お薬の説明</h1>
                <div class="instruction-header">
                    <p class="instruction-main">お薬の番号を入力してください。</p>
                    <p class="instruction-sub">複数ある場合は、スペース または カンマ(,)で区切ってください。</p>
                </div>
                <form method="GET" action="app.py" class="search-form">
                    <input type="hidden" name="s" value="1">
                    <div class="input-wrapper">
                        <input type="text" name="q" value="{html.escape(search_query)}" placeholder="ここに番号を入力">
                    </div>
                    <div class="button-wrapper">
                        <input type="submit" value="🔍 検索する">
                    </div>
                    {"<div class='checkbox-wrapper'><label class='checkbox-label'><input type='checkbox' name='exact' " + ("checked" if display_checked else "") + "> 完全一致</label></div>" if search_query else ""}
                </form>
            </div>
    """

    if search_query:
        if not os.path.exists(EXCEL_FILE):
            html_out += f"<div class='no-print'><p style='color:red;'>エラー: {EXCEL_FILE} が見つかりません。</p></div>"
        else:
            try:
                df = pd.read_excel(EXCEL_FILE, engine='openpyxl')
                # 検索前に全ての列を一括で文字列化して正規化（高速化と型の不一致防止）
                df_str = df.astype(str).apply(lambda x: x.apply(normalize_text))
                
                terms = [normalize_text(t) for t in re.split(r'[,\uff0c\u3001\s]+', search_query) if t.strip()]
                
                if terms:
                    mask = pd.Series([False] * len(df))
                    for term in terms:
                        if exact_match:
                            # 完全一致：どの列かの値が term と完全に一致するか
                            term_mask = (df_str == term).any(axis=1)
                        else:
                            # 部分一致：どの列かの値に term が含まれているか
                            term_mask = df_str.apply(lambda x: x.str.contains(term, na=False)).any(axis=1)
                        mask |= term_mask
                    
                    results = df[mask]
                    if len(results) > 0:
                        html_out += f"<div class='result-count no-print'>{len(results)}件が見つかりました</div>"
                        html_out += f"""
                        <div class="print-button-container no-print">
                            <button class="print-btn" onclick="window.print()">🖨️ 印刷する</button>
                        </div>
                        <div class="print-container">
                            <div class="print-header">
                                <span>お薬の説明</span>
                            </div>
                        """
                        for _, row in results.iterrows():
                            html_out += f"""
                            <div class="result-item">
                                <div class="first-line">
                                    <span class="prescription-name">{html.escape(str(row.get('処方名', '')))}</span>
                                    <span class="search-number">{html.escape(str(row.get('検索番号', '')))}</span>
                                </div>
                                <div class="description-content">{html.escape(str(row.get('説明', '')))}</div>
                            </div>
                            """
                        html_out += "</div>"
                    else:
                        html_out += "<div class='no-print'><p>該当するお薬が見つかりませんでした。</p></div>"
            except Exception as e:
                html_out += f"<div class='no-print'><p style='color:red;'>エラー: {str(e)}</p></div>"

    html_out += "</div></body></html>"
    print_html(html_out)

if __name__ == "__main__":
    main()