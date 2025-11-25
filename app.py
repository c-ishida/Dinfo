import streamlit as st
import pandas as pd
import os

# Page configuration
st.set_page_config(page_title="お薬の説明 - 漢方薬局ハレノヴァ", page_icon="🔍")

st.title("🌿 お薬の説明")

# File path
FILE_PATH = "処方の説明.xlsx"

@st.cache_resource
def load_excel_file(file_path):
    """Loads the Excel file and returns the ExcelFile object."""
    try:
        return pd.ExcelFile(file_path)
    except FileNotFoundError:
        return None
    except Exception as e:
        return str(e)

# Check if file exists
if not os.path.exists(FILE_PATH):
    st.error(f"File not found: {FILE_PATH}")
    st.info("Please make sure the file '処方の説明.xlsx' is in the same directory.")
else:
    xls = load_excel_file(FILE_PATH)
    
    if isinstance(xls, str): # Error occurred
        st.error(f"Error loading file: {xls}")
    elif xls:
        # Sheet selection (hidden - auto-select first sheet)
        sheet_names = xls.sheet_names
        selected_sheet = sheet_names[0] if sheet_names else None
        
        if selected_sheet:
            # Load data from selected sheet
            df = pd.read_excel(xls, sheet_name=selected_sheet)
            
            # Fix: Convert 検索番号 column to string to prevent Arrow serialization errors
            if '検索番号' in df.columns:
                df['検索番号'] = df['検索番号'].astype(str)
            
            # Helper function to normalize text for search
            def normalize_text(text):
                """Normalize text: hyphens + full-width alphanumeric to half-width + lowercase"""
                if not isinstance(text, str):
                    text = str(text)
                
                # Replace various hyphen characters with half-width hyphen
                text = text.replace('−', '-').replace('ー', '-').replace('—', '-').replace('–', '-').replace('‐', '-')
                
                # Convert full-width alphanumeric to half-width
                # Full-width: ０-９, Ａ-Ｚ, ａ-ｚ → Half-width: 0-9, A-Z, a-z
                full_to_half = str.maketrans(
                    '０１２３４５６７８９ＡＢＣＤＥＦＧＨＩＪＫＬＭＮＯＰＱＲＳＴＵＶＷＸＹＺａｂｃｄｅｆｇｈｉｊｋｌｍｎｏｐｑｒｓｔｕｖｗｘｙｚ',
                    '0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz'
                )
                text = text.translate(full_to_half)
                
                # Convert to lowercase for case-insensitive search
                return text.lower()
            
            
            # Search Interface
            # Initialize session state for tracking if search has been performed
            if 'search_performed' not in st.session_state:
                st.session_state.search_performed = False
            
            # Instruction text above input
            st.markdown("""
            <div style='margin-bottom: 10px;'>
                <p style='margin: 0; font-size: 18px;'>お薬の番号を入力してください。</p>
                <p style='margin: 3px 0 0 0; font-size: 13px; color: #666;'>複数ある場合は、スペース または カンマ(,)で区切ってください。</p>
            </div>
            """, unsafe_allow_html=True)
            
            # Show checkbox only after first search
            if st.session_state.search_performed:
                col1, col2, col3 = st.columns([3, 1, 1])
            else:
                col1, col2 = st.columns([3, 1])
            
            with col1:
                search_query = st.text_input("お薬の番号", "", 
                                             placeholder="ここに番号を入力",
                                             help="複数ある場合は、スペース または カンマ(,)で区切ってください",
                                             key="search_input",
                                             label_visibility="collapsed")
            with col2:
                # Add some spacing to align with the button
                search_button = st.button("🔍 検索する", use_container_width=True)
            
            # Show exact match checkbox only after first search
            if st.session_state.search_performed:
                with col3:
                    # Add some spacing to align with the checkbox
                    exact_match = st.checkbox("完全一致", value=True)
            else:
                # Default to exact match for first search
                exact_match = True
            
            # Trigger search if button is clicked or if there's text in the input
            if search_button or search_query:
                # Mark that search has been performed
                st.session_state.search_performed = True
                # Parse multiple search terms (split by comma or space)
                import re
                search_terms = [term.strip() for term in re.split(r'[,\s]+', search_query) if term.strip()]
                
                # Normalize search terms (hyphens, full-width chars, case)
                search_terms = [normalize_text(term) for term in search_terms]
                
                # If no search terms (empty input), show all data
                if not search_terms:
                    with st.expander("すべてのお薬を表示する", expanded=True):
                        st.dataframe(df)
                else:
                    # Filter data
                    if exact_match:
                        # Exact match logic - match any of the search terms
                        masks = []
                        for term in search_terms:
                            # Normalize data before comparison
                            term_mask = df.astype(str).apply(lambda x: x.apply(normalize_text)).apply(lambda x: (x == term).any(), axis=1)
                            masks.append(term_mask)
                        # Combine all masks with OR logic
                        mask = pd.concat(masks, axis=1).any(axis=1)
                    else:
                        # Partial match logic - match any of the search terms
                        masks = []
                        for term in search_terms:
                            # Normalize data before comparison
                            term_mask = df.astype(str).apply(lambda x: x.apply(normalize_text)).apply(lambda x: x.str.contains(term, case=False, na=False).any(), axis=1)
                            masks.append(term_mask)
                        # Combine all masks with OR logic
                        mask = pd.concat(masks, axis=1).any(axis=1)
                    
                    results = df[mask]
                    
                    st.write(f"{len(results)}件 見つかりました")
                    st.dataframe(results)
                    
                    # Print layout - auto-show if search button was clicked, otherwise show button
                    if len(results) > 0:
                        # If search button was clicked, automatically show print layout
                        if search_button:
                            # Generate print-friendly HTML
                            from datetime import datetime
                            import html
                            import streamlit.components.v1 as components
                            
                            # Generate custom layout for each result
                            results_html = ""
                            for idx, row in results.iterrows():
                                # Get column values by name
                                search_num = html.escape(str(row.get('検索番号', '')))
                                prescription_name = html.escape(str(row.get('処方名', '')))
                                description = html.escape(str(row.get('説明', '')))
                                
                                results_html += f"""
                                <div class='result-item'>
                                    <div class='first-line'>
                                        <span class='prescription-name'>{prescription_name}</span>
                                        <span class='search-number'>{search_num}</span>
                                    </div>
                                    <div class='description-section'>
                                        <div class='description-content'>{description}</div>
                                    </div>
                                </div>
                                """
                            
                            # Complete HTML
                            now = datetime.now().strftime("%Y年%m月%d日 %H:%M")
                            html_content = f"""
                            <!DOCTYPE html>
                            <html>
                            <head>
                                <meta charset="UTF-8">
                                <style>
                                    @media print {{
                                        @page {{
                                            size: A5;
                                            margin: 10mm;
                                        }}
                                        .no-print {{
                                            display: none;
                                        }}
                                        .print-info {{
                                            display: none;
                                        }}
                                    }}
                                    body {{
                                        font-family: 'Meiryo', 'MS Gothic', 'Yu Gothic', sans-serif;
                                        margin: 0;
                                        padding: 20px;
                                    }}
                                    .print-container {{
                                        max-width: 148mm;
                                        margin: 0 auto;
                                        background: white;
                                    }}
                                    .print-header {{
                                        display: flex;
                                        justify-content: space-between;
                                        align-items: center;
                                        font-weight: bold;
                                        font-size: 14pt;
                                        margin-bottom: 15px;
                                        border-bottom: 2px solid #000;
                                        padding-bottom: 5px;
                                        color: #000 !important;
                                    }}
                                    .pharmacy-name {{
                                        font-size: 9pt;
                                        font-weight: normal;
                                    }}
                                    .print-info {{
                                        font-size: 8pt;
                                        margin-bottom: 10px;
                                        color: #666;
                                    }}
                                    .result-item {{
                                        margin-bottom: 20px;
                                        page-break-inside: avoid;
                                    }}
                                    .first-line {{
                                        display: flex;
                                        justify-content: space-between;
                                        align-items: center;
                                        border-bottom: 1px solid #000;
                                        padding-bottom: 5px;
                                        margin-bottom: 10px;
                                    }}
                                    .prescription-name {{
                                        font-weight: bold;
                                        font-size: 12pt;
                                        flex-grow: 1;
                                        white-space: nowrap;
                                    }}
                                    .search-number {{
                                        font-size: 10pt;
                                        text-align: right;
                                        margin-left: 20px;
                                    }}
                                    .description-section {{
                                        font-size: 9pt;
                                    }}
                                    .description-label {{
                                        font-weight: bold;
                                        margin-bottom: 5px;
                                    }}
                                    .description-content {{
                                        white-space: pre-wrap;
                                        word-wrap: break-word;
                                        line-height: 1.5;
                                    }}
                                    .print-button {{
                                        text-align: center;
                                        margin-top: 20px;
                                    }}
                                    button {{
                                        padding: 12px 24px;
                                        font-size: 14pt;
                                        cursor: pointer;
                                        background-color: #4CAF50;
                                        color: white;
                                        border: none;
                                        border-radius: 5px;
                                    }}
                                    button:hover {{
                                        background-color: #45a049;
                                    }}
                                </style>
                            </head>
                            <body>
                                <div class='print-container'>
                                    <div class='print-header'>
                                        <span>お薬の説明</span>
                                        <span class='pharmacy-name'>漢方薬局ハレノヴァ</span>
                                    </div>
                                    <div class='print-info'>検索語: {html.escape(', '.join(search_terms))} / 件数: {len(results)}件 / 出力日時: {now}</div>
                                    {results_html}
                                </div>
                                <div class='print-button no-print'>
                                    <button onclick='window.print()'>
                                        🖨️ 印刷する (Ctrl+P)
                                    </button>
                                </div>
                            </body>
                            </html>
                            """
                            
                            components.html(html_content, height=800, scrolling=True)
                            st.info("💡 上の「印刷する」ボタンをクリックするか、Ctrl+P で印刷ダイアログを開いてください。")
                        else:
                            # For Enter key users, show manual button
                            st.markdown("---")
                            
                            # Initialize session state for print layout
                            if 'show_print_layout' not in st.session_state:
                                st.session_state.show_print_layout = False
                            
                            if st.button("📄 印刷用レイアウトを表示"):
                                st.session_state.show_print_layout = True
                                st.rerun()
                            
                            # Show print layout if button was clicked
                            if st.session_state.show_print_layout:
                                # Generate print-friendly HTML
                                from datetime import datetime
                                import html
                                import streamlit.components.v1 as components
                                
                                # Generate custom layout for each result
                                results_html = ""
                                for idx, row in results.iterrows():
                                    # Get column values by name
                                    search_num = html.escape(str(row.get('検索番号', '')))
                                    prescription_name = html.escape(str(row.get('処方名', '')))
                                    description = html.escape(str(row.get('説明', '')))
                                    
                                    results_html += f"""
                                    <div class='result-item'>
                                        <div class='first-line'>
                                            <span class='prescription-name'>{prescription_name}</span>
                                            <span class='search-number'>{search_num}</span>
                                        </div>
                                        <div class='description-section'>
                                            <div class='description-content'>{description}</div>
                                        </div>
                                    </div>
                                    """
                                
                                # Complete HTML
                                now = datetime.now().strftime("%Y年%m月%d日 %H:%M")
                                html_content = f"""
                                <!DOCTYPE html>
                                <html>
                                <head>
                                    <meta charset="UTF-8">
                                    <style>
                                        @media print {{
                                            @page {{
                                                size: A5;
                                                margin: 10mm;
                                            }}
                                            .no-print {{
                                                display: none;
                                            }}
                                            .print-info {{
                                                display: none;
                                            }}
                                        }}
                                        body {{
                                            font-family: 'MS Gothic', 'Yu Gothic', sans-serif;
                                            margin: 0;
                                            padding: 20px;
                                        }}
                                        .print-container {{
                                            max-width: 148mm;
                                            margin: 0 auto;
                                            background: white;
                                        }}
                                        .print-header {{
                                            display: flex;
                                            justify-content: space-between;
                                            align-items: center;
                                            font-weight: bold;
                                            font-size: 14pt;
                                            margin-bottom: 15px;
                                            border-bottom: 2px solid #000;
                                            padding-bottom: 5px;
                                            color: #000 !important;
                                        }}
                                        .pharmacy-name {{
                                            font-size: 9pt;
                                            font-weight: normal;
                                        }}
                                        .print-info {{
                                            font-size: 8pt;
                                            margin-bottom: 10px;
                                            color: #666;
                                        }}
                                        .result-item {{
                                            margin-bottom: 20px;
                                            page-break-inside: avoid;
                                        }}
                                        .first-line {{
                                            display: flex;
                                            justify-content: space-between;
                                            align-items: center;
                                            border-bottom: 1px solid #000;
                                            padding-bottom: 5px;
                                            margin-bottom: 10px;
                                        }}
                                        .prescription-name {{
                                            font-weight: bold;
                                            font-size: 12pt;
                                            flex-grow: 1;
                                            white-space: nowrap;
                                        }}
                                        .search-number {{
                                            font-size: 10pt;
                                            text-align: right;
                                            margin-left: 20px;
                                        }}
                                        .description-section {{
                                            font-size: 9pt;
                                        }}
                                        .description-label {{
                                            font-weight: bold;
                                            margin-bottom: 5px;
                                        }}
                                        .description-content {{
                                            white-space: pre-wrap;
                                            word-wrap: break-word;
                                            line-height: 1.5;
                                        }}
                                        .print-button {{
                                            text-align: center;
                                            margin-top: 20px;
                                        }}
                                        button {{
                                            padding: 12px 24px;
                                            font-size: 14pt;
                                            cursor: pointer;
                                            background-color: #4CAF50;
                                            color: white;
                                            border: none;
                                            border-radius: 5px;
                                        }}
                                        button:hover {{
                                            background-color: #45a049;
                                        }}
                                    </style>
                                </head>
                                <body>
                                    <div class='print-container'>
                                        <div class='print-header'>
                                            <span>お薬の説明</span>
                                            <span class='pharmacy-name'>漢方薬局ハレノヴァ</span>
                                        </div>
                                        <div class='print-info'>検索語: {html.escape(', '.join(search_terms))} / 件数: {len(results)}件 / 出力日時: {now}</div>
                                        {results_html}
                                    </div>
                                    <div class='print-button no-print'>
                                        <button onclick='window.print()'>
                                            🖨️ 印刷する (Ctrl+P)
                                        </button>
                                    </div>
                                </body>
                                </html>
                                """
                                
                                components.html(html_content, height=800, scrolling=True)
                                st.info("💡 上の「印刷する」ボタンをクリックするか、Ctrl+P で印刷ダイアログを開いてください。")
                                
                                # Reset the flag
                                st.session_state.show_print_layout = False



