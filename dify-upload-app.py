import streamlit as st
from dotenv import load_dotenv
import requests
import os
import mammoth
import re
from pathlib import Path
import openpyxl
import tempfile
import io

def get_file_extension(file_path):
    return Path(file_path).suffix

def convert_with_openpyxl(input_file_data, file_name):
    """
    openpyxl を使用してExcelファイルをテキストファイルに変換する
    """
    # 一時ファイルを作成してExcelデータを保存
    with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as temp_file:
        temp_file.write(input_file_data)
        excel_path = temp_file.name
    
    try:
        # Excelファイルを読み込む
        workbook = openpyxl.load_workbook(excel_path, data_only=True)
        
        output_text = ""
        # 各シートの内容を処理
        for sheet_name in workbook.sheetnames:
            sheet = workbook[sheet_name]
            output_text += f"===== シート: {sheet_name} =====\n\n"
            
            # シートの内容を文字列に変換
            data = []
            for row in sheet.rows:
                row_data = [str(cell.value) if cell.value is not None else '' for cell in row]
                data.append(row_data)
            
            # 各列の最大幅を計算
            col_widths = {}
            for row in data:
                for col_idx, cell_value in enumerate(row):
                    col_widths[col_idx] = max(col_widths.get(col_idx, 0), len(cell_value))
            
            # 各行をフォーマットして出力
            for row in data:
                line = ""
                for col_idx, cell_value in enumerate(row):
                    if col_idx < len(row):
                        line += cell_value.ljust(col_widths[col_idx] + 2)
                output_text += line + "\n"
            output_text += "\n\n"
        
        return output_text
    
    finally:
        # 一時ファイルを削除
        os.unlink(excel_path)

def extract_text_with_mammoth(file_data):
    """mammothを使用してテキストを抽出する関数"""
    try:
        # mammothを使用してテキストを抽出
        result = mammoth.extract_raw_text(io.BytesIO(file_data))
        return result.value
    except Exception as e:
        return f"mammoth抽出エラー: {str(e)}"

def clean_text(text):
    """抽出されたテキストを整形する関数"""
    # 複数の空白行を1つにまとめる
    text = re.sub(r'\n\s*\n', '\n\n', text)
    # 行頭と行末の空白を削除
    text = "\n".join([line.strip() for line in text.split("\n")])
    return text

def process_docx_file(file_data):
    """Wordファイルからテキストを抽出し、整形して返す関数"""
    # mammothでテキストを抽出
    extracted_text = extract_text_with_mammoth(file_data)
    
    # テキストを整形
    cleaned_text = clean_text(extracted_text)
    
    return cleaned_text

load_dotenv(override=True)

# Azure Test Dify02
DIFY_API_KEY = "app-xxxxxxxxxx"
DIFY_BASE_URL = "http://xxx.xx.xx.xx/v1"
DIFY_USER = "sample-user"
proxies = {
  'http': 'http://proxy.xxxx.co.jp:8080',
  'https': 'http://proxy.xxxx.co.jp:8080',
}

def upload_file(file_data, file_name, file_type):
    target_url = f"{DIFY_BASE_URL}/files/upload"

    headers = {
        "Authorization": f"Bearer {DIFY_API_KEY}",
    }

    try:
        response = requests.post(
            target_url,
            headers=headers,
#            proxies=proxies,     # for No Proxy Comment Out
            files={"file": (file_name, file_data, file_type)},
            data={"user": DIFY_USER},
        )

        if response.status_code == 201:
            return response.json()
        else:
            st.error(f"アップロードエラー: {response.status_code}")
            st.write(response.text)  # エラー詳細を表示
            return None

    except Exception as e:
        st.error(f"予期しないエラーが発生しました: {str(e)}")
        return None

def run_workflow(file_id):
    target_url = f"{DIFY_BASE_URL}/workflows/run"
    headers = {
        "Authorization": f"Bearer {DIFY_API_KEY}",
        "Content-Type": "application/json"
    }

    input = {
        # Dify ワークフローの入力フィールド名と一致させる
        "doc": {
            "type": "document",
            "transfer_method": "local_file",
            "upload_file_id": file_id
        }   
    }

    payload = {
        "inputs": input,
        "response_mode": "blocking",
        "user": DIFY_USER
    }

    try:
#        response = requests.post(target_url, headers=headers,proxies=proxies, json=payload) # Use Proxy
        response = requests.post(target_url, headers=headers, json=payload) # No Proxy
        response.raise_for_status()
        return response.json()
    except requests.exceptions.RequestException as e:
        st.error(f"ワークフロー実行エラー: {str(e)}")
        st.write(e.response.text if hasattr(e, 'response') else "レスポンスの詳細はありません")
        return None

st.title("AI Technical Document Check")

uploaded_file = st.file_uploader(
    "ファイルをアップロードしてください",
    type=["jpg", "jpeg", "txt", "xlsx", "docx", "pdf"],
)

if uploaded_file is not None:
    file_data = uploaded_file.getvalue()  # バイトデータを取得
    file_type = uploaded_file.type
    file_name = uploaded_file.name
    
    col1, col2 = st.columns(2)
    with col1:
        st.write("ファイル名:", file_name)
        st.write("ファイルタイプ:", file_type)

    # ファイルタイプに応じた処理
    file_extension = get_file_extension(file_name).lower()
    processed_data = None
    processed_type = file_type
    
    if file_extension == ".xlsx":
        st.write("前処理:Excelファイルをテキスト化処理中...")
        processed_data = convert_with_openpyxl(file_data, file_name)
        processed_type = "text/plain"
        # プレビュー表示
        st.text_area("変換後のテキスト (プレビュー)", processed_data[:500] + "..." if len(processed_data) > 500 else processed_data, height=200)
    
    elif file_extension == ".docx":
        st.write("前処理:Wordファイルをテキスト化処理中...")
        processed_data = process_docx_file(file_data)
        processed_type = "text/plain"
        # プレビュー表示
        st.text_area("変換後のテキスト (プレビュー)", processed_data[:500] + "..." if len(processed_data) > 500 else processed_data, height=200)
    
    else:
        # その他のファイルはそのまま使用
        processed_data = file_data

    if st.button("ファイルをアップロード"):
        with st.spinner("ファイルをアップロード中..."):
            # ファイルタイプに応じてアップロード
            if file_extension in [".xlsx", ".docx"]:
                # 処理済みテキストをアップロード
                upload_response = upload_file(
                    processed_data.encode('utf-8') if isinstance(processed_data, str) else processed_data,
                    Path(file_name).stem + ".txt", 
                    "text/plain"
                )
            else:
                # 元のファイルをそのままアップロード
                upload_response = upload_file(file_data, file_name, file_type)
                
            if upload_response is not None:
                st.success(f'アップロード完了: ファイルID - {upload_response["id"]}')
                
                with st.spinner("ワークフローを実行中..."):
                    workflow_result = run_workflow(upload_response["id"])
                    
                    if workflow_result is not None:
                        st.subheader("ワークフロー実行結果")
                        
                        # NoneTypeエラーを回避するための修正
                        try:
                            result_text = workflow_result.get("data", {}).get("outputs", {}).get("text", "テキストが見つかりません")
                            # st.text_area("結果", result_text, height=400)
                            st.write("結果", result_text)
                        except (AttributeError, TypeError):
                            st.error("ワークフローの結果が予期せぬ形式です。")
                            st.write("ワークフロー応答:", workflow_result)
                        
                        # レスポンス全体を表示（デバッグ用）
                        with st.expander("API レスポンス詳細"):
                            st.json(workflow_result)
                    else:
                        st.error('ワークフロー実行結果が取得できませんでした。')
                        st.info('DifyサーバーとAPIキーが正しく設定されているか確認してください。')
            else:
                st.error('ファイルのアップロードに失敗しました。')