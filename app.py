import os
import sys
import tempfile
import threading

from flask import Flask, render_template, request, send_file, flash, redirect, jsonify, make_response
import openpyxl
from werkzeug.utils import secure_filename
from urllib.parse import quote


def get_resource_path(relative_path):
    """PyInstallerでパッケージ化された場合のリソースパスを取得"""
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.join(os.path.dirname(os.path.abspath(__file__)), relative_path)


app = Flask(__name__, template_folder=get_resource_path('templates'))
app.secret_key = os.urandom(24)
app.config['MAX_CONTENT_LENGTH'] = 5 * 1024 * 1024  # 5MB制限（Web版）

ALLOWED_EXTENSIONS = {'xlsm'}


def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


def remove_macro(input_path, output_path):
    """
    マクロ付きExcelファイル(.xlsm)を読み込み、
    マクロなしの.xlsxとして保存する

    openpyxlはマクロを保持しないため、.xlsxで保存するだけでマクロが除去される
    """
    # keep_vba=Falseでマクロを読み込まない（デフォルト動作）
    wb = openpyxl.load_workbook(input_path, keep_vba=False)
    wb.save(output_path)
    wb.close()
    return output_path


@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        if 'file' not in request.files:
            flash('ファイルが選択されていません', 'error')
            return redirect(request.url)

        file = request.files['file']

        if file.filename == '':
            flash('ファイルが選択されていません', 'error')
            return redirect(request.url)

        if not file or not allowed_file(file.filename):
            flash('対応しているファイル形式は .xlsm のみです', 'error')
            return redirect(request.url)

        temp_dir = tempfile.mkdtemp()
        input_filename = secure_filename(file.filename)
        input_path = os.path.join(temp_dir, input_filename)
        file.save(input_path)

        try:
            # 出力ファイル名: 元のファイル名 + (マクロ除去済).xlsx
            original_base_name = os.path.splitext(file.filename)[0]
            output_filename = f"{original_base_name}(マクロ除去済).xlsx"
            safe_output_filename = f"{os.path.splitext(input_filename)[0]}.xlsx"
            output_path = os.path.join(temp_dir, safe_output_filename)

            remove_macro(input_path, output_path)

            response = make_response(send_file(
                output_path,
                mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            ))
            # 日本語ファイル名対応: RFC 5987形式でエンコード
            encoded_filename = quote(output_filename)
            response.headers['Content-Disposition'] = \
                f"attachment; filename*=UTF-8''{encoded_filename}"
            return response
        except Exception as e:
            flash(f'処理中にエラーが発生しました: {str(e)}', 'error')
            return redirect(request.url)

    return render_template('index.html')


@app.route('/api/remove-macro', methods=['POST'])
def api_remove_macro():
    """Ajax用のマクロ除去API"""
    if 'file' not in request.files:
        return jsonify({'error': 'ファイルが選択されていません'}), 400

    file = request.files['file']

    if file.filename == '':
        return jsonify({'error': 'ファイルが選択されていません'}), 400

    if not file or not allowed_file(file.filename):
        return jsonify({'error': '対応しているファイル形式は .xlsm のみです'}), 400

    temp_dir = tempfile.mkdtemp()
    input_filename = secure_filename(file.filename)
    input_path = os.path.join(temp_dir, input_filename)
    file.save(input_path)

    try:
        # 出力ファイル名: 元のファイル名 + (マクロ除去済).xlsx
        original_base_name = os.path.splitext(file.filename)[0]
        output_filename = f"{original_base_name}(マクロ除去済).xlsx"
        safe_output_filename = f"{os.path.splitext(input_filename)[0]}.xlsx"
        output_path = os.path.join(temp_dir, safe_output_filename)

        remove_macro(input_path, output_path)

        response = make_response(send_file(
            output_path,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        ))
        # 日本語ファイル名対応: RFC 5987形式でエンコード
        encoded_filename = quote(output_filename)
        response.headers['Content-Disposition'] = \
            f"attachment; filename*=UTF-8''{encoded_filename}"
        return response
    except Exception as e:
        return jsonify({'error': f'処理中にエラーが発生しました: {str(e)}'}), 500


def start_flask_server():
    """Flaskサーバーをバックグラウンドで起動"""
    app.run(host='127.0.0.1', port=5001, debug=False, threaded=True, use_reloader=False)


def main():
    """メインエントリーポイント: pywebviewでデスクトップアプリとして起動"""
    try:
        import webview

        # Flaskサーバーをバックグラウンドスレッドで起動
        server_thread = threading.Thread(target=start_flask_server, daemon=True)
        server_thread.start()

        # サーバーの起動を待機
        import time
        time.sleep(1)

        # pywebviewでウィンドウを作成
        webview.create_window(
            'MacroRemover - Excelマクロ除去ツール',
            'http://127.0.0.1:5001',
            width=600,
            height=700,
            resizable=True,
            min_size=(450, 550)
        )
        webview.start()

    except ImportError:
        # pywebviewがインストールされていない場合はブラウザで開く
        import webbrowser

        print("pywebviewがインストールされていません。ブラウザで開きます。")
        print("アプリを終了するには Ctrl+C を押してください。")

        # サーバー起動前にブラウザを開く準備
        def open_browser():
            import time
            time.sleep(1.5)
            webbrowser.open('http://127.0.0.1:5001')

        browser_thread = threading.Thread(target=open_browser, daemon=True)
        browser_thread.start()

        # Flaskサーバーを起動（メインスレッド）
        app.run(host='127.0.0.1', port=5001, debug=False, threaded=True)


if __name__ == '__main__':
    main()
