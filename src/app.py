from flask import Flask, render_template, request, send_file, redirect, url_for
import os
import tempfile
from pdf2docx import Converter
import pymupdf
import traceback

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 100 * 1024 * 1024  # 限制上传文件大小为100MB

# 确保临时目录存在
if not os.path.exists('temp'):
    os.makedirs('temp')

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/convert', methods=['POST'])
def convert():
    if 'pdf_file' not in request.files:
        return redirect(request.url)
    
    file = request.files['pdf_file']
    if file.filename == '':
        return redirect(request.url)
    
    if file and file.filename.endswith('.pdf'):
        # 保存上传的PDF文件
        pdf_path = os.path.join(os.getcwd(), 'temp', file.filename)
        file.save(pdf_path)
        
        # 获取用户选择的页码
        start_page = request.form.get('start_page', 1, type=int)
        end_page = request.form.get('end_page', None)
        if end_page:
            end_page = int(end_page)
        
        # 转换PDF到Word
        docx_filename = file.filename.rsplit('.', 1)[0] + '.docx'
        docx_path = os.path.join(os.getcwd(), 'temp', docx_filename)
        
        try:
            cv = Converter(pdf_path)
            cv.convert(docx_path, start=start_page-1, end=end_page)  # start from 0
            cv.close()
            
            # 提供文件下载
            return send_file(os.path.abspath(docx_path), as_attachment=True)
        except pymupdf.mupdf.FzErrorArgument as e:
            if "cannot find builtin font with name 'Arial'" in str(e):
                # 特殊处理字体错误
                return render_template('error.html'), 500
            else:
                # 其他PyMuPDF错误
                traceback.print_exc()
                return render_template('error.html'), 500
        except Exception as e:
            # 其他所有错误
            traceback.print_exc()
            return render_template('error.html'), 500
    
    return redirect(url_for('index'))

if __name__ == '__main__':
    app.run(debug=True)