import os
import shutil
import webbrowser
import socket
import time
from pathlib import Path
from flask import Flask, render_template, request, send_file, flash, redirect, url_for
from werkzeug.utils import secure_filename
from pypdf import PdfWriter, PdfReader

app = Flask(__name__)
app.secret_key = os.urandom(24)
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024  # 50MB 제한

# 업로드 및 다운로드 디렉토리 설정
UPLOAD_FOLDER = 'uploads'
DOWNLOAD_FOLDER = 'downloads'
ALLOWED_EXTENSIONS = {'pdf'}

# 디렉토리 생성
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(DOWNLOAD_FOLDER, exist_ok=True)

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['DOWNLOAD_FOLDER'] = DOWNLOAD_FOLDER


def allowed_file(filename):
    """파일 확장자 검증"""
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


def extract_pages(input_path, output_path, pages):
    """
    PDF 파일에서 지정한 페이지를 추출합니다.
    
    Args:
        input_path: 입력 PDF 파일 경로
        output_path: 출력 PDF 파일 경로
        pages: 추출할 페이지 번호 리스트 (1부터 시작)
    """
    reader = PdfReader(input_path)
    writer = PdfWriter()
    
    total_pages = len(reader.pages)
    extracted_count = 0
    
    for page_num in pages:
        if 1 <= page_num <= total_pages:
            writer.add_page(reader.pages[page_num - 1])
            extracted_count += 1
    
    if extracted_count == 0:
        raise ValueError("추출할 수 있는 페이지가 없습니다.")
    
    with open(output_path, 'wb') as output_file:
        writer.write(output_file)
    
    return extracted_count


def merge_pdfs(input_paths, output_path):
    """
    여러 PDF 파일을 하나로 병합합니다.
    
    Args:
        input_paths: 병합할 PDF 파일 경로 리스트
        output_path: 출력 PDF 파일 경로
    """
    writer = PdfWriter()
    total_pages = 0
    
    for pdf_path in input_paths:
        if not os.path.exists(pdf_path):
            continue
        
        reader = PdfReader(pdf_path)
        for page in reader.pages:
            writer.add_page(page)
            total_pages += 1
    
    if total_pages == 0:
        raise ValueError("병합할 페이지가 없습니다.")
    
    with open(output_path, 'wb') as output_file:
        writer.write(output_file)
    
    return total_pages


@app.route('/')
def index():
    """메인 페이지"""
    return render_template('index.html')


@app.route('/extract', methods=['POST'])
def extract():
    """PDF 페이지 추출 처리"""
    try:
        # 파일 검증
        if 'file' not in request.files:
            flash('파일을 선택해주세요.', 'error')
            return redirect(url_for('index'))
        
        file = request.files['file']
        pages_input = request.form.get('pages', '').strip()
        
        if file.filename == '':
            flash('파일을 선택해주세요.', 'error')
            return redirect(url_for('index'))
        
        if not allowed_file(file.filename):
            flash('PDF 파일만 업로드할 수 있습니다.', 'error')
            return redirect(url_for('index'))
        
        if not pages_input:
            flash('추출할 페이지 번호를 입력해주세요.', 'error')
            return redirect(url_for('index'))
        
        # 파일 저장
        filename = secure_filename(file.filename)
        input_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        file.save(input_path)
        
        try:
            # 페이지 번호 파싱
            pages = []
            for page_arg in pages_input.replace(',', ' ').split():
                if '-' in page_arg:
                    start, end = map(int, page_arg.split('-'))
                    pages.extend(range(start, end + 1))
                else:
                    pages.append(int(page_arg))
            
            pages = sorted(set(pages))  # 중복 제거 및 정렬
            
            # 페이지 추출
            output_filename = f"extracted_{Path(filename).stem}.pdf"
            output_path = os.path.join(app.config['DOWNLOAD_FOLDER'], output_filename)
            
            extracted_count = extract_pages(input_path, output_path, pages)
            
            # 임시 파일 삭제
            os.remove(input_path)
            
            flash(f'성공: {extracted_count}개의 페이지가 추출되었습니다.', 'success')
            return redirect(url_for('download', filename=output_filename))
            
        except ValueError as e:
            os.remove(input_path)
            flash(str(e), 'error')
            return redirect(url_for('index'))
        except Exception as e:
            if os.path.exists(input_path):
                os.remove(input_path)
            flash(f'오류 발생: {str(e)}', 'error')
            return redirect(url_for('index'))
            
    except Exception as e:
        flash(f'오류 발생: {str(e)}', 'error')
        return redirect(url_for('index'))


@app.route('/merge', methods=['POST'])
def merge():
    """PDF 파일 병합 처리"""
    try:
        # 파일 검증
        if 'files' not in request.files:
            flash('파일을 선택해주세요.', 'error')
            return redirect(url_for('index'))
        
        files = request.files.getlist('files')
        
        if not files or files[0].filename == '':
            flash('최소 하나의 파일을 선택해주세요.', 'error')
            return redirect(url_for('index'))
        
        # 파일 저장
        input_paths = []
        saved_files = []
        
        try:
            for file in files:
                if file.filename == '':
                    continue
                
                if not allowed_file(file.filename):
                    flash(f'{file.filename}: PDF 파일만 업로드할 수 있습니다.', 'error')
                    # 이미 저장한 파일들 정리
                    for path in saved_files:
                        if os.path.exists(path):
                            os.remove(path)
                    return redirect(url_for('index'))
                
                filename = secure_filename(file.filename)
                input_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
                file.save(input_path)
                input_paths.append(input_path)
                saved_files.append(input_path)
            
            if not input_paths:
                flash('유효한 파일이 없습니다.', 'error')
                return redirect(url_for('index'))
            
            # 파일 병합
            output_filename = "merged.pdf"
            output_path = os.path.join(app.config['DOWNLOAD_FOLDER'], output_filename)
            
            total_pages = merge_pdfs(input_paths, output_path)
            
            # 임시 파일 삭제
            for path in input_paths:
                if os.path.exists(path):
                    os.remove(path)
            
            flash(f'성공: {len(input_paths)}개의 파일이 병합되었습니다. (총 {total_pages}페이지)', 'success')
            return redirect(url_for('download', filename=output_filename))
            
        except ValueError as e:
            for path in saved_files:
                if os.path.exists(path):
                    os.remove(path)
            flash(str(e), 'error')
            return redirect(url_for('index'))
        except Exception as e:
            for path in saved_files:
                if os.path.exists(path):
                    os.remove(path)
            flash(f'오류 발생: {str(e)}', 'error')
            return redirect(url_for('index'))
            
    except Exception as e:
        flash(f'오류 발생: {str(e)}', 'error')
        return redirect(url_for('index'))


@app.route('/download/<filename>')
def download(filename):
    """파일 다운로드"""
    try:
        file_path = os.path.join(app.config['DOWNLOAD_FOLDER'], filename)
        if not os.path.exists(file_path):
            flash('파일을 찾을 수 없습니다.', 'error')
            return redirect(url_for('index'))
        
        return send_file(file_path, as_attachment=True, download_name=filename)
    except Exception as e:
        flash(f'다운로드 오류: {str(e)}', 'error')
        return redirect(url_for('index'))


def is_port_available(port):
    """포트가 사용 가능한지 확인"""
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        try:
            s.bind(('localhost', port))
            return True
        except OSError:
            return False


def find_available_port(start_port=5000, max_port=5010):
    """사용 가능한 포트 찾기"""
    for port in range(start_port, max_port + 1):
        if is_port_available(port):
            return port
    return None


if __name__ == '__main__':
    # 사용 가능한 포트 찾기
    port = find_available_port(5000, 5010)
    
    if port is None:
        print("오류: 5000-5010 포트 범위에서 사용 가능한 포트를 찾을 수 없습니다.")
        port = 5000  # 기본값 사용 (오류 발생 가능)
    
    url = f'http://localhost:{port}'
    
    print(f"\n{'='*50}")
    print(f"PDF 처리 웹앱이 시작되었습니다!")
    print(f"브라우저에서 다음 주소로 접속하세요:")
    print(f"  {url}")
    print(f"{'='*50}\n")
    
    # 1초 후 브라우저 자동 열기
    def open_browser():
        time.sleep(1)
        webbrowser.open(url)
    
    import threading
    threading.Thread(target=open_browser, daemon=True).start()
    
    # localhost로 변경 (0.0.0.0은 외부 접속용)
    app.run(debug=True, host='127.0.0.1', port=port, use_reloader=False)
