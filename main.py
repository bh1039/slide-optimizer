#!/usr/bin/env python3
from flask import Flask, render_template_string, request, send_file, jsonify
import fitz  # PyMuPDF
from PIL import Image
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from reportlab.lib.utils import ImageReader
import os
import tempfile
import subprocess
from werkzeug.utils import secure_filename

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024
app.config['UPLOAD_FOLDER'] = tempfile.gettempdir()

HTML_TEMPLATE = """
<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width,initial-scale=1" />
  <title>Slide Optimizer Pro</title>
  <script src="https://cdn.tailwindcss.com"></script>
  <style>
    body { background-color: #1a1c1a; color: #e2e8f0; font-family: 'Inter', sans-serif; }
    .card { background: #242624; border: 1px solid #323532; }
    .input-field { background: #1a1c1a; border: 1px solid #3f423f; color: #e2e8f0; }
    .input-field:focus { border-color: #5d6d5d; outline: none; }
    .btn-primary { background: #3e4a3e; color: #d1d5db; transition: all 0.3s ease; }
    .btn-primary:hover { background: #4d5d4d; color: #ffffff; }
    .drop-zone--over { border-color: #8b9a8b; background: #2a2d2a; }
    .accent-brown { color: #a39081; }
    select { appearance: none; background-image: url("data:image/svg+xml;charset=UTF-8,%3csvg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 24 24' fill='none' stroke='currentColor' stroke-width='2' stroke-linecap='round' stroke-linejoin='round'%3e%3cpolyline points='6 9 12 15 18 9'%3e%3c/polyline%3e%3c/svg%3e"); background-repeat: no-repeat; background-position: right 1rem center; background-size: 1em; }
  </style>
</head>
<body class="min-h-screen flex items-center justify-center p-6">
  <div class="max-w-4xl w-full grid grid-cols-1 md:grid-cols-2 gap-8">
    
    <div class="card p-8 rounded-lg shadow-2xl">
      <h1 class="text-2xl font-light mb-1 tracking-tight text-zinc-100">Slide <span class="accent-brown italic font-serif">Optimizer</span></h1>
      <p class="text-[10px] text-zinc-500 mb-8 uppercase tracking-[0.2em] font-bold">Minimalist Handout Engine</p>

      <div id="dropZone" class="border border-dashed border-zinc-800 rounded-md p-10 text-center cursor-pointer transition-all mb-6 hover:border-zinc-600">
        <p class="text-xs text-zinc-500" id="fileLabel">Attach PDF or PowerPoint Source</p>
        <input type="file" id="fileInput" class="hidden" accept=".pdf,.pptx,.ppt" />
      </div>

      <div class="space-y-5">
        <div>
          <label class="block text-[10px] font-bold text-zinc-600 uppercase mb-2 tracking-widest">Filename</label>
          <input id="outName" type="text" placeholder="optimized_handout" class="input-field w-full p-3 rounded text-sm placeholder:text-zinc-700" />
        </div>

        <div class="grid grid-cols-2 gap-4">
          <div>
            <label class="block text-[10px] font-bold text-zinc-600 uppercase mb-2 tracking-widest">Layout</label>
            <select id="slidesPerPage" class="input-field w-full p-3 rounded text-sm">
              <option value="auto" selected>Auto-Detect</option>
              <option value="1">1 Slide</option>
              <option value="2">2 Slides</option>
              <option value="4">4 Slides</option>
              <option value="6">6 Slides</option>
            </select>
          </div>
          <div>
            <label class="block text-[10px] font-bold text-zinc-600 uppercase mb-2 tracking-widest">Quality (DPI MAX 300)</label>
            <input id="dpi" type="number" value="200" max="300" class="input-field w-full p-3 rounded text-sm" />
          </div>
        </div>
      </div>

      <button id="processBtn" class="btn-primary w-full mt-10 py-4 rounded text-[11px] font-bold tracking-[0.25em] uppercase shadow-lg">
        Begin Processing
      </button>

      <div id="status" class="mt-6 hidden text-center">
        <span class="text-[10px] uppercase tracking-widest text-zinc-600 animate-pulse">Reconstructing document...</span>
      </div>
    </div>

    <div class="flex flex-col justify-center p-4">
      <div class="space-y-8 max-w-xs mx-auto">
        <div class="border-l-2 border-zinc-800 pl-4">
          <h3 class="text-xs font-bold uppercase tracking-widest text-zinc-400 mb-2">Automated Logic</h3>
          <p class="text-xs text-zinc-500 leading-relaxed">System analyzes slide dimensions to choose between a 2-up landscape or 4-up portrait grid automatically.</p>
        </div>
        <div class="border-l-2 border-zinc-800 pl-4">
          <h3 class="text-xs font-bold uppercase tracking-widest text-zinc-400 mb-2">Constraints</h3>
          <p class="text-xs text-zinc-500 leading-relaxed">Input limited to 50MB. DPI is server-capped at 300 to ensure rendering stability.</p>
        </div>
      </div>
    </div>
  </div>

  <script>
    const $ = id => document.getElementById(id);
    const dropZone = $('dropZone');
    const fileInput = $('fileInput');
    let selectedFile = null;

    dropZone.onclick = () => fileInput.click();
    fileInput.onchange = (e) => { if (e.target.files.length) handleFile(e.target.files[0]); };

    function handleFile(file) {
        selectedFile = file;
        $('fileLabel').textContent = file.name;
        $('fileLabel').className = "text-xs accent-brown italic";
    }

    $('processBtn').onclick = async () => {
        if (!selectedFile) return alert("Source file missing");
        
        const dpiVal = Math.min(parseInt($('dpi').value) || 200, 300);
        const formData = new FormData();
        formData.append('file', selectedFile);
        formData.append('slides_per_page', $('slidesPerPage').value);
        formData.append('dpi', dpiVal);
        formData.append('out_name', $('outName').value || "optimized_handout");

        $('status').classList.remove('hidden');
        $('processBtn').disabled = true;

        try {
            const response = await fetch('/optimize', { method: 'POST', body: formData });
            if (!response.ok) throw new Error("Processing failure");
            
            const blob = await response.blob();
            const url = window.URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = ($('outName').value || "optimized_handout") + ".pdf";
            a.click();
        } catch (err) { alert(err.message); }
        finally {
            $('status').classList.add('hidden');
            $('processBtn').disabled = false;
        }
    };
  </script>
</body>
</html>
"""

def process_file(input_pdf, output_pdf, slides_per_page, dpi):
    dpi = min(int(dpi), 300)
    pdf_document = fitz.open(input_pdf)
    total_slides = len(pdf_document)
    
    # Render first page to determine orientation for "Auto" mode
    zoom = dpi / 72
    mat = fitz.Matrix(zoom, zoom)
    images = []
    for page in pdf_document:
        pix = page.get_pixmap(matrix=mat)
        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        images.append(img)
    
    img_w, img_h = images[0].size
    is_landscape = img_w > img_h
    pdf_document.close()

    # Automatic Logic
    if slides_per_page == 'auto':
        s_p_p = 2 if is_landscape else 4
    else:
        s_p_p = int(slides_per_page)

    page_width, page_height = letter
    layouts = {1:(1,1), 2:(1,2), 4:(2,2), 6:(2,3)}
    cols, rows = layouts.get(s_p_p, (2,2))
    
    margin, gap = 36, 12
    cw = (page_width - (2*margin) - (cols-1)*gap) / cols
    ch = (page_height - (2*margin) - (rows-1)*gap) / rows
    
    scale = min(cw/img_w, ch/img_h)
    sw, sh = img_w*scale, img_h*scale

    c = canvas.Canvas(output_pdf, pagesize=letter)
    idx = 0
    while idx < total_slides:
        for j in range(s_p_p):
            if idx >= total_slides: break
            col, row = j % cols, j // cols
            x = margin + col*(cw+gap) + (cw-sw)/2
            y = page_height - margin - (row+1)*ch - row*gap + (ch-sh)/2
            c.drawImage(ImageReader(images[idx]), x, y, width=sw, height=sh)
            c.setStrokeColorRGB(0.2, 0.2, 0.2)
            c.rect(x, y, sw, sh)
            idx += 1
        c.showPage()
    c.save()

@app.route('/')
def index():
    return render_template_string(HTML_TEMPLATE)

@app.route('/optimize', methods=['POST'])
def optimize():
    file = request.files.get('file')
    custom_name = secure_filename(request.form.get('out_name', 'optimized_handout'))
    filename = secure_filename(file.filename)
    input_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
    file.save(input_path)

    if filename.lower().endswith(('.pptx', '.ppt')):
        subprocess.run(['soffice', '--headless', '--convert-to', 'pdf', input_path, '--outdir', app.config['UPLOAD_FOLDER']], check=True)
        input_path = input_path.rsplit('.', 1)[0] + '.pdf'

    output_path = os.path.join(app.config['UPLOAD_FOLDER'], f"{custom_name}.pdf")
    process_file(input_path, output_path, request.form.get('slides_per_page'), request.form.get('dpi'))
    return send_file(output_path, as_attachment=True, download_name=f"{custom_name}.pdf")

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 8080))
    app.run(host='0.0.0.0', port=port)
