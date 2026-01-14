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
# Limit file uploads to 50MB
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
    body { background: linear-gradient(180deg,#f8fafc,#ffffff); }
    .card { backdrop-filter: blur(6px); }
    .drop-zone--over { border-color: #4f46e5; background: #f5f3ff; }
  </style>
</head>
<body class="min-h-screen flex items-center justify-center p-6">
  <div class="max-w-4xl w-full grid grid-cols-1 md:grid-cols-2 gap-6">
    <div class="card p-6 rounded-2xl shadow-lg bg-white/80 border border-slate-100">
      <h1 class="text-2xl font-semibold mb-3">Slide Optimizer</h1>
      <p class="text-sm text-slate-600 mb-6">Convert PPTX or PDF slides into a printable handout.</p>

      <div id="dropZone" class="border-2 border-dashed border-slate-200 rounded-xl p-8 text-center cursor-pointer transition-all hover:border-indigo-400 mb-4">
        <div class="text-3xl mb-2">📄</div>
        <p class="text-sm font-medium text-slate-700" id="fileLabel">Click or drag PDF/PPTX here</p>
        <input type="file" id="fileInput" class="hidden" accept=".pdf,.pptx,.ppt" />
      </div>

      <div class="grid grid-cols-2 gap-4 mt-4">
        <div>
          <label class="block text-sm font-medium text-slate-700">Layout</label>
          <select id="slidesPerPage" class="mt-1 block w-full rounded-md border-slate-200 shadow-sm text-sm">
            <option value="1">1 Slide</option>
            <option value="2">2 Slides</option>
            <option value="4" selected>4 Slides</option>
            <option value="6">6 Slides</option>
          </select>
        </div>
        <div>
          <label class="block text-sm font-medium text-slate-700">Quality (DPI)</label>
          <input id="dpi" type="number" value="200" class="mt-1 block w-full rounded-md border-slate-200 shadow-sm text-sm" />
        </div>
      </div>

      <button id="processBtn" class="w-full mt-6 px-4 py-3 bg-indigo-600 text-white font-semibold rounded-xl shadow-md hover:bg-indigo-700 disabled:bg-slate-300 transition-all">
        Optimize & Download
      </button>

      <div id="status" class="mt-4 hidden text-center">
        <div class="inline-block animate-spin rounded-full h-4 w-4 border-b-2 border-indigo-600 mr-2"></div>
        <span class="text-sm text-slate-600">Processing file...</span>
      </div>
    </div>

    <div class="card p-6 rounded-2xl shadow-lg bg-white/80 border border-slate-100 flex flex-col justify-center">
      <h2 class="text-lg font-semibold mb-2">Instructions</h2>
      <ul class="text-sm text-slate-600 space-y-3">
        <li class="flex gap-2"><span>✅</span> Supports .pptx, .ppt, and .pdf</li>
        <li class="flex gap-2"><span>✅</span> Tiles multiple slides on one page</li>
        <li class="flex gap-2"><span>✅</span> Optimized for 8.5x11" paper</li>
      </ul>
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
        $('fileLabel').classList.add('text-indigo-600');
    }

    $('processBtn').onclick = async () => {
        if (!selectedFile) return alert("Select a file first");
        const formData = new FormData();
        formData.append('file', selectedFile);
        formData.append('slides_per_page', $('slidesPerPage').value);
        formData.append('dpi', $('dpi').value);

        $('status').classList.remove('hidden');
        $('processBtn').disabled = true;

        try {
            const response = await fetch('/optimize', { method: 'POST', body: formData });
            if (!response.ok) throw new Error("Conversion failed");
            const blob = await response.blob();
            const url = window.URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = "optimized_handout.pdf";
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
    pdf_document = fitz.open(input_pdf)
    total_slides = len(pdf_document)
    images = []
    zoom = dpi / 72
    mat = fitz.Matrix(zoom, zoom)
    
    for page in pdf_document:
        pix = page.get_pixmap(matrix=mat)
        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        images.append(img)
    pdf_document.close()

    page_width, page_height = letter
    s_p_p = int(slides_per_page)
    layouts = {1:(1,1), 2:(1,2), 4:(2,2), 6:(2,3)}
    cols, rows = layouts.get(s_p_p, (2,2))
    
    margin, gap = 36, 12
    cw = (page_width - (2*margin) - (cols-1)*gap) / cols
    ch = (page_height - (2*margin) - (rows-1)*gap) / rows
    
    img_w, img_h = images[0].size
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
            c.setStrokeColorRGB(0.8, 0.8, 0.8)
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
    filename = secure_filename(file.filename)
    input_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
    file.save(input_path)

    if filename.lower().endswith(('.pptx', '.ppt')):
        subprocess.run(['soffice', '--headless', '--convert-to', 'pdf', input_path, '--outdir', app.config['UPLOAD_FOLDER']], check=True)
        input_path = input_path.rsplit('.', 1)[0] + '.pdf'

    output_path = os.path.join(app.config['UPLOAD_FOLDER'], 'output.pdf')
    process_file(input_path, output_path, request.form.get('slides_per_page'), int(request.form.get('dpi')))
    return send_file(output_path, as_attachment=True)

if __name__ == '__main__':
    # Local dev fallback
    app.run(host='0.0.0.0', port=int(os.environ.get('PORT', 8080)))
