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
  </style>
</head>
<body class="min-h-screen flex items-center justify-center p-6">
  <div class="max-w-4xl w-full grid grid-cols-1 md:grid-cols-2 gap-8">
    
    <div class="card p-8 rounded-lg shadow-2xl">
      <h1 class="text-2xl font-light mb-1 tracking-tight">Slide <span class="accent-brown italic">Optimizer</span></h1>
      <p class="text-xs text-slate-500 mb-8 uppercase tracking-widest">Handout Generation Tool</p>

      <div id="dropZone" class="border border-dashed border-zinc-700 rounded-md p-10 text-center cursor-pointer transition-all mb-6">
        <p class="text-sm text-zinc-400" id="fileLabel">Drop PDF or PowerPoint</p>
        <input type="file" id="fileInput" class="hidden" accept=".pdf,.pptx,.ppt" />
      </div>

      <div class="space-y-4">
        <div>
          <label class="block text-xs font-semibold text-zinc-500 uppercase mb-1">Custom Output Name</label>
          <input id="outName" type="text" placeholder="optimized_handout" class="input-field w-full p-3 rounded text-sm" />
        </div>

        <div class="grid grid-cols-2 gap-4">
          <div>
            <label class="block text-xs font-semibold text-zinc-500 uppercase mb-1">Layout</label>
            <select id="slidesPerPage" class="input-field w-full p-3 rounded text-sm appearance-none">
              <option value="1">1 Slide</option>
              <option value="2">2 Slides</option>
              <option value="4" selected>4 Slides</option>
              <option value="6">6 Slides</option>
            </select>
          </div>
          <div>
            <label class="block text-xs font-semibold text-zinc-500 uppercase mb-1">DPI (Max 300)</label>
            <input id="dpi" type="number" value="200" max="300" class="input-field w-full p-3 rounded text-sm" />
          </div>
        </div>
      </div>

      <button id="processBtn" class="btn-primary w-full mt-8 py-4 rounded font-medium text-sm tracking-widest uppercase">
        Process Document
      </button>

      <div id="status" class="mt-6 hidden text-center">
        <span class="text-xs italic text-zinc-500 animate-pulse">Processing engine running...</span>
      </div>
    </div>

    <div class="flex flex-col justify-center p-4 border-l border-zinc-800">
      <h2 class="text-sm font-semibold uppercase tracking-widest mb-6 accent-brown">Capabilities</h2>
      <ul class="space-y-6">
        <li class="flex items-start gap-4">
          <span class="text-zinc-600 text-xs">01</span>
          <div>
            <p class="text-sm font-medium">Multi-Format Support</p>
            <p class="text-xs text-zinc-500">Native conversion for .pptx and .pdf using LibreOffice headless.</p>
          </div>
        </li>
        <li class="flex items-start gap-4">
          <span class="text-zinc-600 text-xs">02</span>
          <div>
            <p class="text-sm font-medium">Print Optimization</p>
            <p class="text-xs text-zinc-500">Standard 8.5x11" output with smart scaling to eliminate white space.</p>
          </div>
        </li>
        <li class="flex items-start gap-4">
          <span class="text-zinc-600 text-xs">03</span>
          <div>
            <p class="text-sm font-medium">Safety Bounds</p>
            <p class="text-xs text-zinc-500">Auto-capped quality to ensure fast processing and server stability.</p>
          </div>
        </li>
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
        $('fileLabel').style.color = "#a39081";
    }

    $('processBtn').onclick = async () => {
        if (!selectedFile) return alert("Select a file");
        
        let dpiVal = parseInt($('dpi').value);
        if (dpiVal > 300) { alert("DPI capped at 300 for stability."); dpiVal = 300; }

        const formData = new FormData();
        formData.append('file', selectedFile);
        formData.append('slides_per_page', $('slidesPerPage').value);
        formData.append('dpi', dpiVal);
        formData.append('out_name', $('outName').value || "optimized_handout");

        $('status').classList.remove('hidden');
        $('processBtn').disabled = true;

        try {
            const response = await fetch('/optimize', { method: 'POST', body: formData });
            if (!response.ok) throw new Error("Processing error");
            
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
    # Enforce DPI cap on backend
    dpi = min(int(dpi), 300)
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
            c.setStrokeColorRGB(0.3, 0.3, 0.3) # Darker borders for minimalist look
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
    custom_name = secure_filename(request.form.get('out_name', 'output'))
    if not custom_name: custom_name = "output"
    
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
