from flask import Flask, send_file, render_template_string, request
from io import BytesIO
from ngx_report import generate_files_for_today

app = Flask(__name__)

HTML = """
<!doctype html>
<title>NGX Daily Equity Summary</title>
<h1>NGX Daily Equity Summary</h1>
<form method="post" action="/generate">
  <button type="submit">Generate Today’s Report</button>
</form>
"""

@app.get("/")
def home():
    return render_template_string(HTML)

@app.post("/generate")
def generate():
    xlsx, docx, basename = generate_files_for_today()

    # Simple approach: return the Excel, and you can add a second button for Word
    # Better: zip both and download one file. For now, choose based on a query param.
    kind = request.args.get("kind", "xlsx")
    if kind == "docx":
        return send_file(BytesIO(docx), as_attachment=True, download_name=f"{basename}.docx",
                         mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
    return send_file(BytesIO(xlsx), as_attachment=True, download_name=f"{basename}.xlsx",
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000)
