from ngx_report import generate_files_for_today

xlsx, docx, basename = generate_files_for_today()
open(f"/var/reports/{basename}.xlsx", "wb").write(xlsx)
open(f"/var/reports/{basename}.docx", "wb").write(docx)
