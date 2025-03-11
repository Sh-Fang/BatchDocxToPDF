import PyPDF2
import os

# ============================================================

def merge_pdfs(input_folder, output_pdf):
    # 获取文件夹中的所有 PDF 文件
    pdf_files = [f for f in os.listdir(input_folder) if f.lower().endswith(".pdf")]
    pdf_files.sort()  # 按文件名排序

    # 创建一个 PDF 合并对象
    pdf_merger = PyPDF2.PdfMerger()

    # 遍历所有 PDF 文件并将其合并
    for pdf_file in pdf_files:
        pdf_path = os.path.join(input_folder, pdf_file)
        pdf_merger.append(pdf_path)

    # 输出合并后的 PDF 文件
    pdf_merger.write(output_pdf)
    pdf_merger.close()

    print(f"PDFs have been merged into: {output_pdf}")

# ============================================================

# 设置两个文件夹路径
single_print_path = "./单页打印"
double_print_path = "./双面打印"

# 设置输出路径
output_folder = "merge_output"
if not os.path.exists(output_folder):
    os.makedirs(output_folder)

# 输出文件的全路径
single_print_pdf = os.path.join(output_folder, "单面打印.pdf")
double_print_pdf = os.path.join(output_folder, "双面打印.pdf")

# ============================================================

merge_pdfs(single_print_path, single_print_pdf)
merge_pdfs(double_print_path, double_print_pdf)
