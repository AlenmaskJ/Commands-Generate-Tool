from flask import Flask, render_template_string, request, send_file, flash, redirect, url_for
import openpyxl
import xlrd
import re
import os
import tempfile
from werkzeug.utils import secure_filename

# 初始化Flask应用
app = Flask(__name__)
app.secret_key = 'command_generator_2026'  # 用于flash提示
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 限制上传文件最大16MB

# 前端页面模板（和之前一致，无需修改）
HTML_TEMPLATE = '''
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>通用命令生成工具</title>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
    <style>
        .container {
            max-width: 800px;
            margin-top: 50px;
        }
        .result-area {
            margin-top: 20px;
            padding: 15px;
            border: 1px solid #ddd;
            border-radius: 5px;
            min-height: 200px;
            background-color: #f8f9fa;
            white-space: pre-wrap;
            font-family: monospace;
        }
        .form-label {
            font-weight: bold;
        }
    </style>
</head>
<body>
    <div class="container">
        <h2 class="text-center mb-4">通用命令生成工具</h2>

        {% with messages = get_flashed_messages() %}
          {% if messages %}
            {% for message in messages %}
              <div class="alert alert-warning alert-dismissible fade show" role="alert">
                {{ message }}
                <button type="button" class="btn-close" data-bs-dismiss="alert" aria-label="Close"></button>
              </div>
            {% endfor %}
          {% endif %}
        {% endwith %}

        <form method="POST" enctype="multipart/form-data">
            <div class="mb-3">
                <label for="template" class="form-label">命令模板（变量用{变量名}标识，例：Mod Bscex : MSISDN={x},SMSRouterID={y};）</label>
                <textarea class="form-control" id="template" name="template" rows="3" required placeholder="请输入命令模板，如：Mod Bscex : MSISDN={x},SMSRouterID={y};"></textarea>
                <div class="form-text">变量必须用大括号包裹，如{x}、{y}、{z}，Excel列名需与变量名完全一致</div>
            </div>

            <div class="mb-3">
                <label for="excel_file" class="form-label">上传Excel文件（列名对应模板中的变量）</label>
                <input class="form-control" type="file" id="excel_file" name="excel_file" accept=".xlsx,.xls" required>
                <div class="form-text">仅支持.xlsx/.xls格式，Excel中每列对应一个变量，列名需和模板中的变量名一致（如模板有{x}，Excel需有列名为x的列）</div>
            </div>

            <button type="submit" class="btn btn-primary">生成命令</button>
        </form>

        {% if commands %}
            <h4 class="mt-4">生成结果</h4>
            <div class="result-area" id="commandResult">{{ commands }}</div>
            <div class="mt-3">
                <a href="{{ url_for('download_file', filename=output_file) }}" class="btn btn-success">下载命令文件（TXT）</a>
            </div>
        {% endif %}
    </div>

    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/js/bootstrap.bundle.min.js"></script>
</body>
</html>
'''


# 提取模板中的所有变量（和之前一致）
def extract_variables(template):
    pattern = r'\{(\w+)\}'
    variables = re.findall(pattern, template)
    return list(set(variables))


# 轻量版Excel读取+命令生成（替换pandas）
def generate_commands_from_template(template, excel_path):
    variables = extract_variables(template)
    if not variables:
        raise ValueError("命令模板中未检测到有效变量（变量需用{变量名}格式，如{x}）")

    commands = []
    # 处理xlsx格式
    if excel_path.endswith('.xlsx'):
        wb = openpyxl.load_workbook(excel_path, data_only=True)
        ws = wb.active
        # 获取表头
        headers = [cell.value for cell in ws[1]]
        # 检查缺失变量
        missing_vars = [var for var in variables if var not in headers]
        if missing_vars:
            raise ValueError(f"Excel缺少变量列：{', '.join(missing_vars)}")
        # 遍历数据行
        for row in ws.iter_rows(min_row=2, values_only=True):
            row_dict = dict(zip(headers, row))
            if any(row_dict.get(var) is None for var in variables):
                continue
            # 替换变量
            command = template
            for var in variables:
                command = command.replace(f'{{{var}}}', str(row_dict[var]))
            commands.append(command)
    # 处理xls格式
    elif excel_path.endswith('.xls'):
        wb = xlrd.open_workbook(excel_path)
        ws = wb.sheet_by_index(0)
        headers = [ws.cell_value(0, col) for col in range(ws.ncols)]
        missing_vars = [var for var in variables if var not in headers]
        if missing_vars:
            raise ValueError(f"Excel缺少变量列：{', '.join(missing_vars)}")
        # 遍历数据行
        for row_idx in range(1, ws.nrows):
            row = [ws.cell_value(row_idx, col) for col in range(ws.ncols)]
            row_dict = dict(zip(headers, row))
            if any(row_dict.get(var) is None for var in variables):
                continue
            command = template
            for var in variables:
                command = command.replace(f'{{{var}}}', str(row_dict[var]))
            commands.append(command)
    else:
        raise ValueError("仅支持.xlsx/.xls格式的Excel文件")

    if not commands:
        raise ValueError("未生成任何有效命令，请检查Excel数据")

    return commands


# 保存命令到文件（和之前一致）
def save_commands_to_file(commands):
    temp_dir = tempfile.gettempdir()
    file_path = os.path.join(temp_dir, "generated_commands.txt")
    with open(file_path, 'w', encoding='utf-8') as f:
        f.write('\n'.join(commands))
    return os.path.basename(file_path), file_path


# 主页面路由（和之前一致）
@app.route('/', methods=['GET', 'POST'])
def index():
    commands = None
    output_file = None

    if request.method == 'POST':
        try:
            template = request.form['template'].strip()
            excel_file = request.files['excel_file']

            if not excel_file.filename:
                flash("请选择要上传的Excel文件")
                return redirect(url_for('index'))

            filename = secure_filename(excel_file.filename)
            temp_excel_path = os.path.join(tempfile.gettempdir(), filename)
            excel_file.save(temp_excel_path)

            command_list = generate_commands_from_template(template, temp_excel_path)
            commands = '\n'.join(command_list)

            output_file, _ = save_commands_to_file(command_list)
            os.remove(temp_excel_path)

        except ValueError as e:
            flash(f"错误：{str(e)}")
        except Exception as e:
            flash(f"系统错误：{str(e)}")

    return render_template_string(HTML_TEMPLATE, commands=commands, output_file=output_file)


# 下载文件路由（和之前一致）
@app.route('/download/<filename>')
def download_file(filename):
    file_path = os.path.join(tempfile.gettempdir(), filename)
    if os.path.exists(file_path):
        return send_file(file_path, as_attachment=True, download_name="generated_commands.txt", mimetype='text/plain')
    else:
        flash("下载文件不存在")
        return redirect(url_for('index'))


# 启动应用
if __name__ == '__main__':
    print("通用命令生成工具已启动，访问 http://127.0.0.1:5000 使用")
    app.run(host='0.0.0.0', port=5000, debug=False)