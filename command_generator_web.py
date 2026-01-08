from flask import Flask, render_template_string, request, send_file, flash, redirect, url_for
import pandas as pd
import re
import os
import tempfile
from werkzeug.utils import secure_filename

# 初始化Flask应用
app = Flask(__name__)
app.secret_key = 'command_generator_2026'  # 用于flash提示
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 限制上传文件最大16MB

# 前端页面模板（整合HTML+CSS+JS）
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


# 提取模板中的所有变量（如从{x},{y}提取['x','y']）
def extract_variables(template):
    pattern = r'\{(\w+)\}'  # 匹配{变量名}格式，变量名仅包含字母/数字/下划线
    variables = re.findall(pattern, template)
    return list(set(variables))  # 去重


# 生成命令的核心函数
def generate_commands_from_template(template, excel_path):
    # 提取模板中的变量
    variables = extract_variables(template)
    if not variables:
        raise ValueError("命令模板中未检测到有效变量（变量需用{变量名}格式，如{x}）")

    # 读取Excel文件
    try:
        # 自动适配xls/xlsx格式
        if excel_path.endswith('.xls'):
            df = pd.read_excel(excel_path, engine='xlrd')
        else:
            df = pd.read_excel(excel_path, engine='openpyxl')
    except Exception as e:
        raise ValueError(f"读取Excel文件失败：{str(e)}")

    # 检查Excel列名是否包含所有变量
    missing_vars = [var for var in variables if var not in df.columns]
    if missing_vars:
        raise ValueError(f"Excel文件缺少以下变量列：{', '.join(missing_vars)}，请检查Excel列名是否与模板变量一致")

    # 遍历每行生成命令
    commands = []
    for index, row in df.iterrows():
        # 跳过空值行（任意变量为空则跳过）
        row_vars = {var: row[var] for var in variables}
        if any(pd.isna(val) for val in row_vars.values()):
            continue

        # 替换模板中的变量
        command = template
        for var, val in row_vars.items():
            command = command.replace(f'{{{var}}}', str(val))
        commands.append(command)

    if not commands:
        raise ValueError("未生成任何有效命令，请检查Excel数据是否为空或匹配")

    return commands


# 保存命令到临时文件
def save_commands_to_file(commands):
    temp_dir = tempfile.gettempdir()
    file_path = os.path.join(temp_dir, "generated_commands.txt")
    with open(file_path, 'w', encoding='utf-8') as f:
        f.write('\n'.join(commands))
    return os.path.basename(file_path), file_path


# 主页面路由
@app.route('/', methods=['GET', 'POST'])
def index():
    commands = None
    output_file = None

    if request.method == 'POST':
        try:
            # 获取表单数据
            template = request.form['template'].strip()
            excel_file = request.files['excel_file']

            # 校验文件
            if not excel_file.filename:
                flash("请选择要上传的Excel文件")
                return redirect(url_for('index'))

            # 保存上传的Excel到临时文件
            filename = secure_filename(excel_file.filename)
            temp_excel_path = os.path.join(tempfile.gettempdir(), filename)
            excel_file.save(temp_excel_path)

            # 生成命令
            command_list = generate_commands_from_template(template, temp_excel_path)
            commands = '\n'.join(command_list)

            # 保存命令到文件，用于下载
            output_file, _ = save_commands_to_file(command_list)

            # 删除临时Excel文件
            os.remove(temp_excel_path)

        except ValueError as e:
            flash(f"错误：{str(e)}")
        except Exception as e:
            flash(f"系统错误：{str(e)}")

    return render_template_string(HTML_TEMPLATE, commands=commands, output_file=output_file)


# 下载文件路由
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
    # 安装依赖提示
    print("=" * 50)
    print("使用前请先安装依赖：pip install flask pandas openpyxl xlrd werkzeug")
    print("=" * 50)
    # 启动Web服务（允许外部访问，端口5000）
    app.run(host='0.0.0.0', port=5000, debug=True)