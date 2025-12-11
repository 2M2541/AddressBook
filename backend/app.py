from flask import Flask, request, jsonify, send_file
from flask_cors import CORS
import pandas as pd
import os
import io

# 初始化Flask应用
app = Flask(__name__)
CORS(app)  # 允许跨域

# 模拟数据库（内存存储，重启后重置；如需持久化可替换为MySQL）
contacts = [
    {"id": 1, "name": "张三", "phone": "13800138000"},
    {"id": 2, "name": "李四", "phone": "13900139000"}
]
next_id = 3  # 自增ID

# -------------------------- 基础接口：增删查 --------------------------
# 1. 获取所有联系人
@app.route('/api/contacts', methods=['GET'])
def get_contacts():
    return jsonify(contacts), 200

# 2. 新增联系人
@app.route('/api/contacts', methods=['POST'])
def add_contact():
    global next_id
    data = request.get_json()
    if not data or not data.get('name') or not data.get('phone'):
        return jsonify({"error": "姓名和电话不能为空"}), 400
    
    new_contact = {"id": next_id, "name": data['name'], "phone": data['phone']}
    contacts.append(new_contact)
    next_id += 1
    return jsonify({"message": "添加成功", "contact": new_contact}), 201

# 3. 删除联系人
@app.route('/api/contacts/<int:contact_id>', methods=['DELETE'])
def delete_contact(contact_id):
    global contacts
    for idx, contact in enumerate(contacts):
        if contact['id'] == contact_id:
            del contacts[idx]
            return jsonify({"message": "删除成功"}), 200
    return jsonify({"error": "联系人不存在"}), 404

# -------------------------- Excel导出接口（修复版） --------------------------
@app.route('/api/contacts/export', methods=['GET'])
def export_contacts():
    # 将联系人列表转为DataFrame
    df = pd.DataFrame(contacts)
    df = df[['id', 'name', 'phone']]  # 调整列顺序
    
    # 生成Excel文件到内存（确保文件流完整）
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='通讯录', index=False)
    output.seek(0)  # 重置文件指针到开头
    
    # 修复：添加完整响应头，确保Excel识别文件
    response = send_file(
        output,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        as_attachment=True,
        download_name='通讯录导出.xlsx'
    )
    # 强制设置文件大小，避免Excel判定文件损坏
    response.headers['Content-Length'] = str(len(output.getvalue()))
    # 防止缓存导致的文件问题
    response.headers['Cache-Control'] = 'no-cache'
    return response

# -------------------------- Excel导入接口 --------------------------
@app.route('/api/contacts/import', methods=['POST'])
def import_contacts():
    global next_id
    # 检查是否有文件上传
    if 'file' not in request.files:
        return jsonify({"error": "请选择要导入的Excel文件"}), 400
    
    file = request.files['file']
    # 检查文件类型
    if file.filename == '' or not file.filename.endswith('.xlsx'):
        return jsonify({"error": "仅支持.xlsx格式的Excel文件"}), 400
    
    try:
        # 读取Excel文件
        df = pd.read_excel(file)
        # 检查必要列（name和phone）
        if 'name' not in df.columns or 'phone' not in df.columns:
            return jsonify({"error": "Excel文件必须包含'name'（姓名）和'phone'（电话）列"}), 400
        
        # 遍历Excel数据，新增到通讯录
        imported_count = 0
        for _, row in df.iterrows():
            name = str(row['name']).strip()
            phone = str(row['phone']).strip()
            if name and phone:  # 非空验证
                new_contact = {"id": next_id, "name": name, "phone": phone}
                contacts.append(new_contact)
                next_id += 1
                imported_count += 1
        
        return jsonify({
            "message": f"导入成功！共导入 {imported_count} 条联系人",
            "imported_count": imported_count
        }), 200
    
    except Exception as e:
        return jsonify({"error": f"导入失败：{str(e)}"}), 500

# 启动服务
if __name__ == '__main__':
    app.run(debug=True, port=5000)