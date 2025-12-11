import pymysql  # 确保这行是正确的“pymysql”（不是pymysq）
import pandas as pd
from openpyxl import Workbook
import os
def export_contacts_to_excel(conn):
    """导出联系人到Excel"""
    # 查询联系人及联系方式
    cursor = conn.cursor(pymysql.cursors.DictCursor)
    cursor.execute('SELECT * FROM contacts')
    contacts = cursor.fetchall()
    
    # 整理数据
    export_data = []
    for contact in contacts:
        contact_id = contact['id']
        cursor.execute(
            'SELECT type, value FROM contact_details WHERE contact_id = %s',
            (contact_id,)
        )
        details = cursor.fetchall()
        
        # 构建一行数据
        row = {
            '联系人ID': contact_id,
            '姓名': contact['name'],
            '是否收藏': '是' if contact['is_favorite'] == 1 else '否',
            '创建时间': contact['created_at']
        }
        
        # 添加联系方式（电话1、邮箱1、社交账号1等）
        detail_types = {}
        for detail in details:
            dtype = detail['type']
            if dtype not in detail_types:
                detail_types[dtype] = 1
            else:
                detail_types[dtype] += 1
            row[f'{dtype}{detail_types[dtype]}'] = detail['value']
        
        export_data.append(row)
    
    cursor.close()
    
    # 创建Excel文件
    df = pd.DataFrame(export_data)
    file_path = '../contacts_export.xlsx'
    df.to_excel(file_path, index=False, engine='openpyxl')
    
    return file_path

def import_contacts_from_excel(conn, file):
    """从Excel导入联系人"""
    # 读取Excel文件
    df = pd.read_excel(file, engine='openpyxl')
    cursor = conn.cursor()
    
    # 定义支持的联系方式类型
    support_types = ['电话', '邮箱', '社交账号', '地址']
    
    for _, row in df.iterrows():
        # 获取联系人基本信息
        name = row.get('姓名')
        is_favorite = 1 if row.get('是否收藏') == '是' else 0
        
        if not name:
            continue  # 姓名为空跳过
        
        try:
            # 插入联系人表
            cursor.execute(
                'INSERT INTO contacts (name, is_favorite) VALUES (%s, %s)',
                (name, is_favorite)
            )
            contact_id = cursor.lastrowid
            
            # 插入联系方式
            for dtype in support_types:
                # 处理多个同类型联系方式（电话1、电话2...）
                i = 1
                while True:
                    col_name = f'{dtype}{i}' if i > 1 else dtype
                    if col_name not in df.columns:
                        break
                    value = row.get(col_name)
                    if pd.notna(value) and str(value).strip():
                        cursor.execute(
                            'INSERT INTO contact_details (contact_id, type, value) VALUES (%s, %s, %s)',
                            (contact_id, dtype, str(value).strip())
                        )
                    i += 1
        
        except Exception as e:
            conn.rollback()
            raise e
    
    conn.commit()
    cursor.close()