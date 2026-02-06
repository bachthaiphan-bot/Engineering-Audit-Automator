import pandas as pd
import openpyxl
import os
from pathlib import Path
from typing import Dict, List, Any
import logging
import re

# 尝试导入 xlrd 以支持 .xls 文件（如果未安装则仅支持 .xlsx/.xlsm）
try:
    import xlrd  # type: ignore
except Exception:  # 兼容环境无 xlrd 的情况
    xlrd = None  # type: ignore

def is_date_field_by_name(field_name: str) -> bool:
    """根据字段名判断是否为日期字段"""
    if not isinstance(field_name, str):
        return False
    
    date_keywords = [
        '日期', 'date', '时间', 'time', '期间', 'period',
        '开工', '竣工', '完工', '签订', '审核', '编制', '报告',
        '批准', '验收', '交付', '开始', '结束', '截止'
    ]
    
    field_lower = field_name.lower()
    return any(keyword in field_lower for keyword in date_keywords)

def is_numeric_field_by_name(field_name: str) -> bool:
    """根据字段名判断是否为数值字段"""
    if not isinstance(field_name, str):
        return False
    
    numeric_keywords = [
        '金额', 'amount', '数量', 'quantity', '值', 'value', 
        '余额', 'balance', '价格', 'price', '费用', 'cost',
        '造价', '投资', '预算', '决算', '结算', '合同',
        '工程量', '面积', '长度', '重量', '体积', '比例',
        '百分比', '率', '元', '万元', '千元'
    ]
    
    field_lower = field_name.lower()
    return any(keyword in field_lower for keyword in numeric_keywords)

def is_date_format(number_format: str) -> bool:
    """检查Excel数字格式是否为日期格式"""
    if not number_format:
        return False
    
    # Excel内置日期格式代码
    date_format_codes = [
        '14', '15', '16', '17', '18', '19', '20', '21', '22'  # Excel内置日期格式
    ]
    
    if number_format in date_format_codes:
        return True
    
    # 检查格式字符串中是否包含日期指示符
    date_indicators = ['d', 'm', 'y', 'h', 's', '年', '月', '日', '时', '分', '秒']
    format_lower = number_format.lower()
    
    return any(indicator in format_lower for indicator in date_indicators)

def safe_read_cell_value(cell, header: str = "") -> Any:
    """安全读取单元格值，使用格式检测避免误转换"""
    if cell is None or cell.value is None:
        return None
    
    # 1. 首先检查openpyxl的内置日期检测
    if hasattr(cell, 'is_date') and cell.is_date:
        return cell.value  # 保持datetime对象
    
    # 2. 检查数字格式
    if hasattr(cell, 'number_format'):
        if is_date_format(cell.number_format):
            # 只有格式明确为日期时才允许日期转换
            if is_date_field_by_name(header):
                return cell.value
            else:
                # 即使格式是日期，但字段名不像日期字段，返回字符串
                return str(cell.value)
    
    # 3. 字段名启发式检查
    if isinstance(cell.value, (int, float)):
        # 如果字段名明确是数值字段，强制返回原始数值
        if is_numeric_field_by_name(header):
            return cell.value
        
        # 如果字段名明确是日期字段，且数值在合理范围内，才考虑日期转换
        if is_date_field_by_name(header):
            if isinstance(cell.value, (int, float)) and 36526 <= cell.value <= 55196:
                # 只有在字段名明确是日期字段时才转换
                try:
                    from datetime import datetime, timedelta
                    excel_epoch = datetime(1900, 1, 1)
                    return excel_epoch + timedelta(days=int(cell.value) - 2)
                except:
                    return cell.value
        
        # 默认返回原始数值
        return cell.value
    
    # 4. 对于其他类型，返回字符串形式以保持原始性
    return str(cell.value) if cell.value is not None else None

def read_excel_data_from_data_folder(data_folder: str = "data") -> Dict[str, Any]:
    """
    从data文件夹中读取所有Excel文件的数据
    返回格式: {
        'headers': ['项目编码', '编制日期', ...],  # P2:P155的标题
        'projects': [
            {
                'filename': '项目文件名',
                'data': {'项目编码': '值1', '编制日期': '值2', ...}
            },
            ...
        ]
    }
    """
    # 支持 .xlsx / .xlsm / .xls，过滤临时文件
    data_files = [
        f for f in os.listdir(data_folder)
        if (f.lower().endswith(('.xlsx', '.xlsm', '.xls'))) and not f.startswith('~$')
    ]
    
    if not data_files:
        raise ValueError(f"在 {data_folder} 文件夹中未找到Excel文件")
    
    headers = []
    projects = []
    
    # 读取第一个文件获取标题
    first_file = os.path.join(data_folder, data_files[0])
    ext = os.path.splitext(first_file)[1].lower()
    if ext in ('.xlsx', '.xlsm'):
        wb = openpyxl.load_workbook(first_file, data_only=True)
        # 查找隐藏工作表（包括隐藏的）
        target_sheet = None
        for sheet_name in wb.sheetnames:
            if '_calc' in sheet_name:
                target_sheet = sheet_name
                break
        
        if target_sheet is None:
            raise ValueError(f"在文件 {first_file} 中未找到包含'_calc'的工作表")
        
        ws = wb[target_sheet]
        # 读取P2:P155作为标题
        for i in range(2, 156):
            cell_value = ws[f'P{i}'].value
            if cell_value is not None:
                headers.append(str(cell_value))
            else:
                headers.append("")  # 空标题
    else:  # .xls
        if xlrd is None:
            raise ValueError("检测到 .xls 文件，但未安装 xlrd，无法读取 .xls。请安装 xlrd 或将文件转换为 .xlsx。")
        wb_xls = xlrd.open_workbook(first_file)
        # 查找隐藏工作表（包括隐藏的）
        target_sheet = None
        for sheet in wb_xls.sheets():
            if '_calc' in sheet.name:
                target_sheet = sheet.name
                break
        
        if target_sheet is None:
            raise ValueError(f"在文件 {first_file} 中未找到包含'_calc'的工作表")
        
        sh = wb_xls.sheet_by_name(target_sheet)
        # xlrd 行列从0开始，P列为索引15，范围P2:P155 => 行索引1..154
        for row_idx in range(1, 155):
            try:
                cell_value = sh.cell_value(row_idx, 15)
            except Exception:
                cell_value = None
            if cell_value not in (None, ""):
                headers.append(str(cell_value))
            else:
                headers.append("")
    
    print(f"读取到 {len(headers)} 个标题字段")
    
    # 读取所有文件的数据
    for filename in data_files:
        file_path = os.path.join(data_folder, filename)
        try:
            ext = os.path.splitext(filename)[1].lower()
            project_data = {}
            if ext in ('.xlsx', '.xlsm'):
                # 使用 openpyxl
                wb = openpyxl.load_workbook(file_path, data_only=True)
                # 查找隐藏工作表（包括隐藏的）
                target_sheet = None
                for sheet_name in wb.sheetnames:
                    if '_calc' in sheet_name:
                        target_sheet = sheet_name
                        break
                
                if target_sheet is None:
                    print(f"警告: 文件 {filename} 中未找到包含'_calc'的工作表，跳过")
                    continue
                
                ws = wb[target_sheet]
                # 读取Q2:Q155的数据 - 使用格式检测避免误转换
                for i, header in enumerate(headers):
                    if header:  # 只处理非空标题
                        cell = ws[f'Q{i+2}']
                        cell_value = safe_read_cell_value(cell, header)
                        project_data[header] = cell_value
            else:  # .xls
                if xlrd is None:
                    print(f"警告: 检测到 .xls 文件 {filename} 但未安装 xlrd，跳过")
                    continue
                wb_xls = xlrd.open_workbook(file_path)
                # 查找隐藏工作表（包括隐藏的）
                target_sheet = None
                for sheet in wb_xls.sheets():
                    if '_calc' in sheet.name:
                        target_sheet = sheet.name
                        break
                
                if target_sheet is None:
                    print(f"警告: 文件 {filename} 中未找到包含'_calc'的工作表，跳过")
                    continue
                
                sh = wb_xls.sheet_by_name(target_sheet)
                # 读取Q2:Q155，Q列索引为16，行索引1..(1+len(headers)-1)
                for i, header in enumerate(headers):
                    if header:
                        row_idx = 1 + i
                        try:
                            cell_value = sh.cell_value(row_idx, 16)
                        except Exception:
                            cell_value = None
                        project_data[header] = cell_value
            
            projects.append({
                'filename': filename,
                'data': project_data
            })
            
            print(f"成功读取文件: {filename}")
            
        except Exception as e:
            print(f"读取文件 {filename} 时出错: {e}")
            continue
    
    return {
        'headers': headers,
        'projects': projects
    }

def create_date_xlsx_from_data(data_folder: str = "data", output_file: str = "date.xlsx"):
    """
    根据data文件夹中的Excel文件自动生成date.xlsx
    第一行: 替换符标题 (P2:P155)
    后续行: 各个工作簿的数据 (Q2:Q155)
    """
    print("开始自动映射数据...")
    
    # 读取数据
    excel_data = read_excel_data_from_data_folder(data_folder)
    
    # 创建DataFrame：增加前导两列用于追踪源文件
    # 第一行: [源文件名, 源文件路径, P2:P155 标题...]
    df_headers = ["源文件名", "源文件路径"] + excel_data['headers']
    df_data = [df_headers]

    # 后续行: 每个项目的数据，前两列填充文件名与完整路径
    for project in excel_data['projects']:
        row_data = [
            project.get('filename', ''),
            os.path.join(data_folder, project.get('filename', '')),
        ]
        for header in excel_data['headers']:
            if header:
                row_data.append(project['data'].get(header, ""))
            else:
                row_data.append("")
        df_data.append(row_data)
    
    # 创建DataFrame
    df = pd.DataFrame(df_data)
    
    # 保存为Excel文件，保留原始格式
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='Sheet1', index=False, header=False)
        
        # 获取工作表对象以进行格式调整
        worksheet = writer.sheets['Sheet1']
        
        # 设置列宽
        for col in range(1, len(excel_data['headers']) + 1):
            worksheet.column_dimensions[openpyxl.utils.get_column_letter(col)].width = 20
    
    print(f"自动映射完成，生成文件: {output_file}")
    print(f"包含 {len(excel_data['projects'])} 个项目的数据")
    
    return output_file

if __name__ == "__main__":
    # 测试自动映射功能
    try:
        create_date_xlsx_from_data()
        print("自动映射测试成功!")
    except Exception as e:
        print(f"自动映射测试失败: {e}")
        import traceback
        traceback.print_exc()