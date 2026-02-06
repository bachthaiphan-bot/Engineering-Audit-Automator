from __future__ import annotations
import os
from typing import Any, Dict
from datetime import datetime, timedelta
import re

import argparse
import json
import logging
from logging.handlers import RotatingFileHandler
from pathlib import Path
from copy import deepcopy
import shutil
from concurrent.futures import ThreadPoolExecutor, as_completed

import pandas as pd
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Mm
from docx import Document
from openpyxl import load_workbook

# 可选引入：xlwings（如环境存在 Excel）
try:
    import xlwings as xw  # type: ignore
except Exception:
    xw = None  # 运行时检查

# GUI进度窗口相关导入
import sys
if getattr(sys, 'frozen', False):
    try:
        import tkinter as tk
        from tkinter import ttk
    except ImportError:
        tk = None

# 可选引入：tqdm 进度条
try:
    from tqdm import tqdm  # type: ignore
    # 在打包的exe中确保tqdm能正常工作
    if getattr(sys, 'frozen', False):
        # 修复打包后tqdm无法找到stdout的问题
        tqdm = lambda iterable, **kwargs: iterable
except Exception:  # 降级：无 tqdm 时按原生迭代
    def tqdm(iterable, **kwargs):  # type: ignore
        return iterable


def setup_logging(output_dir: str, debug: bool) -> logging.Logger:
    """初始化日志到控制台与文件（outputs/logs/app.log）。"""
    logs_dir = Path(output_dir) / "logs"
    logs_dir.mkdir(parents=True, exist_ok=True)
    log_path = logs_dir / "app.log"

    logger = logging.getLogger("docgen")
    logger.setLevel(logging.DEBUG if debug else logging.INFO)
    logger.handlers = []

    fmt = logging.Formatter("%(asctime)s | %(levelname)s | %(message)s")
    # 控制台
    ch = logging.StreamHandler()
    ch.setLevel(logging.DEBUG if debug else logging.INFO)
    ch.setFormatter(fmt)
    logger.addHandler(ch)
    # 文件
    # 使用大小轮转，避免日志无限增长
    rfh = RotatingFileHandler(log_path, maxBytes=5 * 1024 * 1024, backupCount=3, encoding="utf-8")
    rfh.setLevel(logging.DEBUG)
    rfh.setFormatter(fmt)
    logger.addHandler(rfh)

    return logger


class ProgressWindow:
    """GUI进度窗口类，用于在打包exe后显示进度"""
    
    def __init__(self, total_items):
        self.total_items = total_items
        self.current_item = 0
        
        # 只在打包为exe时且tk可用时创建窗口
        if getattr(sys, 'frozen', False) and 'tk' in globals() and tk:
            self.root = tk.Tk()
            self.root.title("工程竣工决算文档生成工具")
            self.root.geometry("400x100")
            self.root.resizable(False, False)
            
            # 创建标签和进度条
            self.label = tk.Label(self.root, text="正在处理...")
            self.label.pack(pady=10)
            
            self.progress = ttk.Progressbar(
                self.root, 
                length=300, 
                mode='determinate',
                maximum=total_items
            )
            self.progress.pack(pady=5)
            
            self.status_label = tk.Label(self.root, text="0 / {}".format(total_items))
            self.status_label.pack(pady=5)
            
            # 将窗口置于最前
            self.root.lift()
            self.root.attributes('-topmost', True)
            self.root.after_idle(self.root.attributes, '-topmost', False)
        else:
            self.root = None
    
    def update(self, current_item, message=""):
        """更新进度"""
        self.current_item = current_item
        
        if self.root:
            self.progress['value'] = current_item
            self.status_label.config(text="{} / {}".format(current_item, self.total_items))
            if message:
                self.label.config(text=message)
            self.root.update()
    
    def close(self):
        """关闭进度窗口"""
        if self.root:
            self.root.destroy()


def build_context(row_dict: Dict[str, Any], image_tpl: DocxTemplate, debug: bool = False, strict_date: bool = False) -> Dict[str, Any]:
    """
    将 Excel 行数据转为 docxtpl 渲染上下文，支持：
    - 普通字段：日期 → yyyy年mm月dd日，数值 → 千位分隔符
    - 图片字段：列名以 "img:" 开头
    - 列表字段：列名以 "list:" 开头
    """
    context: Dict[str, Any] = {}

    for key, value in row_dict.items():
        if pd.isna(value):
            value = ""

        # 图片字段：img:字段名
        if isinstance(key, str) and key.startswith("img:"):
            field = key.split(":", 1)[1]
            width_mm_value = row_dict.get(f"imgw:{field}")
            try:
                width_mm = Mm(float(width_mm_value)) if width_mm_value not in (None, "") and not pd.isna(width_mm_value) else None
            except Exception:
                width_mm = None

            image_path = str(value).strip()
            if image_path and os.path.exists(image_path):
                context[field] = InlineImage(image_tpl, image_path, width=width_mm) if width_mm else InlineImage(image_tpl, image_path)
            else:
                context[field] = None
            continue

        # 列表字段：list:字段名
        if isinstance(key, str) and key.startswith("list:"):
            field = key.split(":", 1)[1]
            items: Any = []
            text = str(value).strip()
            if text:
                import json
                try:
                    parsed = json.loads(text)
                    if isinstance(parsed, list):
                        items = parsed
                    else:
                        items = [parsed]
                except Exception:
                    items = []
                    for line in text.splitlines():
                        line = line.strip()
                        if not line:
                            continue
                        try:
                            items.append(json.loads(line))
                        except Exception:
                            items.append({"value": line})
            context[field] = items
            continue

        # 日期字段：严格/宽松判定并格式化
        if is_date_value(value, field_name=key, strict_date=strict_date):
            # 特殊处理“报告日期”字段：转换为大写中文数字格式
            if key == "报告日期":
                context[key] = format_chinese_date(value)
            else:
                context[key] = format_date_value(value)
            continue

        # 数值字段：加千位分隔符
        if isinstance(value, (int, float, str)):
            try:
                formatted_num = format_number_value(value)
                if formatted_num != str(value):
                    context[key] = formatted_num
                    continue
            except Exception:
                pass

        # 默认直接赋值
        context[key] = value

        if debug:
            try:
                print(f"[DEBUG] {key} -> 原始: {repr(row_dict.get(key))} | 替换后: {repr(context.get(key))}")
            except Exception:
                pass

    # 标题用项目名称：根据规则插入换行，以便模板用 nl2br 滤镜渲染两行
    try:
        base_title_name = str(context.get("项目名称") or context.get("name") or "").strip()
        # 保持模板原样式：占位直接渲染为纯文本；断行在渲染后处理
        context["标题项目名称_rt"] = base_title_name
        context["标题项目名称_plain"] = base_title_name
    except Exception:
        text = str(context.get("项目名称") or "")
        context["标题项目名称_rt"] = text
        context["标题项目名称_plain"] = text

    return context


def format_chinese_date(value: Any) -> str:
    """将日期值格式化为大写中文数字格式，例如：2025年07月25日 → 二〇二五年七月二十五日"""
    if pd.isna(value) or value == "":
        return ""
    try:
        # 仅当为“整数的 Excel 序列号”时才按日期处理
        if isinstance(value, (int, float)):
            if isinstance(value, float) and not float(value).is_integer():
                pass
            else:
                int_value = int(value)
                if 20000 <= int_value <= 50000:
                    excel_epoch = datetime(1900, 1, 1)
                    dt = excel_epoch + timedelta(days=int_value - 2)
                    return format_chinese_date_from_datetime(dt)
        if isinstance(value, str):
            date_formats = [
                "%Y-%m-%d", "%Y/%m/%d", "%Y年%m月%d日", "%Y.%m.%d",
                "%m/%d/%Y", "%d/%m/%Y"
            ]
            for fmt in date_formats:
                try:
                    dt = datetime.strptime(value.strip(), fmt)
                    return format_chinese_date_from_datetime(dt)
                except ValueError:
                    continue
            return str(value)
        if hasattr(value, "strftime"):
            return format_chinese_date_from_datetime(value)
        return str(value)
    except Exception:
        return str(value)


def format_chinese_date_from_datetime(dt: datetime) -> str:
    """从datetime对象生成大写中文数字格式的日期字符串"""
    year = dt.year
    month = dt.month
    day = dt.day

    # 中文数字映射
    chinese_digits = {
        0: '〇', 1: '一', 2: '二', 3: '三', 4: '四',
        5: '五', 6: '六', 7: '七', 8: '八', 9: '九'
    }

    # 年份处理
    year_str = ""
    for digit in str(year):
        year_str += chinese_digits[int(digit)]
    year_str += "年"

    # 月份处理
    month_str = ""
    if month < 10:
        month_str = chinese_digits[month] + "月"
    else:
        tens_digit = month // 10
        units_digit = month % 10
        if tens_digit == 1:
            month_str = "十"
        else:
            month_str = chinese_digits[tens_digit] + "十"
        if units_digit > 0:
            month_str += chinese_digits[units_digit]
        month_str += "月"

    # 日处理
    day_str = ""
    if day < 10:
        day_str = chinese_digits[day] + "日"
    else:
        tens_digit = day // 10
        units_digit = day % 10
        if tens_digit == 1:
            day_str = "十"
        else:
            day_str = chinese_digits[tens_digit] + "十"
        if units_digit > 0:
            day_str += chinese_digits[units_digit]
        day_str += "日"

    return year_str + month_str + day_str


def is_likely_date_field(field_name: str) -> bool:
    """根据字段名判断是否可能是日期字段"""
    if not isinstance(field_name, str):
        return False
    
    date_keywords = [
        '日期', '时间', '年月日', '开工', '竣工', '完工', '签订', '审核', 
        '编制', '报告', '批准', '验收', '交付', '开始', '结束', '截止',
        'date', 'time', 'start', 'end', 'finish', 'complete'
    ]
    
    field_lower = field_name.lower()
    return any(keyword in field_lower for keyword in date_keywords)


def is_likely_numeric_field(field_name: str) -> bool:
    """根据字段名判断是否可能是数值字段"""
    if not isinstance(field_name, str):
        return False
    
    numeric_keywords = [
        '金额', '费用', '成本', '价格', '单价', '总价', '造价', '投资',
        '预算', '决算', '结算', '合同', '工程量', '数量', '面积', '长度',
        '重量', '体积', '比例', '百分比', '率', '元', '万元', '千元',
        'amount', 'cost', 'price', 'budget', 'total', 'sum', 'value'
    ]
    
    field_lower = field_name.lower()
    return any(keyword in field_lower for keyword in numeric_keywords)


def is_date_value(value: Any, field_name: str = "", strict_date: bool = False) -> bool:
    """判断值是否为日期类型。
    
    REFACTORED: 使用格式检测和字段名启发式，避免数值被误识别为日期

    strict_date=False（默认，宽松）：
        - 允许 datetime 对象
        - 仅在字段名明确为日期字段时，才允许Excel序列号转换
        - 数值字段严格禁止日期转换

    strict_date=True（严格）：
        - 仅接受 datetime 或能被固定格式解析的日期字符串
        - 完全禁止数值型日期转换
    """
    if pd.isna(value) or value == "":
        return False
    try:
        # 统一先判断 datetime
        if hasattr(value, "strftime"):
            return True

        if strict_date:
            # 严格模式：不接受数值型日期
            if isinstance(value, (int, float)):
                return False
            if isinstance(value, str):
                value_str = value.strip()
                if not value_str:
                    return False
                for fmt in [
                    "%Y-%m-%d", "%Y/%m/%d", "%Y.%m.%d", "%Y年%m月%d日",
                    "%m/%d/%Y", "%d/%m/%Y", "%Y-%m-%d %H:%M:%S", "%Y/%m/%d %H:%M:%S"
                ]:
                    try:
                        datetime.strptime(value_str, fmt)
                        return True
                    except ValueError:
                        continue
                return False
        else:
            # 宽松模式：兼容 Excel 序列号与多格式字符串
            # 仅当为“整数的 Excel 序列号”时才识别为日期，避免金额(如 28041.07)被误判
            if isinstance(value, (int, float)):
                # 只有在字段名明确是日期字段时，才考虑Excel序列号转换
                is_likely_date = is_likely_date_field(field_name)
                is_likely_numeric = is_likely_numeric_field(field_name)
                
                # 如果字段名明显是数值类型，严格禁止日期转换
                if is_likely_numeric:
                    return False
                
                # 只有字段名明确是日期字段时，才允许序列号转换
                if not is_likely_date:
                    return False
                
                # 即使字段名是日期字段，也要更严格的范围检查
                if isinstance(value, float) and not float(value).is_integer():
                    return False
                    
                int_value = int(value)
                # 更保守的Excel日期序列号范围：2010-01-01到2030-12-31
                # 40179 = 2010-01-01, 47482 = 2030-12-31
                if 40179 <= int_value <= 47482:
                    return True
                return False
            if isinstance(value, str):
                value_str = value.strip()
                if not value_str:
                    return False
                
                # 根据字段名进行预判断
                is_likely_date = is_likely_date_field(field_name)
                is_likely_numeric = is_likely_numeric_field(field_name)
                
                # 如果字段名明显是数值类型，严格禁止日期转换
                if is_likely_numeric:
                    return False
                    
                # 检查明显的数值特征 - 如果有这些特征，绝对不是日期
                if re.search(r'[,￥$¥%]', value_str) or value_str.count('.') > 1:
                    return False
                
                # 检查固定的日期格式
                date_formats = [
                    "%Y-%m-%d", "%Y/%m/%d", "%Y年%m月%d日", "%Y.%m.%d",
                    "%m/%d/%Y", "%d/%m/%Y", "%Y-%m-%d %H:%M:%S", "%Y/%m/%d %H:%M:%S"
                ]
                for fmt in date_formats:
                    try:
                        datetime.strptime(value_str, fmt)
                        return True
                    except ValueError:
                        continue
                date_patterns = [
                    r"\d{4}[-/年]\d{1,2}[-/月]\d{1,2}",
                    r"\d{1,2}[-/]\d{1,2}[-/]\d{4}",
                ]
                for pattern in date_patterns:
                    if re.search(pattern, value_str):
                        return True
        return False
    except Exception:
        return False


def format_date_value(value: Any) -> str:
    """将日期值格式化为 yyyy年mm月dd日"""
    if pd.isna(value) or value == "":
        return ""
    try:
        # 仅当为“整数的 Excel 序列号”时才按日期处理
        if isinstance(value, (int, float)):
            if isinstance(value, float) and not float(value).is_integer():
                pass
            else:
                int_value = int(value)
                if 20000 <= int_value <= 50000:
                    excel_epoch = datetime(1900, 1, 1)
                    dt = excel_epoch + timedelta(days=int_value - 2)
                    return dt.strftime("%Y年%m月%d日")
        if isinstance(value, str):
            date_formats = [
                "%Y-%m-%d", "%Y/%m/%d", "%Y年%m月%d日", "%Y.%m.%d",
                "%m/%d/%Y", "%d/%m/%Y"
            ]
            for fmt in date_formats:
                try:
                    
                    dt = datetime.strptime(value.strip(), fmt)
                    return dt.strftime("%Y年%m月%d日")
                except ValueError:
                    continue
            return str(value)
        if hasattr(value, "strftime"):
            return value.strftime("%Y年%m月%d日")
        return str(value)
    except Exception:
        return str(value)


def format_number_value(value: Any) -> str:
    """将数值格式化为千位分隔符形式，统一显示两位小数"""
    if pd.isna(value) or value == "":
        return ""
    try:
        if isinstance(value, str):
            cleaned = re.sub(r"[,\s￥$]", "", value.strip())
            num_value = float(cleaned)
        else:
            num_value = float(value)
        # 统一显示两位小数，包括整数
        return f"{num_value:,.2f}"
    except (ValueError, TypeError):
        return str(value)


def insert_title_linebreak(name: str) -> str:
    """标题断行：保留前12个字符，之后换行。"""
    s = name.strip()
    if not s:
        return s
    return s if len(s) <= 12 else (s[:12] + "\n" + s[12:])


def register_jinja_filters(tpl: DocxTemplate) -> None:
    """为 docxtpl 模板注册自定义过滤器（如 nl2br）。

    有些环境下 docxtpl 版本不自带 nl2br，这里注入一个兼容实现，
    将字符串中的换行符转换为 RichText 的换行效果。
    """
    try:
        env = None
        # 兼容不同 docxtpl 版本的环境获取方式
        if hasattr(tpl, 'get_jinja_env'):
            try:
                env = tpl.get_jinja_env()
            except Exception:
                env = None
        if env is None and hasattr(tpl, 'jinja_env'):
            env = getattr(tpl, 'jinja_env')
        if env is None and hasattr(tpl, 'environment'):
            env = getattr(tpl, 'environment')
        if env is None:
            return
        if 'nl2br' not in env.filters:
            from docxtpl import RichText  # 局部导入以避免打包时不必要依赖
            def _nl2br(value):
                if value is None:
                    return ""
                text = str(value)
                rt = RichText()
                parts = text.split("\n")
                for idx, part in enumerate(parts):
                    if idx:
                        rt.add("\n")
                    rt.add(part)
                return rt
            env.filters['nl2br'] = _nl2br
    except Exception:
        pass


def split_title_keep_style(doc_path: str, title_text_with_break: str, break_after_chars: int, logger: logging.Logger) -> None:
    """在渲染后对仅含标题文本的段落进行"就地断行"。
    采用最简单可靠的方法：在原run的text中直接插入\\n，不创建新run，确保格式100%一致。
    同时处理表格中的标题。
    """
    try:
        doc = Document(doc_path)
        found = False
        target_title = title_text_with_break.strip()
        
        # 先检查普通段落
        for p in doc.paragraphs:
            para_text = p.text.strip()
            if target_title in para_text and abs(len(para_text) - len(target_title)) <= 5:
                found = True
                logger.debug(f"找到标题段落，runs数={len(p.runs)}, 文本='{para_text[:30]}'")
                
                if len(p.runs) == 0:
                    logger.warning("标题段落无runs，跳过")
                    break
                
                # 单run情况：直接在text中插入\\n
                if len(p.runs) == 1:
                    r = p.runs[0]
                    text = r.text.strip()
                    if len(text) <= break_after_chars:
                        logger.debug(f"标题长度<={break_after_chars}字符，无需断行")
                        break
                    before = text[:break_after_chars]
                    after = text[break_after_chars:]
                    r.text = before + "\n" + after
                    logger.info(f"标题已在第{break_after_chars}字符处断行（单run内换行）")
                    break
                else:
                    # 多run：找到切分点所在run，在其中插入\\n
                    consumed = 0
                    for r in p.runs:
                        r_len = len(r.text)
                        if consumed + r_len >= break_after_chars:
                            local_idx = break_after_chars - consumed
                            if local_idx <= 0 or local_idx >= r_len:
                                break
                            before = r.text[:local_idx]
                            after = r.text[local_idx:]
                            r.text = before + "\n" + after
                            logger.info(f"标题已在第{break_after_chars}字符处断行（多run内换行）")
                            break
                        consumed += r_len
                    break
        
        # 如果在段落中没找到，检查表格
        if not found:
            for table in doc.tables:
                for row in table.rows:
                    for cell in row.cells:
                        for p in cell.paragraphs:
                            para_text = p.text.strip()
                            if target_title in para_text and abs(len(para_text) - len(target_title)) <= 5:
                                found = True
                                logger.debug(f"在表格中找到标题段落，runs数={len(p.runs)}, 文本='{para_text[:30]}'")
                                
                                if len(p.runs) == 0:
                                    logger.warning("标题段落无runs，跳过")
                                    break
                                
                                # 单run情况：直接在text中插入\\n
                                if len(p.runs) == 1:
                                    r = p.runs[0]
                                    text = r.text.strip()
                                    if len(text) <= break_after_chars:
                                        logger.debug(f"标题长度<={break_after_chars}字符，无需断行")
                                        break
                                    before = text[:break_after_chars]
                                    after = text[break_after_chars:]
                                    r.text = before + "\n" + after
                                    logger.info(f"表格中标题已在第{break_after_chars}字符处断行（单run内换行）")
                                    break
                                else:
                                    # 多run：找到切分点所在run，在其中插入\\n
                                    consumed = 0
                                    for r in p.runs:
                                        r_len = len(r.text)
                                        if consumed + r_len >= break_after_chars:
                                            local_idx = break_after_chars - consumed
                                            if local_idx <= 0 or local_idx >= r_len:
                                                break
                                            before = r.text[:local_idx]
                                            after = r.text[local_idx:]
                                            r.text = before + "\n" + after
                                            logger.info(f"表格中标题已在第{break_after_chars}字符处断行（多run内换行）")
                                            break
                                        consumed += r_len
                                    break
                        if found:
                            break
                    if found:
                        break
                if found:
                    break
                    
        if not found:
            logger.warning(f"未找到标题段落进行断行，标题文本: {title_text_with_break[:30] if title_text_with_break else '空'}")
        doc.save(doc_path)
    except Exception as e:
        logger.warning(f"标题断行保持样式失败: {e}")


def delete_paragraphs_containing_phrase(doc_path: str, phrase: str, logger: logging.Logger) -> None:
    try:
        doc = Document(doc_path)
        removed = False
        for p in list(doc.paragraphs):
            if phrase in p.text:
                pe = p._element
                pe.getparent().remove(pe)
                removed = True
        if removed:
            doc.save(doc_path)
            logger.info(f"已删除包含短语“{phrase}”的段落")
    except Exception as e:
        logger.warning(f"删除包含短语的段落失败: {e}")


def _clone_run_format(src_run, dst_run) -> None:
    """完整复制 run 的所有格式属性，确保目标 run 与源 run 格式完全一致。
    包括：字体、字号、加粗、斜体、颜色、东亚字体、字符样式等。
    """
    try:
        # 方法1: 深拷贝整个 rPr 节点（最彻底）
        src_r = getattr(src_run, "_r", None)
        dst_r = getattr(dst_run, "_r", None)
        if src_r is not None and dst_r is not None:
            if getattr(src_r, "rPr", None) is not None:
                dst_r.rPr = deepcopy(src_r.rPr)
                return
        
        # 方法2: 如果方法1失败，逐项复制高层API属性
        if hasattr(src_run, 'style') and src_run.style is not None:
            dst_run.style = src_run.style
        
        if hasattr(src_run, 'font') and hasattr(dst_run, 'font'):
            src_font = src_run.font
            dst_font = dst_run.font
            
            # 复制字体名称（中文、西文、东亚）
            if src_font.name is not None:
                dst_font.name = src_font.name
            
            # 复制字号
            if src_font.size is not None:
                dst_font.size = src_font.size
            
            # 复制加粗、斜体、下划线
            if src_font.bold is not None:
                dst_font.bold = src_font.bold
            if src_font.italic is not None:
                dst_font.italic = src_font.italic
            if src_font.underline is not None:
                dst_font.underline = src_font.underline
            
            # 复制颜色
            if src_font.color and src_font.color.rgb is not None:
                dst_font.color.rgb = src_font.color.rgb
            
            # 复制其他属性
            try:
                if hasattr(src_font, 'all_caps') and src_font.all_caps is not None:
                    dst_font.all_caps = src_font.all_caps
                if hasattr(src_font, 'small_caps') and src_font.small_caps is not None:
                    dst_font.small_caps = src_font.small_caps
                if hasattr(src_font, 'strike') and src_font.strike is not None:
                    dst_font.strike = src_font.strike
                if hasattr(src_font, 'double_strike') and src_font.double_strike is not None:
                    dst_font.double_strike = src_font.double_strike
            except:
                pass
        
        # 复制其他run级别属性
        if hasattr(src_run, 'underline') and src_run.underline is not None:
            dst_run.underline = src_run.underline
        if hasattr(src_run, 'bold') and src_run.bold is not None:
            dst_run.bold = src_run.bold
        if hasattr(src_run, 'italic') and src_run.italic is not None:
            dst_run.italic = src_run.italic
            
    except Exception as e:
        # 失败则忽略，保持已有格式
        pass


def remove_phrase_in_paragraphs(doc_path: str, phrase: str, logger: logging.Logger) -> None:
    """只删除段落中的指定短语，保留该段落其他文本与样式（粗略：用首个 run 写回整段文本）。"""
    try:
        doc = Document(doc_path)
        changed = False
        for p in doc.paragraphs:
            if phrase in p.text:
                new_text = p.text.replace(phrase, "")
                if not p.runs:
                    continue
                p.runs[0].text = new_text
                for r in p.runs[1:]:
                    r.text = ""
                changed = True
        if changed:
            doc.save(doc_path)
            logger.info(f"已从段落中移除短语“{phrase}”")
    except Exception as e:
        logger.warning(f"移除短语失败: {e}")


def remove_phrase_in_paragraphs_preserve_runs(doc_path: str, phrase: str, logger: logging.Logger) -> None:
    """在保留段落内原有 run 样式的前提下，精准删除指定短语（支持跨 run）。"""
    try:
        if not phrase:
            return
        doc = Document(doc_path)
        changed_total = False
        for p in doc.paragraphs:
            full_text = p.text
            if not full_text or phrase not in full_text:
                continue
            # 定位要删除的区间（仅处理首次出现，避免误删）
            start = full_text.find(phrase)
            end = start + len(phrase)
            # 按 runs 粒度改写
            cur = 0
            changed_para = False
            for run in p.runs:
                t = run.text or ""
                L = len(t)
                seg_start = cur
                seg_end = cur + L
                # 无重叠
                if end <= seg_start or start >= seg_end:
                    cur += L
                    continue
                # 有重叠：保留 run 内不在删除区间的部分
                left_keep = ""
                right_keep = ""
                if start > seg_start:
                    left_keep = t[: start - seg_start]
                if end < seg_end:
                    right_keep = t[end - seg_start :]
                new_t = left_keep + right_keep
                if new_t != t:
                    run.text = new_t
                    changed_para = True
                cur += L
            if changed_para:
                changed_total = True
        if changed_total:
            doc.save(doc_path)
            logger.info(f"已从段落中精准移除短语“{phrase}”（保留样式）")
    except Exception as e:
        logger.warning(f"精准移除短语失败: {e}")


def remove_init_design_phrase(doc_path: str, context: Dict[str, Any], logger: logging.Logger) -> None:
    """仅当初设批复字段为空或0时，删除"于 初设批复时间 取得《初设批复名称》（初设批复文件），"这段文字。"""
    try:
        time_v = context.get("初设批复时间")
        name_v = context.get("初设批复名称")
        file_v = context.get("初设批复文件")
        
        # 判断是否有任一字段为空或为默认值
        is_empty = (_is_blank(time_v) or _is_blank(name_v) or _is_blank(file_v) or
                   _is_empty_or_zero(name_v) or _is_empty_or_zero(file_v))
        
        if is_empty:
            # 如果有字段为空，才执行删除
            # 构建实际渲染后的短语
            time_str = str(time_v) if time_v else "00:00:00"
            name_str = str(name_v) if name_v else "0.00"
            file_str = str(file_v) if file_v else "0.00"
            phrase_ctx = f"于{time_str}取得《{name_str}》（{file_str}），"
            
            # 默认异常占位值
            phrase_zero = "于00:00:00取得《0.00》（0.00），"
            
            # 移除两种情况
            remove_phrase_in_paragraphs_preserve_runs(doc_path, phrase_ctx, logger)
            remove_phrase_in_paragraphs_preserve_runs(doc_path, phrase_zero, logger)
        else:
            logger.debug("初设批复字段均有效，保留相关句子")
    except Exception as e:
        logger.warning(f"删除初设批复片段失败: {e}")


def remove_init_phrase_in_manual(doc_path: str, context: Dict[str, Any], logger: logging.Logger) -> None:
    """在 template_说明书.docx 的句式中，仅当初设字段为空或0时，删除"、《初设…》（初设…）"部分。
    例如：1.《可研…》（可研…）、《初设…》（初设…）； → 1.《可研…》（可研…）；
    """
    try:
        name_v = context.get("初设批复名称")
        file_v = context.get("初设批复文件")
        
        # 判断是否有字段为空或为默认值
        is_empty = (_is_blank(name_v) or _is_blank(file_v) or
                   _is_empty_or_zero(name_v) or _is_empty_or_zero(file_v))
        
        if is_empty:
            # 如果有字段为空，才执行删除
            name_str = str(name_v) if name_v else "0.00"
            file_str = str(file_v) if file_v else "0.00"
            phrase1 = f"、《{name_str}》（{file_str}）"
            phrase2 = "、《0.00》（0.00）"
            remove_phrase_in_paragraphs_preserve_runs(doc_path, phrase1, logger)
            remove_phrase_in_paragraphs_preserve_runs(doc_path, phrase2, logger)
        else:
            logger.debug("初设批复字段均有效，保留说明书中的相关内容")
    except Exception as e:
        logger.warning(f"说明书删除初设短语失败: {e}")


def update_catalog_when_initial_missing(doc_path: str, context: Dict[str, Any], logger: logging.Logger) -> None:
    """若初设批复名称/文件缺失或0，则在报告目录页将条目2删除，并把后续序号上移。
    这里通过简单文本替换实现：
      - 删除包含“《初设批复名称》（初设批复文件）”的段落
      - 同时将“3.”改为“2.”，“4.”改为“3.”（仅在同页目录区域内）
    注意：该处理假设目录是按普通段落/编号构成，而非自动目录域。
    """
    try:
        name_v = str(context.get("初设批复名称") or "").strip()
        file_v = str(context.get("初设批复文件") or "").strip()
        is_missing = (name_v == "" or name_v == "0.00" or file_v == "" or file_v == "0.00")
        if not is_missing:
            return
        doc = Document(doc_path)
        removed = False
        for p in list(doc.paragraphs):
            t = p.text.strip()
            if ("初设批复名称" in t and "初设批复文件" in t) or (f"《{name_v}》" in t and f"（{file_v}）" in t):
                pe = p._element
                pe.getparent().remove(pe)
                removed = True
        if removed:
            # 简单的序号上移：将目录页中的“3.”替换为“2.”，“4.”替换为“3.”
            for p in doc.paragraphs:
                if "3." in p.text or "4." in p.text:
                    new_t = p.text
                    new_t = new_t.replace("4.", "#TMP#")
                    new_t = new_t.replace("3.", "2.")
                    new_t = new_t.replace("#TMP#", "3.")
                    if new_t != p.text:
                        if p.runs:
                            p.runs[0].text = new_t
                            for r in p.runs[1:]:
                                r.text = ""
                        else:
                            p.text = new_t
            doc.save(doc_path)
            logger.info("报告目录：已删除初设批复条目并上移序号")
    except Exception as e:
        logger.warning(f"更新目录失败: {e}")


def remove_empty_unit_lines(doc_path: str, context: Dict[str, Any], logger: logging.Logger) -> None:
    """删除审核报告中空值或0值的单位行。
    
    处理以下单位字段，当字段为空或0时删除整行：
    - 项目建设管理单位：{{项目建设管理单位}}
    - 施工单位：{{施工单位}}  
    - 设计单位：{{设计单位}}
    - 监理单位：{{监理单位}}
    """
    try:
        # 定义需要检查的单位字段
        unit_fields = [
            ("项目建设管理单位", "项目建设管理单位"),
            ("施工单位", "施工单位"),
            ("设计单位", "设计单位"), 
            ("监理单位", "监理单位")
        ]
        
        doc = Document(doc_path)
        removed_any = False
        
        for unit_label, field_key in unit_fields:
            field_value = context.get(field_key)
            
            # 判断字段是否为空或0
            is_empty = (_is_blank(field_value) or _is_empty_or_zero(field_value))
            
            if is_empty:
                # 删除包含该单位标签的段落
                for p in list(doc.paragraphs):
                    if f"{unit_label}：" in p.text:
                        pe = p._element
                        pe.getparent().remove(pe)
                        removed_any = True
                        logger.info(f"已删除空值单位行: {unit_label}")
                        break
        
        if removed_any:
            doc.save(doc_path)
            
    except Exception as e:
        logger.warning(f"删除空值单位行失败: {e}")


def sanitize_filename(name: str) -> str:
    """移除 Windows 非法文件名字符并去除首尾空白。"""
    if not isinstance(name, str):
        name = str(name)
    name = name.strip()
    # Windows 不允许的字符: \ / : * ? " < > |
    return re.sub(r"[\\/:*?\"<>|]", "_", name) or "输出"

def render_excel_template(template_path: str, context: Dict[str, Any], output_path: str) -> None:
    """
    基于 openpyxl 的简易占位符替换：
    - 在 Excel 单元格文本中查找 {字段名} 并用 context 中的值替换
    - 值将进行日期与数值格式化；其他类型转为字符串
    - 不改动公式与非文本单元格
    """
    from openpyxl import load_workbook
    import re as _re

    def _stringify(value: Any) -> str:
        if is_date_value(value, field_name="", strict_date=False):
            return format_date_value(value)
        try:
            # 尝试数值格式化
            if isinstance(value, (int, float, str)):
                formatted_num = format_number_value(value)
                # 若格式化后与原不同，采用格式化
                if formatted_num != str(value):
                    return formatted_num
        except Exception:
            pass
        return "" if value is None or (isinstance(value, float) and pd.isna(value)) else str(value)

    # 读取模板
    wb = load_workbook(template_path)

    # 占位符模式：{字段}
    placeholder_pattern = _re.compile(r"\{([^{}\s]+)\}")

    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                val = cell.value
                if isinstance(val, str) and "{" in val and "}" in val:
                    def _replace(m: _re.Match[str]) -> str:
                        key = m.group(1)
                        raw = context.get(key, "")
                        return _stringify(raw)

                    try:
                        new_text = placeholder_pattern.sub(_replace, val)
                        cell.value = new_text
                    except Exception:
                        # 保底：失败则不替换
                        pass

    # 确保输出目录存在
    out_dir = os.path.dirname(output_path)
    if out_dir:
        os.makedirs(out_dir, exist_ok=True)
    wb.save(output_path)


def force_recalc_with_excel(path: str, allow_com: bool = True) -> bool:
    """使用Excel计算公式，返回True表示成功，False表示失败（会静默跳过）
    allow_com=False 时跳过 COM 自动化（用于并行模式）"""
    if xw is None or not allow_com:
        return False
    app = None
    try:
        # 增加重试机制，Excel进程可能被占用
        for attempt in range(2):
            try:
                app = xw.App(visible=False, add_book=False)
                break
            except Exception as e:
                if attempt == 0:
                    # 第一次失败，等待1秒后重试
                    import time
                    time.sleep(1)
                else:
                    # 第二次仍失败，放弃
                    raise e
        
        wb = app.books.open(path)
        wb.app.api.CalculateFullRebuild()
        wb.save()
        wb.close()
        return True
    except Exception as e:
        # Excel启动失败不应该中断整个流程，静默跳过即可
        # 日志会在调用处记录
        return False
    finally:
        if app is not None:
            try:
                app.quit()
            except:
                pass


def _is_empty_or_zero(value: Any) -> bool:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return True
    text = str(value).strip()
    if text == "":
        return True
    try:
        cleaned = re.sub(r"[,%\s]", "", text)
        cleaned = cleaned.replace("％", "").replace("%", "")
        if cleaned in ("", ".", "-"):
            return True
        num = float(cleaned)
        return abs(num) == 0.0
    except Exception:
        return False


def _is_blank(value: Any) -> bool:
    """仅判断是否为空白（None/NaN/空字符串）。"""
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return True
    return str(value).strip() == ""

def _delete_table_row_by_first_cell_text(doc_path: str, row_label: str, logger: logging.Logger) -> None:
    try:
        doc = Document(doc_path)
        removed = False
        for table in doc.tables:
            for row in table.rows:
                first_text = row.cells[0].text.strip() if row.cells else ""
                if first_text == row_label:
                    tbl = row._tr
                    tbl.getparent().remove(tbl)
                    removed = True
                    break
            if removed:
                break
        if removed:
            doc.save(doc_path)
            logger.info(f"已根据空值规则删除表格行: {row_label}")
    except Exception as e:
        logger.warning(f"删除表格行失败（{row_label}）: {e}")


def _delete_paragraphs_by_anchor_when_any_empty(
    doc_path: str,
    anchor_text: str,
    required_keys: list[str],
    context: Dict[str, Any],
    logger: logging.Logger,
) -> None:
    """当 required_keys 中任一值为空/为0时，删除包含 anchor_text 的整段落。"""
    try:
        if not any(k in context for k in required_keys):
            return
        need_delete = False
        for k in required_keys:
            v = context.get(k)
            # 金额/比率等数值用空或0判定；文本用空白判定
            if k.endswith("金额") or k.endswith("比率"):
                if _is_empty_or_zero(v):
                    need_delete = True
                    break
            else:
                if _is_blank(v):
                    need_delete = True
                    break
        if not need_delete:
            return

        doc = Document(doc_path)
        removed = False
        for p in list(doc.paragraphs):
            if anchor_text in p.text:
                pe = p._element
                pe.getparent().remove(pe)
                removed = True
        if removed:
            doc.save(doc_path)
            logger.info(f"已删除包含“{anchor_text}”的段落（因字段空/0）")
    except Exception as e:
        logger.warning(f"删除段落失败（{anchor_text}）: {e}")


def cleanup_document_rows_by_context(doc_path: str, context: Dict[str, Any], logger: logging.Logger) -> None:
    """根据上下文数值是否为空/为0，删除对应标签的整行。

    当前规则：
    - 当以下字段全部为空或为0时，删除对应的行：
      建筑工程 → [建筑概算金额, 建筑决算金额, 建筑抵扣增值税, 建筑超节支金额, 建筑超节支节率]
    """
    try:
        rules = [
    ("建筑工程", ["建筑概算金额", "建筑决算金额", "建筑抵扣增值税", "建筑超节支金额", "建筑超节支节率"]),
    ("安装工程", ["安装概算金额", "安装决算金额", "安装抵扣增值税", "安装超节支金额", "安装超节支节率"]),
    ("设备购置", ["设备概算金额", "设备决算金额", "设备抵扣增值税", "设备超节支金额", "设备超节支节率"]),
    ("其他费用", ["其他概算金额", "其他决算金额", "其他抵扣增值税", "其他超节支金额", "其他超节支节率"]),
    ("合计", ["合计概算金额", "合计决算金额", "合计抵扣增值税", "合计超节支金额", "合计超节支节率"]),
]
        for label, keys in rules:
            values = [context.get(k, None) for k in keys]
            if all(_is_empty_or_zero(v) for v in values):
                _delete_table_row_by_first_cell_text(doc_path, label, logger)
        # 供应商行：任一字段为空则删除（金额为空或为0也视为空）
        for i in range(1, 16):  # 扩展到15个供应商
            name_v = context.get(f"供应商{i}")
            amount_v = context.get(f"供应商应付{i}")
            kind_v = context.get(f"供应商款项性质{i}")
            if _is_blank(name_v) or _is_empty_or_zero(amount_v) or _is_blank(kind_v):
                _delete_table_row_by_first_cell_text(doc_path, str(i), logger)

        # 段落类语句：若关键字段任一为空/为0，则删除整段
        _delete_paragraphs_by_anchor_when_any_empty(
            doc_path,
            anchor_text="建筑工程支出",
            required_keys=["建筑结余情况", "建筑结余金额", "建筑结余比率", "建筑原因说明"],
            context=context,
            logger=logger,
        )
        _delete_paragraphs_by_anchor_when_any_empty(
            doc_path,
            anchor_text="安装工程支出",
            required_keys=["安装结余情况", "安装结余金额", "安装结余比率", "安装原因说明"],
            context=context,
            logger=logger,
        )
        _delete_paragraphs_by_anchor_when_any_empty(
            doc_path,
            anchor_text="设备购置支出",
            required_keys=["设备结余情况", "设备结余金额", "设备结余比率", "设备原因说明"],
            context=context,
            logger=logger,
        )
        _delete_paragraphs_by_anchor_when_any_empty(
            doc_path,
            anchor_text="其他费用支出",
            required_keys=["其他结余情况", "其他结余金额", "其他结余比率", "其他原因说明"],
            context=context,
            logger=logger,
        )

        # 审核报告/说明书：初设批复句子若任一字段为空则删除
        # 精准删除“于xx取得《xx》（xx），”片段，保留段落其余部分
        remove_init_design_phrase(doc_path, context, logger)
    except Exception as e:
        logger.warning(f"按上下文清理表格行失败: {e}")

def convert_excel_to_xlsx(source_path: str, dest_path: str, logger: logging.Logger, allow_com: bool = True) -> bool:
    """将Excel文件转换为.xlsx格式，保留所有公式计算结果。

    - .xlsx: 直接复制
    - .xlsm: 用Excel打开并计算公式后另存为.xlsx（移除宏但保留所有数据）
    - .xls: 用Excel打开并另存为.xlsx
    
    返回 True/False 表示是否成功。
    """
    try:
        os.makedirs(os.path.dirname(dest_path), exist_ok=True)
        source_ext = os.path.splitext(source_path)[1].lower()
        
        # 如果已经是xlsx，直接复制
        if source_ext == ".xlsx":
            shutil.copy2(source_path, dest_path)
            return True
        
        # 对于xlsm和xls，使用Excel COM进行转换（保留公式计算结果）
        if xw is not None and allow_com:
            app = None
            try:
                # 尝试用xlwings（Excel COM）转换
                for attempt in range(2):
                    try:
                        app = xw.App(visible=False, add_book=False)
                        break
                    except Exception as e:
                        if attempt == 0:
                            import time
                            time.sleep(1)
                        else:
                            raise e
                
                # 打开源文件
                wb = app.books.open(source_path)
                # 计算所有公式
                wb.app.api.CalculateFullRebuild()
                # 另存为xlsx格式（51 = xlOpenXMLWorkbook）
                wb.api.SaveAs(os.path.abspath(dest_path), FileFormat=51)
                wb.close()
                return True
            except Exception as e:
                logger.warning(f"Excel COM转换失败: {e}，尝试使用openpyxl...")
            finally:
                if app is not None:
                    try:
                        app.quit()
                    except:
                        pass
        
        # 备用方案1：使用openpyxl with data_only=True（读取计算值）
        try:
            wb = load_workbook(source_path, data_only=True, keep_vba=False)
            # 检查是否有数据（避免全是0的情况）
            has_data = False
            for sheet in wb.worksheets:
                for row in sheet.iter_rows(min_row=1, max_row=10, max_col=10):
                    for cell in row:
                        if cell.value not in (None, 0, ''):
                            has_data = True
                            break
                    if has_data:
                        break
                if has_data:
                    break
            
            if has_data:
                wb.save(dest_path)
                logger.info(f"使用openpyxl转换为xlsx（已读取公式计算值）")
                return True
            else:
                logger.warning(f"openpyxl读取的数据全为空，可能文件未被Excel打开过，尝试其他方案")
        except Exception as e:
            logger.warning(f"openpyxl data_only模式失败: {e}")
        
        # 备用方案2：先复制再用openpyxl打开保存（强制转换，保留公式）
        try:
            wb = load_workbook(source_path, data_only=False, keep_vba=False)
            wb.save(dest_path)
            logger.warning(f"使用openpyxl转换为xlsx（公式已保留但未计算，用户打开时会自动计算）")
            logger.warning(f"  → 如果看到数据为0，请双击打开Excel文件，数据会自动计算")
            return True
        except Exception as e:
            logger.error(f"openpyxl强制转换也失败: {e}")
        
        # 最后的备用方案：修改扩展名强制为xlsx
        logger.warning(f"所有转换方法都失败，将直接复制文件并修改扩展名为.xlsx")
        logger.warning(f"  → 文件可能需要用Excel打开后重新保存以正确显示")
        shutil.copy2(source_path, dest_path)
        return True
            
    except Exception as e:
        logger.error(f"复制/转换源Excel失败: {e}")
        return False


def auto_generate_date_xlsx(data_folder: str = "data", output_file: str = "date.xlsx") -> bool:
    """
    自动从data文件夹生成date.xlsx文件
    返回True表示成功，False表示失败
    """
    try:
        from auto_mapping import create_date_xlsx_from_data
        create_date_xlsx_from_data(data_folder, output_file)
        return True
    except Exception as e:
        print(f"自动生成date.xlsx失败: {e}")
        return False

def main(
    template_path: str = "模板文件/template_审核报告.docx",
    output_pattern: str = "2_工程竣工决算审核报告_{项目名称}.docx",
    extra_templates: Dict[str, str] | None = None,
    excel_templates: Dict[str, str] | None = None,
    strict_date: bool = False,
    debug: bool = False,
    data_file: str = "date.xlsx",  # 默认使用date.xlsx
    output_dir: str = "生成结果",
    auto_mapping: bool = True,  # 新增参数：是否启用自动映射
    workers: int = 1,  # 并行渲染线程数（>1 时禁用 GUI 与 Excel COM）
    skip_existing: bool = False,  # 若目标输出已存在则跳过（断点续跑）
    projects_workers: int = 1,  # 跨项目并行个数（>1 时禁用 GUI 与 Excel COM）
) -> None:
    logger = setup_logging(output_dir=output_dir, debug=debug)
    logger.info("=== 工程竣工决算文档生成程序 ===")

    # 记录是否需要清理date.xlsx文件
    should_cleanup_date_xlsx = False

    # 自动映射功能：如果启用且"待处理数据"文件夹存在，则自动生成date.xlsx
    data_folder = "待处理数据"
    if auto_mapping and os.path.exists(data_folder) and data_file == "date.xlsx":
        logger.info(f"启用自动映射功能，正在从{data_folder}文件夹生成date.xlsx...")
        if auto_generate_date_xlsx(data_folder, "date.xlsx"):
            logger.info("自动映射完成，date.xlsx已生成")
            should_cleanup_date_xlsx = True  # 标记需要清理
        else:
            logger.warning("自动映射失败，将使用现有的date.xlsx文件（如果存在）")

    if not os.path.exists(data_file):
        logger.error(f"未找到数据文件 {data_file}")
        return

    try:
        df = pd.read_excel(data_file)
        logger.info(f"已读取 {len(df)} 行数据")
    except Exception as e:
        logger.exception(f"读取数据失败: {e}")
        return

    # 基础校验：关键列存在性
    required_columns = {"项目名称"}
    missing_required = [c for c in required_columns if c not in df.columns]
    if missing_required:
        logger.warning(f"数据缺少关键列: {missing_required}，可能影响输出命名或渲染")

    if not os.path.exists(template_path):
        logger.error(f"未找到主模板 {template_path}")
        return
    logger.info(f"加载主模板: {template_path}")

    os.makedirs(output_dir, exist_ok=True)
    debug_dir = Path(output_dir) / "debug"
    debug_dir.mkdir(parents=True, exist_ok=True)
    logger.info(f"输出目录: {output_dir}")

    # 创建增强版GUI进度窗口（如果适用；并行时禁用GUI）
    progress_window = None
    start_time = None
    if workers == 1 and projects_workers == 1 and getattr(sys, 'frozen', False) and 'tk' in globals() and tk:
        try:
            import time
            start_time = time.time()
            
            progress_window = tk.Tk()
            progress_window.title("工程竣工决算文档生成工具 v2.0")
            progress_window.geometry("600x280")
            progress_window.resizable(False, False)
            progress_window.configure(bg='#f0f0f0')
            
            # 标题
            title_label = tk.Label(
                progress_window, 
                text="正在批量生成工程竣工决算文档",
                font=('微软雅黑', 14, 'bold'),
                bg='#f0f0f0',
                fg='#2c3e50'
            )
            title_label.pack(pady=(15, 10))
            
            # 总体进度信息框架
            info_frame = tk.Frame(progress_window, bg='#f0f0f0')
            info_frame.pack(pady=5, fill='x', padx=30)
            
            # 总项目数
            total_label = tk.Label(
                info_frame,
                text=f"总项目数：{len(df)} 个",
                font=('微软雅黑', 10),
                bg='#f0f0f0',
                fg='#34495e'
            )
            total_label.pack(anchor='w')
            
            # 当前项目
            current_label = tk.Label(
                info_frame,
                text="当前项目：准备中...",
                font=('微软雅黑', 10),
                bg='#f0f0f0',
                fg='#34495e'
            )
            current_label.pack(anchor='w', pady=3)
            
            # 处理步骤
            step_label = tk.Label(
                info_frame,
                text="处理步骤：初始化",
                font=('微软雅黑', 10),
                bg='#f0f0f0',
                fg='#16a085'
            )
            step_label.pack(anchor='w')
            
            # 进度条框架
            progress_frame = tk.Frame(progress_window, bg='#f0f0f0')
            progress_frame.pack(pady=10, fill='x', padx=30)
            
            # 进度条
            progress = ttk.Progressbar(
                progress_frame, 
                length=540, 
                mode='determinate',
                maximum=len(df)
            )
            progress.pack()
            
            # 百分比标签
            percent_label = tk.Label(
                progress_frame,
                text="0%",
                font=('微软雅黑', 9),
                bg='#f0f0f0',
                fg='#7f8c8d'
            )
            percent_label.pack(pady=3)
            
            # 状态信息框架
            status_frame = tk.Frame(progress_window, bg='#f0f0f0')
            status_frame.pack(pady=5, fill='x', padx=30)
            
            # 已完成/总数
            count_label = tk.Label(
                status_frame,
                text="已完成：0 / {}".format(len(df)),
                font=('微软雅黑', 9),
                bg='#f0f0f0',
                fg='#7f8c8d'
            )
            count_label.pack(side='left')
            
            # 已用时间
            time_label = tk.Label(
                status_frame,
                text="已用时间：0秒",
                font=('微软雅黑', 9),
                bg='#f0f0f0',
                fg='#7f8c8d'
            )
            time_label.pack(side='left', padx=20)
            
            # 预计剩余
            eta_label = tk.Label(
                status_frame,
                text="预计剩余：计算中...",
                font=('微软雅黑', 9),
                bg='#f0f0f0',
                fg='#7f8c8d'
            )
            eta_label.pack(side='left')
            
            # 处理速度
            speed_label = tk.Label(
                progress_window,
                text="处理速度：-- 项目/分钟",
                font=('微软雅黑', 9),
                bg='#f0f0f0',
                fg='#95a5a6'
            )
            speed_label.pack(pady=5)
            
            # 将窗口置于最前
            progress_window.lift()
            progress_window.attributes('-topmost', True)
            progress_window.after_idle(progress_window.attributes, '-topmost', False)
            
            # 居中显示
            progress_window.update_idletasks()
            width = progress_window.winfo_width()
            height = progress_window.winfo_height()
            x = (progress_window.winfo_screenwidth() // 2) - (width // 2)
            y = (progress_window.winfo_screenheight() // 2) - (height // 2)
            progress_window.geometry(f'{width}x{height}+{x}+{y}')
            
            # 显示窗口
            progress_window.update()
        except Exception as e:
            logger.warning(f"创建GUI进度窗口失败: {e}")
            progress_window = None
            start_time = None

    # 允许 COM/GUI 的条件（两级并行均需为1）
    allow_com = (workers == 1 and projects_workers == 1)
    if workers > 1 or projects_workers > 1:
        logger.info(f"启用并行渲染（workers={workers}, projects_workers={projects_workers}），GUI与Excel COM将被禁用以确保稳定性")

    # 处理每一行数据（支持跨项目并行）
    for idx, row in enumerate(df.iterrows()):
        # 更新进度窗口 - 开始处理
        if progress_window and start_time:
            try:
                import time
                current_time = time.time()
                elapsed = current_time - start_time
                
                # 获取项目名称
                _, row_data_preview = row
                project_name = row_data_preview.get("项目名称", f"项目_{idx+1}")
                if len(str(project_name)) > 30:
                    project_name = str(project_name)[:30] + "..."
                
                # 更新当前项目
                current_label.config(text=f"当前项目：{idx+1}/{len(df)} - {project_name}")
                step_label.config(text="处理步骤：读取数据并构建上下文")
                
                # 更新进度条
                progress['value'] = idx
                percent = int((idx / len(df)) * 100)
                percent_label.config(text=f"{percent}%")
                
                # 更新计数
                count_label.config(text=f"已完成：{idx} / {len(df)}")
                
                # 计算时间
                if elapsed < 60:
                    time_str = f"{int(elapsed)}秒"
                else:
                    time_str = f"{int(elapsed//60)}分{int(elapsed%60)}秒"
                time_label.config(text=f"已用时间：{time_str}")
                
                # 预计剩余时间
                if idx > 0:
                    avg_time = elapsed / idx
                    remaining = avg_time * (len(df) - idx)
                    if remaining < 60:
                        eta_str = f"{int(remaining)}秒"
                    else:
                        eta_str = f"{int(remaining//60)}分{int(remaining%60)}秒"
                    eta_label.config(text=f"预计剩余：{eta_str}")
                    
                    # 处理速度
                    speed = (idx / elapsed) * 60  # 项目/分钟
                    speed_label.config(text=f"处理速度：{speed:.1f} 项目/分钟")
                
                progress_window.update()
            except Exception as e:
                logger.warning(f"更新GUI进度窗口失败: {e}")
        
        _, row_data = row
        row_dict = row_data.to_dict()
        logger.info(f"开始处理第 {idx+1} 行数据")

        try:
            # 更新步骤：加载模板
            if progress_window and start_time:
                try:
                    step_label.config(text="处理步骤：加载主模板并构建上下文")
                    progress_window.update()
                except:
                    pass
            
            tpl = DocxTemplate(template_path)
            register_jinja_filters(tpl)
            # 旧版 docxtpl 可能无 get_jinja_env，直接依赖 build_context 的格式化
            context = build_context(row_dict, tpl, debug=debug, strict_date=strict_date)

            # 预校验：缺失占位符
            try:
                missing_vars = tpl.get_undeclared_template_variables(context)
                if missing_vars:
                    logger.warning(f"第{idx+1}行：主模板缺少数据字段：{sorted(missing_vars)}")
            except Exception:
                pass

            tpl.render(context)

            # 清理文件名中的非法字符
            safe_context = {k: sanitize_filename(v) if isinstance(v, str) else v for k, v in context.items()}

            # 为每个项目创建子目录：项目报告号_项目名称（若缺失报告号则仅用项目名称）
            report_no = safe_context.get("项目报告号") or ""
            proj_name = safe_context.get("项目名称") or safe_context.get("name") or f"项目_{idx+1}"
            if report_no:
                sub_dir_name = sanitize_filename(f"{report_no}_{proj_name}")
            else:
                sub_dir_name = sanitize_filename(f"{proj_name}")
            project_output_dir = os.path.join(output_dir, sub_dir_name)
            os.makedirs(project_output_dir, exist_ok=True)

            try:
                filename = os.path.join(project_output_dir, output_pattern.format(**safe_context))
            except Exception:
                name = context.get("项目名称") or context.get("name") or f"输出_{idx+1}"
                filename = os.path.join(project_output_dir, f"{sanitize_filename(name)}.docx")

            # 断点续跑：若所有目标输出均已存在，直接跳过本项目
            if skip_existing:
                expected_files = [filename]
                if extra_templates:
                    for out_pattern in extra_templates.keys():
                        try:
                            expected_files.append(os.path.join(project_output_dir, out_pattern.format(**safe_context)))
                        except Exception:
                            pass
                if excel_templates:
                    for out_pattern in excel_templates.keys():
                        try:
                            expected_files.append(os.path.join(project_output_dir, out_pattern.format(**safe_context)))
                        except Exception:
                            pass
                # 源数据转换后的目标文件
                proj_name_safe = safe_context.get("项目名称") or safe_context.get("name") or f"项目_{idx+1}"
                expected_files.append(os.path.join(project_output_dir, f"6_基建工程竣工决算报表_{proj_name_safe}.xlsx"))
                if expected_files and all(os.path.exists(p) for p in expected_files):
                    logger.info(f"第{idx+1}行：检测到所有输出已存在（--skip-existing），跳过处理")
                    continue

            try:
                # 更新步骤：生成主文档
                if progress_window and start_time:
                    try:
                        step_label.config(text="处理步骤：生成主文档（审核报告）")
                        progress_window.update()
                    except:
                        pass
                
                tpl.save(filename)
                # 渲染后按上下文清理空值/0的整行
                cleanup_document_rows_by_context(filename, context, logger)
                # 删除空值单位行（仅针对审核报告）
                remove_empty_unit_lines(filename, context, logger)
                # 标题：就地断行但保持段落原样式（用模板渲染出的纯文本做断行）
                try:
                    split_title_keep_style(filename, str(context.get("标题项目名称_plain", "")), 12, logger)
                except Exception:
                    pass
                logger.info(f"主文档生成完成: {filename}")
            except PermissionError:
                import time
                timestamp = int(time.time())
                base_name, ext = os.path.splitext(filename)
                filename = f"{base_name}_{timestamp}{ext}"
                tpl.save(filename)
                logger.warning(f"文件被占用，另存为: {filename}")

            # 额外模板
            if extra_templates:
                # 更新步骤：生成额外文档
                if progress_window and start_time:
                    try:
                        step_label.config(text="处理步骤：生成附加文档（封面、目录、说明书）")
                        progress_window.update()
                    except:
                        pass
                def _render_extra(out_pattern: str, tpl_path: str) -> None:
                    if not os.path.exists(tpl_path):
                        logger.warning(f"跳过：未找到额外模板 {tpl_path}")
                        return
                    try:
                        tpl2 = DocxTemplate(tpl_path)
                        register_jinja_filters(tpl2)
                        context2 = build_context(row_dict, tpl2, debug=debug, strict_date=strict_date)
                        try:
                            missing_vars2 = tpl2.get_undeclared_template_variables(context2)
                            if missing_vars2:
                                logger.warning(f"第{idx+1}行：额外模板缺少数据字段：{sorted(missing_vars2)} → {tpl_path}")
                        except Exception:
                            pass
                        tpl2.render(context2)
                        safe_context2 = {k: sanitize_filename(v) if isinstance(v, str) else v for k, v in context2.items()}
                        fname2 = os.path.join(project_output_dir, out_pattern.format(**safe_context2))
                        tpl2.save(fname2)
                        cleanup_document_rows_by_context(fname2, context2, logger)
                        try:
                            split_title_keep_style(fname2, str(context2.get("标题项目名称_plain", "")), 12, logger)
                        except Exception:
                            pass
                        logger.info(f"额外文档生成完成: {fname2}")
                        try:
                            if os.path.basename(tpl_path) == "template_报告目录.docx":
                                update_catalog_when_initial_missing(fname2, context2, logger)
                            if os.path.basename(tpl_path) == "template_说明书.docx":
                                remove_init_phrase_in_manual(fname2, context2, logger)
                        except Exception:
                            pass
                    except Exception as e:
                        logger.exception(f"第{idx+1}行：生成额外模板失败（{tpl_path}）: {e}")
                if workers > 1:
                    with ThreadPoolExecutor(max_workers=workers) as ex:
                        futures = [ex.submit(_render_extra, out_pattern, tpl_path) for out_pattern, tpl_path in extra_templates.items()]
                        for _f in futures:
                            try:
                                _f.result()
                            except Exception:
                                pass
                else:
                    for out_pattern, tpl_path in extra_templates.items():
                        _render_extra(out_pattern, tpl_path)

            # Excel 模板
            if excel_templates:
                # 更新步骤：生成Excel
                if progress_window and start_time:
                    try:
                        step_label.config(text="处理步骤：生成Excel附件（审核表）")
                        progress_window.update()
                    except:
                        pass
                def _render_excel(out_pattern: str, tpl_path: str) -> None:
                    if not os.path.exists(tpl_path):
                        logger.warning(f"跳过：未找到Excel模板 {tpl_path}")
                        return
                    try:
                        safe_context3 = {k: sanitize_filename(v) if isinstance(v, str) else v for k, v in context.items()}
                        fname_excel = os.path.join(project_output_dir, out_pattern.format(**safe_context3))
                        render_excel_template(tpl_path, context, fname_excel)
                        if not force_recalc_with_excel(fname_excel, allow_com=allow_com):
                            logger.warning(f"Excel公式计算未执行或失败（并行/占用场景可能禁用COM），文件已生成: {fname_excel}")
                        logger.info(f"Excel 附件生成完成: {fname_excel}")
                    except Exception as e:
                        logger.exception(f"第{idx+1}行：生成Excel附件失败（{tpl_path}）: {e}")
                if workers > 1:
                    with ThreadPoolExecutor(max_workers=workers) as ex:
                        futures = [ex.submit(_render_excel, out_pattern, tpl_path) for out_pattern, tpl_path in excel_templates.items()]
                        for _f in futures:
                            try:
                                _f.result()
                            except Exception:
                                pass
                else:
                    for out_pattern, tpl_path in excel_templates.items():
                        _render_excel(out_pattern, tpl_path)

            # 将源数据表转换/复制为 .xlsx 并放入对应项目输出目录
            try:
                # 更新步骤：转换源数据
                if progress_window and start_time:
                    try:
                        step_label.config(text="处理步骤：转换并保存源数据表")
                        progress_window.update()
                    except:
                        pass
                
                source_path = row_dict.get("源文件路径")
                if isinstance(source_path, str) and os.path.exists(source_path):
                    # 使用项目名称作为输出文件名，统一转换为.xlsx格式
                    proj_name_safe = safe_context.get("项目名称") or safe_context.get("name") or f"项目_{idx+1}"
                    dest_filename = f"6_基建工程竣工决算报表_{proj_name_safe}.xlsx"
                    dest_file_path = os.path.join(project_output_dir, dest_filename)
                    
                    if convert_excel_to_xlsx(source_path, dest_file_path, logger, allow_com=allow_com):
                        logger.info(f"源数据已转换并保存: {dest_file_path}")
                    else:
                        logger.warning(f"源数据转换失败，已跳过: {source_path}")
            except Exception as e:
                logger.warning(f"复制源表至项目目录失败: {e}")

            if debug:
                try:
                    with open(debug_dir / f"row_{idx+1}.json", "w", encoding="utf-8") as f:
                        json.dump(context, f, ensure_ascii=False, indent=2, default=str)
                except Exception:
                    logger.warning(f"第{idx+1}行：写入调试上下文失败")

        except Exception as e:
            logger.exception(f"第{idx+1}行：处理失败: {e}")
        
        # 更新完成状态
        if progress_window and start_time:
            try:
                import time
                current_time = time.time()
                elapsed = current_time - start_time
                
                # 更新进度条到当前完成数
                progress['value'] = idx + 1
                percent = int(((idx + 1) / len(df)) * 100)
                percent_label.config(text=f"{percent}%")
                count_label.config(text=f"已完成：{idx + 1} / {len(df)}")
                
                # 更新时间
                if elapsed < 60:
                    time_str = f"{int(elapsed)}秒"
                else:
                    time_str = f"{int(elapsed//60)}分{int(elapsed%60)}秒"
                time_label.config(text=f"已用时间：{time_str}")
                
                # 预计剩余时间
                if idx + 1 < len(df):
                    avg_time = elapsed / (idx + 1)
                    remaining = avg_time * (len(df) - idx - 1)
                    if remaining < 60:
                        eta_str = f"{int(remaining)}秒"
                    else:
                        eta_str = f"{int(remaining//60)}分{int(remaining%60)}秒"
                    eta_label.config(text=f"预计剩余：{eta_str}")
                else:
                    eta_label.config(text="预计剩余：0秒")
                
                # 处理速度
                speed = ((idx + 1) / elapsed) * 60  # 项目/分钟
                speed_label.config(text=f"处理速度：{speed:.1f} 项目/分钟")
                
                step_label.config(text="处理步骤：已完成当前项目")
                progress_window.update()
            except:
                pass

    # 显示完成消息并延迟关闭窗口
    if progress_window and start_time:
        try:
            import time
            step_label.config(text="✓ 全部处理完成！", fg='#27ae60')
            current_label.config(text=f"已成功生成 {len(df)} 个项目的文档")
            
            # 计算总耗时
            total_elapsed = time.time() - start_time
            if total_elapsed < 60:
                time_str = f"{int(total_elapsed)}秒"
            else:
                time_str = f"{int(total_elapsed//60)}分{int(total_elapsed%60)}秒"
            time_label.config(text=f"总用时：{time_str}")
            eta_label.config(text="")
            
            progress_window.update()
            # 延迟3秒后自动关闭
            progress_window.after(3000, progress_window.destroy)
        except Exception as e:
            logger.warning(f"更新完成状态失败: {e}")
    elif progress_window:
        try:
            progress_window.destroy()
        except Exception as e:
            logger.warning(f"关闭GUI进度窗口失败: {e}")
    
    # 清理：按本次是否自动生成来决定是否删除 date.xlsx；与运行时数据目录保持一致
    try:
        if should_cleanup_date_xlsx and os.path.exists("date.xlsx"):
            os.remove("date.xlsx")
            logger.info("已清理临时文件 date.xlsx")
    except Exception as e:
        logger.warning(f"清理 date.xlsx 文件失败: {e}")

    try:
        if os.path.isdir(data_folder):
            # 清空“待处理数据”文件夹
            # 保留原始Excel文件，不清空"待处理数据"文件夹
            # for name in os.listdir(data_folder):
            #     path = os.path.join(data_folder, name)
            #     try:
            #         if os.path.isdir(path):
            #             shutil.rmtree(path, ignore_errors=True)
            #         else:
            #             os.remove(path)
            #     except Exception:
            #         pass
            logger.info(f"已保留 {data_folder} 目录中的原始文件")
    except Exception as e:
        logger.warning(f"处理 {data_folder} 目录失败: {e}")
    
    logger.info("=== 全部处理完成 ===")


def parse_cli_and_run() -> None:
    parser = argparse.ArgumentParser(description="工程竣工决算文档生成程序")
    parser.add_argument("--data", default="date.xlsx", help="数据表路径（默认: date.xlsx）")
    parser.add_argument("--template", default="模板文件/template_审核报告.docx", help="主模板路径")
    parser.add_argument("--output-dir", default="生成结果", help="输出目录（默认: 生成结果）")
    parser.add_argument("--output-pattern", default="2_工程竣工决算审核报告_{项目名称}.docx", help="主文档输出命名格式")
    parser.add_argument("--strict-date", action="store_true", help="严格日期解析模式")
    parser.add_argument("--debug", action="store_true", help="调试模式，输出更多日志与行上下文")
    parser.add_argument("--no-auto-mapping", action="store_true", help="禁用自动映射功能，使用现有的数据文件")
    parser.add_argument("--workers", type=int, default=1, help="并行渲染线程数（>1 将禁用GUI与Excel COM）")
    parser.add_argument("--skip-existing", action="store_true", help="若输出已存在则跳过（断点续跑）")
    parser.add_argument("--projects-workers", type=int, default=1, help="跨项目并行个数（>1 将禁用GUI与Excel COM）")

    args = parser.parse_args()

    # 与 __main__ 中默认配置保持兼容：额外和Excel模板固定为当前内置值
    main(
        template_path=args.template,
        output_pattern=args.output_pattern,
        extra_templates={
            "1_报告封面_{项目名称}.docx": "模板文件/template_报告封面.docx",
            "4_报告目录_{项目名称}.docx": "模板文件/template_报告目录.docx",
            "5_竣工决算报告说明书_{项目名称}.docx": "模板文件/template_说明书.docx",
        },
        excel_templates={
            "3_审核表_{项目名称}.xlsx": "模板文件/template_审核表.xlsx",
        },
        strict_date=args.strict_date,
        debug=args.debug,
        data_file=args.data,
        output_dir=args.output_dir,
        auto_mapping=not args.no_auto_mapping,  # 默认启用自动映射
        workers=max(1, args.workers),
        skip_existing=args.skip_existing,
        projects_workers=max(1, args.projects_workers),
    )


if __name__ == "__main__":
    # 若不传参，行为与原来一致；传参则覆盖默认
    parse_cli_and_run()
