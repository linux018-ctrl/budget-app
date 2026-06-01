import io
import re
import xml.etree.ElementTree as ET
import pandas as pd
from datetime import datetime


def inspect_and_repair_xml_bytes(xml_bytes):
    """
    檢查並修復 XML bytes，回傳 (repaired_bytes, report)

    report 範例:
    {
        "original_valid": False,
        "repaired_valid": True,
        "original_error": "...",
        "repaired_error": "",
        "replaced_amp_count": 3,
        "removed_control_count": 0
    }
    """
    report = {
        "original_valid": False,
        "repaired_valid": False,
        "original_error": "",
        "repaired_error": "",
        "replaced_amp_count": 0,
        "removed_control_count": 0,
    }

    # 先測試原始 XML 是否有效
    try:
        ET.fromstring(xml_bytes)
        report["original_valid"] = True
    except ET.ParseError as e:
        report["original_error"] = str(e)

    text = xml_bytes.decode('utf-8-sig', errors='ignore')

    control_pattern = r'[\x00-\x08\x0B\x0C\x0E-\x1F]'
    amp_pattern = r'&(?!#?\w+;)'

    report["removed_control_count"] = len(re.findall(control_pattern, text))
    report["replaced_amp_count"] = len(re.findall(amp_pattern, text))

    repaired_text = re.sub(control_pattern, '', text)
    repaired_text = re.sub(amp_pattern, '&amp;', repaired_text)
    repaired_bytes = repaired_text.encode('utf-8')

    # 再測試修復後 XML 是否有效
    try:
        ET.fromstring(repaired_bytes)
        report["repaired_valid"] = True
    except ET.ParseError as e:
        report["repaired_error"] = str(e)

    return repaired_bytes, report


def _parse_xml_safely(xml_bytes):
    """
    容錯 XML 解析：
    1) 移除非法控制字元
    2) 修正常見未轉義的 '&'
    """
    # 先嘗試原始解析（最快）
    try:
        return ET.fromstring(xml_bytes)
    except Exception:
        pass

    text = xml_bytes.decode('utf-8-sig', errors='ignore')
    # 移除 XML 1.0 不允許的控制字元（保留 \t \n \r）
    text = re.sub(r'[\x00-\x08\x0B\x0C\x0E-\x1F]', '', text)
    # 修正常見的未轉義 '&'（例如備註中有 A&B）
    text = re.sub(r'&(?!#?\w+;)', '&amp;', text)

    return ET.fromstring(text)

def parse_cwmoney_xml(xml_bytes, year=None, month=None):
    """
    解析 CWMoney 匯出的 XML 檔案，回傳 DataFrame
    """
    try:
        root = _parse_xml_safely(xml_bytes)
    except ET.ParseError as e:
        raise ValueError(f"XML 格式錯誤，無法解析：{e}")

    # 嘗試判斷是否為 SpreadsheetML (Excel XML)
    ns = {
        'ss': 'urn:schemas-microsoft-com:office:spreadsheet',
        'default': 'urn:schemas-microsoft-com:office:spreadsheet'
    }
    worksheet = root.find('.//ss:Worksheet[@ss:Name="Detail"]', ns)
    if worksheet is not None:
        table = worksheet.find('.//ss:Table', ns)
        rows = table.findall('.//ss:Row', ns)
        data = []
        headers = []
        for i, row in enumerate(rows):
            cells = row.findall('.//ss:Cell', ns)
            values = []
            for cell in cells:
                data_elem = cell.find('.//ss:Data', ns)
                if data_elem is not None:
                    values.append(data_elem.text)
                else:
                    values.append("")
            if i == 0:
                headers = values
            else:
                # 補齊缺漏欄位
                while len(values) < len(headers):
                    values.append("")
                data.append(values)
        df = pd.DataFrame(data, columns=headers)
        # 可選：依 year/month 過濾
        if year or month:
            def date_filter(row):
                try:
                    d = pd.to_datetime(row['日期']).date()
                except Exception:
                    return False
                if year and d.year != year:
                    return False
                if month and d.month != month:
                    return False
                return True
            df = df[df.apply(date_filter, axis=1)]
        return df.reset_index(drop=True)

    # fallback: 舊版 <Record> 格式
    records = []
    for rec in root.findall('.//Record'):
        date_str = rec.get('Date')
        try:
            record_date = datetime.strptime(date_str, "%Y-%m-%d").date()
        except Exception:
            continue
        if year and record_date.year != year:
            continue
        if month and record_date.month != month:
            continue
        records.append({
            'date': record_date.isoformat(),
            'type': rec.get('Type', ''),
            'main_category': rec.get('MainClass', ''),
            'sub_category': rec.get('SubClass', ''),
            'account': rec.get('Account', ''),
            'project': rec.get('Project', ''),
            'amount': float(rec.get('Money', 0)),
            'note': rec.get('Note', ''),
            'location': rec.get('Address', ''),
            'invoice': rec.get('Invoice', ''),
        })
    return pd.DataFrame(records)
