import io
import xml.etree.ElementTree as ET
import pandas as pd
from datetime import datetime

def parse_cwmoney_xml(xml_bytes, year=None, month=None):
    """
    解析 CWMoney 匯出的 XML 檔案，回傳 DataFrame
    """
    tree = ET.parse(io.BytesIO(xml_bytes))
    root = tree.getroot()

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
