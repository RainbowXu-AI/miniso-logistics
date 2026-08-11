#!/usr/bin/env python3
"""Process Tracking Report .xls into JSON for the dashboard.
Reads from data/ directory, outputs logistics_data.json in root.
"""
import xlrd
import json
import os
import glob
from datetime import datetime, timedelta
from collections import Counter

DATA_DIR = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), 'data')
OUTPUT_FILE = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), 'logistics_data.json')

COUNTRY_CN = {
    'AD':'安道尔','AE':'阿联酋','AF':'阿富汗','AG':'安提瓜和巴布达','AI':'安圭拉','AL':'阿尔巴尼亚','AM':'亚美尼亚','AO':'安哥拉','AR':'阿根廷','AT':'奥地利','AU':'澳大利亚','AW':'阿鲁巴','AZ':'阿塞拜疆',
    'BA':'波黑','BB':'巴巴多斯','BD':'孟加拉国','BE':'比利时','BF':'布基纳法索','BG':'保加利亚','BH':'巴林','BI':'布隆迪','BJ':'贝宁','BM':'百慕大','BN':'文莱','BO':'玻利维亚','BR':'巴西','BS':'巴哈马','BT':'不丹','BW':'博茨瓦纳','BY':'白俄罗斯','BZ':'伯利兹',
    'CA':'加拿大','CD':'刚果(金)','CF':'中非','CG':'刚果(布)','CH':'瑞士','CI':'科特迪瓦','CK':'库克群岛','CL':'智利','CM':'喀麦隆','CN':'中国','CO':'哥伦比亚','CR':'哥斯达黎加','CU':'古巴','CV':'佛得角','CY':'塞浦路斯','CZ':'捷克',
    'DE':'德国','DJ':'吉布提','DK':'丹麦','DO':'多米尼加','DZ':'阿尔及利亚',
    'EC':'厄瓜多尔','EE':'爱沙尼亚','EG':'埃及','ER':'厄立特里亚','ES':'西班牙','ET':'埃塞俄比亚',
    'FI':'芬兰','FJ':'斐济','FR':'法国',
    'GA':'加蓬','GB':'英国','GD':'格林纳达','GE':'格鲁吉亚','GH':'加纳','GM':'冈比亚','GN':'几内亚','GR':'希腊','GT':'危地马拉','GW':'几内亚比绍','GY':'圭亚那',
    'HK':'中国香港','HN':'洪都拉斯','HR':'克罗地亚','HT':'海地','HU':'匈牙利',
    'ID':'印度尼西亚','IE':'爱尔兰','IL':'以色列','IN':'印度','IQ':'伊拉克','IR':'伊朗','IS':'冰岛','IT':'意大利',
    'JM':'牙买加','JO':'约旦','JP':'日本',
    'KE':'肯尼亚','KG':'吉尔吉斯斯坦','KH':'柬埔寨','KI':'基里巴斯','KM':'科摩罗','KN':'圣基茨和尼维斯','KP':'朝鲜','KR':'韩国','KW':'科威特','KY':'开曼群岛','KZ':'哈萨克斯坦',
    'LA':'老挝','LB':'黎巴嫩','LC':'圣卢西亚','LI':'列支敦士登','LK':'斯里兰卡','LR':'利比里亚','LS':'莱索托','LT':'立陶宛','LU':'卢森堡','LV':'拉脱维亚','LY':'利比亚',
    'MA':'摩洛哥','MD':'摩尔多瓦','ME':'黑山','MG':'马达加斯加','MK':'北马其顿','ML':'马里','MM':'缅甸','MN':'蒙古','MO':'中国澳门','MT':'马耳他','MU':'毛里求斯','MV':'马尔代夫','MW':'马拉维','MX':'墨西哥','MY':'马来西亚','MZ':'莫桑比克',
    'NA':'纳米比亚','NE':'尼日尔','NG':'尼日利亚','NI':'尼加拉瓜','NL':'荷兰','NO':'挪威','NP':'尼泊尔','NZ':'新西兰',
    'OM':'阿曼',
    'PA':'巴拿马','PE':'秘鲁','PG':'巴布亚新几内亚','PH':'菲律宾','PK':'巴基斯坦','PL':'波兰','PR':'波多黎各','PT':'葡萄牙','PY':'巴拉圭',
    'QA':'卡塔尔',
    'RE':'留尼汪','RO':'罗马尼亚','RS':'塞尔维亚','RU':'俄罗斯','RW':'卢旺达',
    'SA':'沙特阿拉伯','SB':'所罗门群岛','SC':'塞舌尔','SD':'苏丹','SE':'瑞典','SG':'新加坡','SI':'斯洛文尼亚','SK':'斯洛伐克','SL':'塞拉利昂','SN':'塞内加尔','SO':'索马里','SR':'苏里南','SS':'南苏丹','SV':'萨尔瓦多','SY':'叙利亚','SZ':'斯威士兰',
    'TD':'乍得','TG':'多哥','TH':'泰国','TJ':'塔吉克斯坦','TM':'土库曼斯坦','TN':'突尼斯','TO':'汤加','TR':'土耳其','TT':'特立尼达和多巴哥','TW':'中国台湾','TZ':'坦桑尼亚',
    'UA':'乌克兰','UG':'乌干达','US':'美国','UY':'乌拉圭','UZ':'乌兹别克斯坦',
    'VA':'梵蒂冈','VC':'圣文森特和格林纳丁斯','VE':'委内瑞拉','VN':'越南','VU':'瓦努阿图',
    'WS':'萨摩亚',
    'YE':'也门',
    'ZA':'南非','ZM':'赞比亚','ZW':'津巴布韦',
    'UNKNOWN':'未知'
}

CONTINENT_CN = {
    'Asia': '亚洲', 'LATAM': '拉丁美洲', 'NA': '北美', 'EU': '欧洲',
    'AFR': '非洲', 'Oceania': '大洋洲', '': '未知'
}

def excel_date(serial):
    if not serial or serial == '':
        return ''
    try:
        epoch = datetime(1899, 12, 30)
        d = epoch + timedelta(days=float(serial))
        return d.strftime('%Y-%m-%d')
    except (ValueError, TypeError):
        return ''

def excel_datetime(serial):
    if not serial or serial == '':
        return ''
    try:
        epoch = datetime(1899, 12, 30)
        d = epoch + timedelta(days=float(serial))
        return d.strftime('%Y-%m-%d %H:%M')
    except (ValueError, TypeError):
        return ''

def iso_week(date_str):
    if not date_str:
        return ''
    try:
        d = datetime.strptime(date_str[:10], '%Y-%m-%d')
        iso = d.isocalendar()
        return f'{iso[0]}-W{iso[1]:02d}'
    except (ValueError, TypeError):
        return ''

def days_between(d1, d2):
    if not d1 or not d2:
        return None
    try:
        dd1 = datetime.strptime(d1[:10], '%Y-%m-%d')
        dd2 = datetime.strptime(d2[:10], '%Y-%m-%d')
        return (dd2 - dd1).days
    except (ValueError, TypeError):
        return None

def to_int(val):
    try:
        return int(float(val))
    except (ValueError, TypeError):
        return None

def find_excel_file():
    patterns = ['*.xls', '*.xlsx']
    for p in patterns:
        files = glob.glob(os.path.join(DATA_DIR, p))
        if files:
            return files[0]
    return None

def main():
    input_file = find_excel_file()
    if not input_file:
        print('No Excel file found in data/ directory!')
        return
    print(f'Processing: {input_file}')

    wb = xlrd.open_workbook(input_file)
    s = wb.sheet_by_index(0)
    print(f'Reading {s.nrows - 1} rows, {s.ncols} cols')

    headers = [str(s.cell_value(0, c)).replace('\xa0', ' ').strip() for c in range(s.ncols)]

    col_map = {}
    for i, h in enumerate(headers):
        hl = h.lower()
        if 'bkg' in hl or 'bl no' in hl: col_map['bkg'] = i
        elif h == 'Cntr No.': col_map['cntr'] = i
        elif h == 'Weekly': col_map['weekly'] = i
        elif h == 'Carrier': col_map['carrier'] = i
        elif h == 'Cntr Size': col_map['cntrSize'] = i
        elif '1st leg' in hl: col_map['vessel'] = i
        elif h == 'POL': col_map['pol'] = i
        elif 'laden gate' in hl: col_map['gateIn'] = i
        elif 'original etd' in hl and 'pol' in hl: col_map['origEtd'] = i
        elif h == 'POL ETD': col_map['polEtd'] = i
        elif 'pol atd' in hl: col_map['polAtd'] = i
        elif h == 'POD': col_map['pod'] = i
        elif 'original' in hl and 'eta' in hl and 'delivery' in hl: col_map['origEta'] = i
        elif 'updated' in hl and 'eta' in hl and 'delivery' in hl: col_map['updEta'] = i
        elif 'ata' in hl and 'delivery' in hl and 'month' not in hl: col_map['ata'] = i
        elif h == 'Delay': col_map['delay'] = i
        elif 'history' in hl and 'eta' in hl: col_map['etaHistory'] = i
        elif 'original transit' in hl: col_map['origTransit'] = i
        elif 'updated transit' in hl: col_map['updTransit'] = i
        elif 'ata month' in hl: col_map['ataMonth'] = i
        elif h == 'Exceptions': col_map['exceptions'] = i
        elif h == 'Continents': col_map['continents'] = i
        elif 'transportation status' in hl: col_map['transStatus'] = i
        elif 'create time' in hl: col_map['createTime'] = i
        elif 'update time' in hl: col_map['updateTime'] = i
        elif 'final destination' in hl: col_map['finalDest'] = i

    def get_val(row, key, default=''):
        idx = col_map.get(key)
        if idx is None:
            return default
        return row[idx]

    records = []
    stats = Counter()

    for r in range(1, s.nrows):
        row = [s.cell_value(r, c) for c in range(s.ncols)]
        bkg = str(get_val(row, 'bkg')).strip()
        cntr = str(get_val(row, 'cntr')).strip()
        if not bkg and not cntr:
            continue
        carrier = str(get_val(row, 'carrier')).strip()
        cntrSize = str(get_val(row, 'cntrSize')).strip()
        vesselRaw = str(get_val(row, 'vessel')).strip()
        pol = str(get_val(row, 'pol')).strip()
        pod = str(get_val(row, 'pod')).strip()
        vesselName = ''
        vesselVoy = ''
        if vesselRaw:
            parts = vesselRaw.rsplit('/', 1)
            if len(parts) == 2:
                vesselName = parts[0].strip()
                vesselVoy = parts[1].strip()
            else:
                vesselName = vesselRaw
        routeKey = f'{carrier} | {vesselName}' if vesselName and carrier else ''
        origEtd = excel_date(get_val(row, 'origEtd'))
        polEtd = excel_date(get_val(row, 'polEtd'))
        polAtd = excel_date(get_val(row, 'polAtd'))
        origEta = excel_date(get_val(row, 'origEta'))
        updEta = excel_date(get_val(row, 'updEta'))
        ata = excel_date(get_val(row, 'ata'))
        gateIn = excel_datetime(get_val(row, 'gateIn'))
        finalDest = str(get_val(row, 'finalDest')).strip()
        if not finalDest:
            finalDest = 'UNKNOWN'
        countryCN = COUNTRY_CN.get(finalDest, finalDest)
        marketDisplay = f'{finalDest} {countryCN}'
        continents = str(get_val(row, 'continents')).strip()
        continentCN = CONTINENT_CN.get(continents, continents or '未知')
        transStatus = str(get_val(row, 'transStatus')).strip()
        delay = to_int(get_val(row, 'delay'))
        if delay is None and origEta and ata:
            delay = days_between(origEta, ata)
        origTransit = to_int(get_val(row, 'origTransit'))
        if origTransit is None:
            origTransit = days_between(origEtd, origEta)
        updTransit = to_int(get_val(row, 'updTransit'))
        actualTransit = days_between(polEtd, ata)
        if actualTransit is None:
            actualTransit = days_between(polEtd, updEta)
        status = 'transit'
        statusText = '在途'
        if ata:
            if delay is not None and delay > 0:
                status = 'delayed'
                statusText = '延误到港'
            else:
                status = 'arrived'
                statusText = '准班到港'
        elif transStatus in ('Empty Returned', 'AT POD'):
            status = 'arrived'
            statusText = '已到港(无ATA)'
        week = iso_week(polEtd or origEtd)
        ataMonth = ''
        if ata:
            try:
                d = datetime.strptime(ata[:10], '%Y-%m-%d')
                ataMonth = f'{d.year}-{d.month:02d}'
            except:
                pass
        etaHistory = str(get_val(row, 'etaHistory')).strip()
        exceptions = str(get_val(row, 'exceptions')).strip()
        records.append({
            'bkg': bkg, 'cntr': cntr, 'carrier': carrier, 'cntrSize': cntrSize,
            'vessel': vesselRaw, 'vesselName': vesselName, 'vesselVoy': vesselVoy,
            'routeKey': routeKey, 'pol': pol, 'pod': pod,
            'finalDest': finalDest, 'countryCN': countryCN, 'marketDisplay': marketDisplay,
            'continents': continents, 'continentCN': continentCN,
            'transStatus': transStatus, 'gateIn': gateIn,
            'origEtd': origEtd, 'polEtd': polEtd, 'polAtd': polAtd,
            'origEta': origEta, 'updEta': updEta, 'ata': ata,
            'delay': delay, 'origTransit': origTransit, 'updTransit': updTransit,
            'actualTransit': actualTransit, 'status': status, 'statusText': statusText,
            'week': week, 'ataMonth': ataMonth, 'etaHistory': etaHistory, 'exceptions': exceptions
        })
        stats[status] += 1

    original_count = len(records)
    seen = set()
    deduped = []
    for r in records:
        cntr = r.get('cntr', '')
        polAtd = r.get('polAtd', '')
        if cntr and polAtd:
            key = (cntr, polAtd)
            if key in seen:
                continue
            seen.add(key)
        deduped.append(r)
    records = deduped
    dup_removed = original_count - len(records)
    print(f'Dedup: removed {dup_removed} duplicates ({original_count} -> {len(records)})')

    with open(OUTPUT_FILE, 'w', encoding='utf-8') as f:
        json.dump(records, f, ensure_ascii=False)

    print(f'Output: {OUTPUT_FILE} ({len(records)} records)')
    print(f'Status: {dict(stats)}')

if __name__ == '__main__':
    main()
