import os
import re
import json
import shutil
import subprocess
from datetime import datetime
import openpyxl

SVOD_PATH       = r'C:\Users\zinov\OneDrive\Desktop\РАБОЧАЯ\001 ИМПОРТ\006 excel учет\СВОД.xlsx'
CONTAINERS_ROOT = r'C:\Users\zinov\OneDrive\Desktop\РАБОЧАЯ\001 ИМПОРТ\001 containers\2026\001 UAE'
REPO_PATH       = r'C:\Users\zinov\OneDrive\Desktop\РАБОЧАЯ\001 ИМПОРТ\007 container tracing'
FILES_DIR       = os.path.join(REPO_PATH, 'files')


def read_tracking():
    wb = openpyxl.load_workbook(SVOD_PATH, data_only=True)
    ws = wb['Трекинг']
    result = []
    for row in ws.iter_rows(min_row=5, values_only=True):
        if not row[0] or not str(row[0]).startswith('UAE'):
            continue
        result.append({
            'name':           str(row[0]).strip(),
            'ship_plan':      row[1],
            'ship_fact':      row[2],
            'customs_plan':   row[3],
            'customs_fact':   row[4],
            'warehouse_plan': row[5],
            'warehouse_fact': row[6],
            'status':         str(row[7]) if row[7] else '',
        })
    return result


def find_container_folder(name):
    # UAE#5 -> папка UAE#5 или UAE#5_ЧТО_УГОДНО
    for folder in os.listdir(CONTAINERS_ROOT):
        if 'DO_NOT_COUNT' in folder.upper():
            continue
        code = folder.split('_')[0]  # UAE#5, UAE#6 из UAE#6_STPARTS
        if code == name:
            return os.path.join(CONTAINERS_ROOT, folder)
    return None


def find_latest_fpz(folder):
    files = [f for f in os.listdir(folder)
             if 'ФПЗ' in f and not f.startswith('~$') and f.lower().endswith('.xlsx')]
    if not files:
        return None

    def sort_key(f):
        m = re.search(r'(\d{2})\.(\d{2})\.(\d{4})', f)
        if m:
            d, mo, y = m.groups()
            return datetime(int(y), int(mo), int(d))
        return datetime.fromtimestamp(os.path.getmtime(os.path.join(folder, f)))

    files.sort(key=sort_key, reverse=True)
    return os.path.join(folder, files[0])


def find_ffz(folder):
    files = [f for f in os.listdir(folder)
             if 'ФФЗ' in f and not f.startswith('~$') and f.lower().endswith('.xlsx')]
    return os.path.join(folder, files[0]) if files else None


def fmt(d):
    if d is None:
        return None
    if isinstance(d, datetime):
        return d.strftime('%Y-%m-%d')
    return str(d)


def calc_delays(c):
    today = datetime.today().replace(hour=0, minute=0, second=0, microsecond=0)
    delays = {}
    pairs = [
        ('ship',      c['ship_plan'],      c['ship_fact']),
        ('customs',   c['customs_plan'],   c['customs_fact']),
        ('warehouse', c['warehouse_plan'], c['warehouse_fact']),
    ]
    for key, plan, fact in pairs:
        if plan and not fact:
            if isinstance(plan, datetime):
                delta = (today - plan).days
                if delta > 0:
                    delays[key] = delta
    return delays


def main():
    os.makedirs(FILES_DIR, exist_ok=True)

    containers = read_tracking()
    output = []

    for c in containers:
        name = c['name']
        folder = find_container_folder(name)

        fpz_url = None
        ffz_url = None

        if folder:
            fpz = find_latest_fpz(folder)
            if fpz:
                dest = os.path.join(FILES_DIR, os.path.basename(fpz))
                shutil.copy2(fpz, dest)
                fpz_url = 'files/' + os.path.basename(fpz)

            ffz = find_ffz(folder)
            if ffz:
                dest = os.path.join(FILES_DIR, os.path.basename(ffz))
                shutil.copy2(ffz, dest)
                ffz_url = 'files/' + os.path.basename(ffz)

        output.append({
            'name':           name,
            'status':         c['status'],
            'ship_plan':      fmt(c['ship_plan']),
            'ship_fact':      fmt(c['ship_fact']),
            'customs_plan':   fmt(c['customs_plan']),
            'customs_fact':   fmt(c['customs_fact']),
            'warehouse_plan': fmt(c['warehouse_plan']),
            'warehouse_fact': fmt(c['warehouse_fact']),
            'delays':         calc_delays(c),
            'fpz_url':        fpz_url,
            'ffz_url':        ffz_url,
        })

    data = {
        'updated_at': datetime.now().strftime('%d.%m.%Y %H:%M'),
        'containers': output,
    }

    json_path = os.path.join(REPO_PATH, 'data.json')
    with open(json_path, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

    print(f'data.json готов — {len(output)} контейнеров')

    # git add -> commit -> push
    subprocess.run(['git', '-C', REPO_PATH, 'add', '-A'], check=True)
    result = subprocess.run(
        ['git', '-C', REPO_PATH, 'commit', '-m',
         f'tracking update {datetime.now().strftime("%d.%m.%Y %H:%M")}'],
        capture_output=True, text=True
    )
    if 'nothing to commit' in result.stdout + result.stderr:
        print('Нет изменений для публикации')
    else:
        subprocess.run(['git', '-C', REPO_PATH, 'push'], check=True)
        print('Опубликовано на GitHub Pages')


if __name__ == '__main__':
    main()
