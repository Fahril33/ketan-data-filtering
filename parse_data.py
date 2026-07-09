import zipfile
import xml.etree.ElementTree as ET
import re
import json
from collections import Counter

def map_desa(val_str):
    if not val_str:
        return None
    val_lower = str(val_str).lower().strip()
    if 'bulili' in val_lower:
        return 'Bulili'
    if 'kadidia' in val_lower:
        return 'Kadidia'
    if 'kamarora b' in val_lower or 'kamarora_b' in val_lower:
        return 'Kamarora B'
    if 'kamarora a' in val_lower or 'kamarora_a' in val_lower or 'kamarora' in val_lower:
        return 'Kamarora A'
    if 'lemban' in val_lower or 'tongoa' in val_lower:
        return 'Lembantongoa'
    if 'sopu' in val_lower:
        return 'Sopu'
    if 'uwenuni' in val_lower:
        return 'Uwenuni'
    return None

def read_xlsx(file_path):
    with zipfile.ZipFile(file_path, 'r') as zip_ref:
        shared_strings = []
        try:
            ss_data = zip_ref.read('xl/sharedStrings.xml')
            root = ET.fromstring(ss_data)
            ns = {'ns': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
            for si in root.findall('.//ns:t', ns):
                shared_strings.append(si.text)
        except KeyError:
            pass
        
        sheet_data = zip_ref.read('xl/worksheets/sheet1.xml')
        root = ET.fromstring(sheet_data)
        ns = {'ns': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
        
        rows = {}
        for r in root.findall('.//ns:row', ns):
            row_idx = int(r.get('r'))
            rows[row_idx] = {}
            for c in r.findall('ns:c', ns):
                cell_ref = c.get('r')
                col_letter = re.match(r'^([A-Z]+)', cell_ref).group(1)
                val_el = c.find('ns:v', ns)
                val = val_el.text if val_el is not None else None
                t_attr = c.get('t')
                if t_attr == 's' and val is not None:
                    try:
                        val = shared_strings[int(val)]
                    except (IndexError, ValueError):
                        pass
                rows[row_idx][col_letter] = val
        
        return rows

def parse_months(detail_str, age_years):
    if not detail_str:
        if age_years is not None:
            return int(age_years * 12)
        return None
        
    detail_lower = str(detail_str).lower()
    months = 0
    years = 0
    
    m_yr_mth = re.search(r'(\d+)\s*tahun\s*(\d+)\s*bulan', detail_lower)
    if m_yr_mth:
        years = int(m_yr_mth.group(1))
        months = int(m_yr_mth.group(2))
        return years * 12 + months
        
    m_mth = re.search(r'(\d+)\s*bulan', detail_lower)
    if m_mth:
        m_yr = re.search(r'(\d+)\s*tahun', detail_lower)
        if m_yr:
            years = int(m_yr.group(1))
        months = int(m_mth.group(1))
        return years * 12 + months
        
    m_yr = re.search(r'(\d+)\s*tahun', detail_lower)
    if m_yr:
        years = int(m_yr.group(1))
        return years * 12
        
    m_day = re.search(r'(\d+)\s*hari', detail_lower)
    if m_day:
        days = int(m_day.group(1))
        return round(days / 30.0, 1)
        
    if age_years is not None:
        return int(age_years * 12)
        
    return None




def clean_name(name_str):
    if not name_str:
        return None
    name_str = str(name_str).strip()
    if re.match(r'^\d+$', name_str) or 'KR-' in name_str or len(name_str) > 25 or any(word in name_str.lower() for word in ['layak huni', 'sakit', 'sehat', 'revisi', 'kk', 'ktp', 'alamat', 'tidak ada', 'belum ada']):
        return None
    return name_str

def parse_database(file_path):
    rows = read_xlsx(file_path)
    surveyor_names = ['Agelia Magi', '7210031309120007', 'DEPRIN', 'Mirawati', 'HERLAMBANG P. PRATAMA', 'MARIA TALANTAN', 'Marta Tamolo', 'Melani', 'Aqila Ramadani']
    
    parsed_records = []
    
    for r_idx in range(2, len(rows)+1):
        if r_idx not in rows:
            continue
        row = rows[r_idx]
        
        if not any(row.values()):
            continue
            
        record = {
            'row_idx': r_idx,
            'timestamp': None,
            'dusun': 'Lainnya',
            'desa': 'Lainnya',
            'kecamatan': 'Lainnya',
            'village_id': None,
            'hh_id': None,
            'nama_kk': None,
            'phone': None,
            'asal_rt_rw': None,
            'nik_kk': None,
            'surveyor': 'Tidak Diketahui',
            'umur': None,
            'gender': 'Tidak Diketahui',
            'kategori_rentan': [],
            'detail_usia_penyakit': None,
            'needs_dropdown': [],
            'needs_specific': [],
            'notes': None,
            'nama_rentan': None,
            'umur_bulan': None
        }
        
        unassigned = {}
        for col, val in row.items():
            if val is not None and str(val).strip():
                unassigned[col] = str(val).strip()
                
        # 1. Timestamp
        for col in list(unassigned.keys()):
            val_str = unassigned[col]
            if (re.match(r'^\d+(\.\d+)?$', val_str) and 40000 < float(val_str) < 50000) or ('/' in val_str and ':' in val_str):
                record['timestamp'] = val_str
                del unassigned[col]
                break
                
        # 2. Desa and Kecamatan
        for col in list(unassigned.keys()):
            val_str = unassigned[col]
            mapped_d = map_desa(val_str)
            if mapped_d:
                record['desa'] = mapped_d
                del unassigned[col]
            elif val_str in ['Nokilalaki', 'Palolo']:
                record['kecamatan'] = val_str
                del unassigned[col]
                
        # 3. Surveyor
        for col in list(unassigned.keys()):
            val_str = unassigned[col]
            if val_str in surveyor_names:
                record['surveyor'] = val_str
                del unassigned[col]
                
        # 4. Gender
        for col in list(unassigned.keys()):
            val_str = unassigned[col]
            if val_str in ['Laki-laki', 'Perempuan']:
                record['gender'] = val_str
                del unassigned[col]
                
        # 5. KR IDs
        kr_ids = []
        for col in list(unassigned.keys()):
            val_str = unassigned[col]
            if val_str.startswith('KR-') and re.match(r'^KR-\d+$', val_str):
                kr_ids.append((col, val_str))
                
        for col, val_str in kr_ids:
            num = int(val_str.split('-')[1])
            if num in [1, 60, 90, 200, 417, 581, 661, 754, 839, 864, 976, 986]:
                record['village_id'] = val_str
            else:
                record['hh_id'] = val_str
            del unassigned[col]
            
        # 6. NIK/No KK (Long number >= 15 digits)
        for col in list(unassigned.keys()):
            val_str = unassigned[col]
            if re.match(r'^\d+$', val_str) and len(val_str) >= 15:
                record['nik_kk'] = val_str
                del unassigned[col]
                
        # 7. Phone number
        for col in list(unassigned.keys()):
            val_str = unassigned[col]
            if re.match(r'^08\d+$', val_str) or (re.match(r'^\d+$', val_str) and 9 <= len(val_str) <= 13) or ('E10' in val_str):
                record['phone'] = val_str
                del unassigned[col]
                
        # 8. Dusun (Usually in column C)
        if 'C' in unassigned:
            val_str = unassigned['C']
            if re.match(r'^\d+(\.0)?$', val_str):
                record['dusun'] = str(int(float(val_str)))
                del unassigned['C']
                
        # 9. Umur (Usually in column L)
        if 'L' in unassigned:
            val_str = unassigned['L']
            if re.match(r'^\d+(\.0)?$', val_str):
                record['umur'] = float(val_str)
                del unassigned['L']
                
        # 10. Kebutuhan Dropdown
        dropdown_options = ['Selimut, Sembako, Kelambu', 'Susu Formula, Popok, Obat obantan balita', 'Gangguan perilaku']
        for col in list(unassigned.keys()):
            val_str = unassigned[col]
            if val_str in dropdown_options:
                record['needs_dropdown'] = [n.strip() for n in val_str.split(',')]
                del unassigned[col]
                
        # 11. Kategori Rentan
        options = ['Bayi / Balita', 'Lanjut Usia (Lansia)', 'Ibu Hamil', 'Ibu Menyusui', 'Disabilitas (Fisik / Sensorik / Mental)', 'Penyakit Kronis']
        for col in list(unassigned.keys()):
            val_str = unassigned[col]
            if any(opt == val_str or val_str.startswith(opt) for opt in options):
                record['kategori_rentan'] = [c.strip() for c in val_str.split(',')]
                del unassigned[col]

        # Extract details and specific needs
        for col in list(unassigned.keys()):
            val_str = unassigned[col]
            if any(word in val_str.lower() for word in ['bulan', 'tahun', 'hari', 'kandungan', 'menyusui', 'penyakit', 'stroke', 'diabetes', 'asma', 'hipertensi', 'jantung', 'gula', 'katarak', 'gatal', 'syaraf', 'disabilitas', 'tuna', 'kolestrol', 'urat']):
                if any(need in val_str.lower() for need in ['susu', 'obat', 'popok', 'pempers', 'selimut', 'kelambu', 'sembako', 'terpal', 'kasur', 'bantuan', 'alat', 'mandi', 'telon', 'biskuit', 'kursi roda', 'tongkat']):
                    record['needs_specific'] = [n.strip() for n in re.split(r',|dan|&', val_str) if n.strip()]
                else:
                    record['detail_usia_penyakit'] = val_str
                del unassigned[col]
                
        # Classify the rest
        for col in list(unassigned.keys()):
            val_str = unassigned[col]
            if any(need in val_str.lower() for need in ['susu', 'obat', 'popok', 'pempers', 'selimut', 'kelambu', 'sembako', 'terpal', 'kasur', 'bantuan', 'alat', 'mandi', 'telon', 'biskuit', 'kursi roda', 'tongkat']):
                record['needs_specific'].extend([n.strip() for n in re.split(r',|dan|&', val_str) if n.strip()])
                del unassigned[col]
                continue
            if len(val_str) > 25:
                record['notes'] = val_str
                del unassigned[col]
                continue
            if any(w in val_str.lower() for w in ['rt', 'rw', 'dusun', 'asal', 'desa', 'dila', 'mando', 'bose']):
                record['asal_rt_rw'] = val_str
                del unassigned[col]
                continue

        # Extract names
        if 'F' in unassigned:
            cleaned = clean_name(unassigned['F'])
            if cleaned:
                record['nama_kk'] = cleaned
                del unassigned['F']
        if 'J' in unassigned:
            cleaned = clean_name(unassigned['J'])
            if cleaned:
                record['nama_rentan'] = cleaned
                del unassigned['J']

        for col in list(unassigned.keys()):
            val_str = unassigned[col]
            cleaned = clean_name(val_str)
            if cleaned:
                if col in ['G', 'H', 'I']:
                    if not record['nama_kk']:
                        record['nama_kk'] = cleaned
                        del unassigned[col]
                else:
                    if not record['nama_rentan']:
                        record['nama_rentan'] = cleaned
                        del unassigned[col]
                        
        for col in list(unassigned.keys()):
            cleaned = clean_name(unassigned[col])
            if cleaned:
                if not record['nama_rentan']:
                    record['nama_rentan'] = cleaned
                elif not record['nama_kk']:
                    record['nama_kk'] = cleaned
                    
        if not record['nama_kk']:
            record['nama_kk'] = "Tidak Diketahui"
        if not record['nama_rentan']:
            record['nama_rentan'] = record['nama_kk']
            
        record['umur_bulan'] = parse_months(record['detail_usia_penyakit'], record['umur'])
        
        # Determine categories dynamically if empty
        if not record['kategori_rentan']:
            if record['umur'] is not None:
                if record['umur'] >= 60:
                    record['kategori_rentan'].append('Lanjut Usia (Lansia)')
                elif record['umur_bulan'] is not None:
                    if record['umur_bulan'] < 12:
                        record['kategori_rentan'].append('Bayi')
                    elif record['umur_bulan'] <= 60:
                        record['kategori_rentan'].append('Balita')
            
            det = str(record['detail_usia_penyakit'] or '').lower()
            if 'hamil' in det or 'kandungan' in det:
                record['kategori_rentan'].append('Ibu Hamil')
            if 'menyusui' in det:
                record['kategori_rentan'].append('Ibu Menyusui')
            if 'disabilitas' in det:
                record['kategori_rentan'].append('Disabilitas')
            if 'kronis' in det or 'stroke' in det or 'diabetes' in det or 'asma' in det:
                record['kategori_rentan'].append('Penyakit Kronis')
                
        # Normalize categories to standard parent categories
        kategori_updated = []
        for kat in record['kategori_rentan']:
            kat_lower = str(kat).lower().strip()
            if 'lansia' in kat_lower or 'lanjut usia' in kat_lower:
                kategori_updated.append('Lanjut Usia (Lansia)')
            elif 'hamil' in kat_lower or 'bumil' in kat_lower:
                kategori_updated.append('Ibu Hamil')
            elif 'menyusui' in kat_lower or 'busui' in kat_lower:
                kategori_updated.append('Ibu Menyusui')
            elif 'disabilitas' in kat_lower:
                kategori_updated.append('Disabilitas')
            elif 'kronis' in kat_lower or 'penyakit' in kat_lower or any(ill in kat_lower for ill in ['stroke', 'diabetes', 'asma', 'jantung', 'hipertensi', 'tensi', 'gula', 'paru', 'bronkitis', 'tumor', 'komplikasi', 'saraf']):
                kategori_updated.append('Penyakit Kronis')
            elif 'bayi' in kat_lower or 'balita' in kat_lower:
                if record['umur_bulan'] is not None:
                    if record['umur_bulan'] < 12:
                        kategori_updated.append('Bayi (<12 Bulan)')
                    else:
                        kategori_updated.append('Balita (1-5 Tahun)')
                else:
                    if record['umur'] is not None and record['umur'] < 1:
                        kategori_updated.append('Bayi (<12 Bulan)')
                    else:
                        kategori_updated.append('Balita (1-5 Tahun)')
            else:
                kategori_updated.append(kat)
        
        # Re-verify based on age if age is provided
        if record['umur'] is not None:
            if record['umur'] >= 60 and 'Lanjut Usia (Lansia)' not in kategori_updated:
                kategori_updated.append('Lanjut Usia (Lansia)')
            elif record['umur_bulan'] is not None:
                if record['umur_bulan'] < 12 and 'Bayi (<12 Bulan)' not in kategori_updated:
                    kategori_updated.append('Bayi (<12 Bulan)')
                elif record['umur_bulan'] <= 60 and 'Balita (1-5 Tahun)' not in kategori_updated:
                    kategori_updated.append('Balita (1-5 Tahun)')
                    
        record['kategori_rentan'] = list(set(kategori_updated))
        
        if not record['kategori_rentan'] or record['kategori_rentan'] == ['Umum']:
            record['kategori_rentan'] = ['Umum']

        # Build all_needs: merge dropdown + specific, keep original text, title-case for consistency
        all_needs = []
        seen = set()
        for n in record['needs_dropdown'] + record['needs_specific']:
            cleaned = n.strip()
            if not cleaned:
                continue
            # Title-case for display consistency but preserve specificity
            display = cleaned[0].upper() + cleaned[1:] if len(cleaned) > 1 else cleaned.upper()
            key = display.lower()
            if key not in seen:
                seen.add(key)
                all_needs.append(display)
        record['all_needs'] = all_needs
            
        parsed_records.append(record)
        
    return parsed_records

# Main parse logic
db = parse_database("./Untitled spreadsheet (1).xlsx")
with open("./survey_data.json", "w") as f:
    json.dump(db, f, indent=2)
print("Written parsed data to survey_data.json successfully!")
