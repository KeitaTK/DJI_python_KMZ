import xml.etree.ElementTree as ET
import csv

# 離陸地点のASL（メートル）
TAKEOFF_ASL_M = 613.5  # 適宜変更すること

# 名前空間
NS = {
    'kml': 'http://www.opengis.net/kml/2.2',
    'wpml': 'http://www.dji.com/wpmz/1.0.6'
}
ET.register_namespace('', NS['kml'])
ET.register_namespace('wpml', NS['wpml'])

def extract_altitudes(kml_path, csv_path):
    tree = ET.parse(kml_path)
    root = tree.getroot()

    output_rows = []
    output_rows.append(['Waypoint Index', 'ALT (relative)', 'ASL (absolute)'])

    for placemark in root.findall('.//kml:Placemark', NS):
        # index
        idx_elem = placemark.find('.//wpml:index', NS)
        index = idx_elem.text.strip() if idx_elem is not None else ''

        # 相対高度を取得（wpml:height）
        height_elem = placemark.find('.//wpml:height', NS)
        if height_elem is not None:
            try:
                alt_rel = float(height_elem.text.strip())
                alt_abs = alt_rel + TAKEOFF_ASL_M
                output_rows.append([index, f'{alt_rel:.2f}', f'{alt_abs:.2f}'])
            except:
                continue

    # CSVに書き出し
    with open(csv_path, 'w', newline='', encoding='utf-8') as f:
        writer = csv.writer(f)
        writer.writerows(output_rows)

    print(f'✅ CSV出力完了: {csv_path}')

if __name__ == '__main__':
    input_kml = r"C:\Users\keita\Documents\local\M30_GPS\1Q_RL\wpmz\template.kml"
    output_csv = r"C:\Users\keita\Documents\local\M30_GPS\1Q_RL\wpmz\waypoints_alt_asl.csv"
    extract_altitudes(input_kml, output_csv)
