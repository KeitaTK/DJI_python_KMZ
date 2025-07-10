import xml.etree.ElementTree as ET

# --- 離陸地点の海抜高度 (ASL) を定義 ---
TAKEOFF_ASL_M = 612.0  # メートル単位

# --- KML名前空間の登録 ---
NS = {
    'kml': 'http://www.opengis.net/kml/2.2',
    'wpml': 'http://www.dji.com/wpmz/1.0.6'
}
ET.register_namespace('', NS['kml'])
ET.register_namespace('wpml', NS['wpml'])

def convert_relative_to_asl(kml_input, kml_output, takeoff_asl_m):
    tree = ET.parse(kml_input)
    root = tree.getroot()

    # 高度基準のモードを "absolute" に変更
    height_mode_elem = root.find('.//wpml:waylineCoordinateSysParam/wpml:heightMode', NS)
    if height_mode_elem is not None:
        height_mode_elem.text = 'absolute'  # 海抜基準に変更

    # Placemarkごとの相対高度をASLに変換
    for placemark in root.findall('.//kml:Placemark', NS):
        for tag in ['height', 'ellipsoidHeight']:
            elem = placemark.find(f'.//wpml:{tag}', NS)
            if elem is not None and elem.text:
                try:
                    rel_alt = float(elem.text.strip())
                    abs_alt = rel_alt + takeoff_asl_m
                    elem.text = f"{abs_alt:.6f}"
                except ValueError:
                    print(f"⚠️ 無効な高度値: {elem.text}")

    # 保存
    tree.write(kml_output, encoding='utf-8', xml_declaration=True)
    print(f"✅ ASL変換完了: {kml_output}")

if __name__ == '__main__':
    input_kml_path = r"C:\Users\keita\Documents\local\M30_GPS\1Q30m\wpmz\template.kml"
    output_kml_path = r"C:\Users\keita\Documents\local\M30_GPS\1Q30m\wpmz\template_ASL.kml"
    convert_relative_to_asl(input_kml_path, output_kml_path, TAKEOFF_ASL_M)
