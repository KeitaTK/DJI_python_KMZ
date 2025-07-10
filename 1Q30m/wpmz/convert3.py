import re
from pathlib import Path

# --- オフセット設定 ---
REF_LAT = 36.0737551548803
REF_LON = 136.556923378546
TODAY_LAT = 36.073747
TODAY_LON = 136.556924

delta_lat = TODAY_LAT - REF_LAT
delta_lon = TODAY_LON - REF_LON

def shift_coordinates_in_text(text, delta_lon, delta_lat):
    """
    text から <coordinates> タグを探し、緯度経度をオフセットして返す。
    """
    def replace_coord(match):
        coord_line = match.group(1).strip()
        # 高度があってもなくても対応
        parts = coord_line.split(',')
        if len(parts) < 2:
            return match.group(0)  # 無効な座標は無視
        lon = float(parts[0])
        lat = float(parts[1])
        alt = float(parts[2]) if len(parts) > 2 else 0.0

        new_lon = lon + delta_lon
        new_lat = lat + delta_lat
        new_coord = f"{new_lon:.15f},{new_lat:.15f},{alt:.1f}"
        return f"<coordinates>\n            {new_coord}\n          </coordinates>"

    # coordinatesタグをすべて変換
    return re.sub(r"<coordinates>\s*([^<]+)\s*</coordinates>", replace_coord, text)

def process_kml_file(input_path, output_path):
    text = Path(input_path).read_text(encoding='utf-8')
    updated_text = shift_coordinates_in_text(text, delta_lon, delta_lat)
    Path(output_path).write_text(updated_text, encoding='utf-8')
    print(f"✅ 変換完了: {output_path}")

if __name__ == '__main__':
    input_kml = r"C:\Users\keita\Documents\local\M30_GPS\1Q30m\wpmz\template.kml"
    output_kml = r"C:\Users\keita\Documents\local\M30_GPS\1Q30m\wpmz\template_corrected.kml"
    process_kml_file(input_kml, output_kml)
