import xml.etree.ElementTree as ET

def extract_coordinates(kml_file):
    # KML namespace definitions
    ns = {
        'kml': 'http://www.opengis.net/kml/2.2',
        'wpml': 'http://www.dji.com/wpmz/1.0.6'
    }

    # パース
    tree = ET.parse(kml_file)
    root = tree.getroot()

    # Placemark 要素を検索し、Point/coordinates を取得
    coords = []
    for placemark in root.findall('.//kml:Placemark', ns):
        pt = placemark.find('.//kml:Point/kml:coordinates', ns)
        if pt is not None and pt.text:
            coords.append(pt.text.strip())

    return coords

if __name__ == '__main__':
    kml_path = r"C:\Users\keita\Documents\local\M30_GPS\1Q30m\wpmz\template.kml"
    coordinates = extract_coordinates(kml_path)
    for idx, coord in enumerate(coordinates):
        print(f'Waypoint {idx}: {coord}')
