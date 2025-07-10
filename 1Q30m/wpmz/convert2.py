import xml.etree.ElementTree as ET

# --- 入力パラメータ ---
REF_LAT = 36.0737551548803    # 想定の基準離陸点（緯度）
REF_LON = 136.556923378546    # 想定の基準離陸点（経度）
TODAY_LAT = 36.073747         # 実際の今日のGPS（緯度）
TODAY_LON = 136.556924        # 実際の今日のGPS（経度）

# --- オフセット計算 ---
delta_lat = TODAY_LAT - REF_LAT
delta_lon = TODAY_LON - REF_LON

# --- KMLの名前空間定義 ---
NS = {
    'kml': 'http://www.opengis.net/kml/2.2',
    'wpml': 'http://www.dji.com/wpmz/1.0.6'
}

def shift_kml_coordinates(kml_path, output_path, delta_lat, delta_lon):
    """
    KMLファイルの座標をオフセットで補正する関数
    
    Args:
        kml_path (str): 入力KMLファイルパス
        output_path (str): 出力KMLファイルパス
        delta_lat (float): 緯度オフセット
        delta_lon (float): 経度オフセット
        
    Returns:
        int: 変更されたウェイポイント数
    """
    # XMLパーサーの設定
    parser = ET.XMLParser(encoding='utf-8')
    tree = ET.parse(kml_path, parser)
    root = tree.getroot()

    # 変更されたウェイポイントカウント
    changed_count = 0
    
    print("=== GPS座標補正処理開始 ===")
    print(f"緯度オフセット: {delta_lat:.10f}°")
    print(f"経度オフセット: {delta_lon:.10f}°")
    print()
    
    # 各Placemarkの座標を検索・更新
    for placemark in root.findall('.//kml:Placemark', NS):
        coord_elem = placemark.find('.//kml:Point/kml:coordinates', NS)
        if coord_elem is not None and coord_elem.text:
            # 座標文字列を分割（経度,緯度,高度の順）
            coord_text = coord_elem.text.strip()
            parts = coord_text.split(',')
            
            if len(parts) >= 2:
                lon, lat = float(parts[0]), float(parts[1])
                # 高度がある場合は保持
                alt = float(parts[2]) if len(parts) >= 3 else None

                # 元の座標を表示
                print(f"ウェイポイント {changed_count + 1}:")
                print(f"  元の座標: {lon:.12f}, {lat:.12f}")
                
                # 緯度経度を補正
                new_lon = lon + delta_lon
                new_lat = lat + delta_lat
                
                # 新しい座標を表示
                print(f"  新しい座標: {new_lon:.12f}, {new_lat:.12f}")
                
                # 元の形式を保持して座標を更新
                if alt is not None:
                    coord_elem.text = f"{new_lon:.12f},{new_lat:.12f},{alt:.10f}"
                else:
                    coord_elem.text = f"{new_lon:.12f},{new_lat:.12f}"
                    
                changed_count += 1
                print(f"  → 更新完了")
                print()

    # 保存（UTF-8エンコーディング、XML宣言付き）
    tree.write(output_path, encoding='utf-8', xml_declaration=True, method='xml')
    
    print(f'✅ GPS座標補正完了: {changed_count} 個のウェイポイントを更新')
    print(f'✅ 保存先: {output_path}')
    
    return changed_count

# メイン実行部分
if __name__ == '__main__':
    # ファイルパスの設定
    input_kml = r"C:\Users\keita\Documents\local\M30_GPS\1Q30m\wpmz\template.kml"
    output_kml = r"C:\Users\keita\Documents\local\M30_GPS\1Q30m\wpmz\template_corrected.kml"
    
    print("=== DJI KMLファイル GPS座標補正プログラム ===")
    print(f"入力ファイル: {input_kml}")
    print(f"出力ファイル: {output_kml}")
    print()
    
    # 座標補正実行
    try:
        changed = shift_kml_coordinates(input_kml, output_kml, delta_lat, delta_lon)
        print(f"\n=== 処理完了 ===")
        print(f"変更されたウェイポイント数: {changed}")
        print("補正されたKMLファイルをDJI Pilot 2で読み込み可能です。")
    except FileNotFoundError:
        print(f"❌ エラー: 入力ファイルが見つかりません: {input_kml}")
        print("ファイルパスを確認してください。")
    except Exception as e:
        print(f"❌ エラーが発生しました: {e}")
