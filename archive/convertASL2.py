import xml.etree.ElementTree as ET

# 離陸地点のASL
TAKEOFF_ASL_M = 612.0

# 名前空間
NS = {
    'kml': 'http://www.opengis.net/kml/2.2',
    'wpml': 'http://www.dji.com/wpmz/1.0.6'
}
ET.register_namespace('', NS['kml'])
ET.register_namespace('wpml', NS['wpml'])

def convert_to_asl_gps_compatible(input_path, output_path, takeoff_asl):
    """
    GPS使用時でもASL高度を認識させるための変換プログラム
    """
    tree = ET.parse(input_path)
    root = tree.getroot()

    print("=== GPS使用時 ASL高度変換処理 ===")
    print(f"離陸地点ASL: {takeoff_asl:.1f}m")
    print()
    
    # 1. heightMode を absolute に設定
    height_mode_elem = root.find('.//wpml:waylineCoordinateSysParam/wpml:heightMode', NS)
    if height_mode_elem is not None:
        height_mode_elem.text = 'absolute'
        print("✅ heightMode を absolute に設定")
    
    # 2. positioningType を GPS に設定（RTKではなく）
    positioning_elem = root.find('.//wpml:waylineCoordinateSysParam/wpml:positioningType', NS)
    if positioning_elem is not None:
        positioning_elem.text = 'GPS'
        print("✅ positioningType を GPS に設定")
    
    # 3. coordinateMode を WGS84 に確認・設定
    coord_mode_elem = root.find('.//wpml:waylineCoordinateSysParam/wpml:coordinateMode', NS)
    if coord_mode_elem is not None:
        coord_mode_elem.text = 'WGS84'
        print("✅ coordinateMode を WGS84 に設定")
    
    # 4. takeOffSecurityHeight を0に設定（重要）
    takeoff_height_elem = root.find('.//wpml:missionConfig/wpml:takeOffSecurityHeight', NS)
    if takeoff_height_elem is not None:
        takeoff_height_elem.text = '0'
        print("✅ takeOffSecurityHeight を 0 に設定")
    
    # 5. globalHeight を適切なASL値に設定
    global_height_elem = root.find('.//wpml:globalHeight', NS)
    if global_height_elem is not None:
        # 最小ASL高度を設定（離陸地点+安全マージン）
        min_asl = takeoff_asl + 30  # 離陸地点+30m
        global_height_elem.text = f"{min_asl:.1f}"
        print(f"✅ globalHeight を {min_asl:.1f}m (ASL) に設定")
    
    # 6. caliFlightEnable を 0 に設定
    cali_elem = root.find('.//wpml:caliFlightEnable', NS)
    if cali_elem is not None:
        cali_elem.text = '0'
        print("✅ caliFlightEnable を 0 に設定")
    
    # 7. finishAction を goHome に設定
    finish_action_elem = root.find('.//wpml:missionConfig/wpml:finishAction', NS)
    if finish_action_elem is not None:
        finish_action_elem.text = 'goHome'
        print("✅ finishAction を goHome に設定")
    
    # 8. 各ウェイポイントの高度を変換
    updated_count = 0
    print("\n=== ウェイポイント高度変換 ===")
    
    for i, placemark in enumerate(root.findall('.//kml:Placemark', NS)):
        print(f"\nウェイポイント {i + 1}:")
        
        # wpml:height を ASL に変換
        height_elem = placemark.find('.//wpml:height', NS)
        if height_elem is not None and height_elem.text:
            try:
                rel_alt = float(height_elem.text.strip())
                abs_alt = rel_alt + takeoff_asl
                height_elem.text = f"{abs_alt:.6f}"
                print(f"  height: {rel_alt:.1f}m (相対) → {abs_alt:.1f}m (ASL)")
            except ValueError:
                print(f"  ⚠️ 無効な height 値: {height_elem.text}")
        
        # wpml:ellipsoidHeight を ASL に変換
        ellipsoid_elem = placemark.find('.//wpml:ellipsoidHeight', NS)
        if ellipsoid_elem is not None and ellipsoid_elem.text:
            try:
                rel_alt = float(ellipsoid_elem.text.strip())
                abs_alt = rel_alt + takeoff_asl
                ellipsoid_elem.text = f"{abs_alt:.6f}"
                print(f"  ellipsoidHeight: {rel_alt:.1f}m (相対) → {abs_alt:.1f}m (ASL)")
            except ValueError:
                print(f"  ⚠️ 無効な ellipsoidHeight 値: {ellipsoid_elem.text}")
        
        # coordinates にも高度を追加・更新
        coord_elem = placemark.find('.//kml:Point/kml:coordinates', NS)
        if coord_elem is not None and coord_elem.text:
            parts = coord_elem.text.strip().split(',')
            if len(parts) >= 2:
                lon, lat = float(parts[0]), float(parts[1])
                
                # wpml:height から ASL 高度を取得
                if height_elem is not None and height_elem.text:
                    try:
                        alt = float(height_elem.text.strip())
                        coord_elem.text = f"{lon:.6f},{lat:.6f},{alt:.6f}"
                        print(f"  coordinates: {alt:.1f}m (ASL) に更新")
                        updated_count += 1
                    except ValueError:
                        print(f"  ⚠️ coordinates 更新エラー")
    
    # 9. 追加の GPS 互換性設定
    # gimbalPitchMode を適切に設定
    gimbal_pitch_elem = root.find('.//wpml:gimbalPitchMode', NS)
    if gimbal_pitch_elem is not None:
        gimbal_pitch_elem.text = 'manual'
        print("✅ gimbalPitchMode を manual に設定")
    
    # globalWaypointTurnMode を適切に設定
    turn_mode_elem = root.find('.//wpml:globalWaypointTurnMode', NS)
    if turn_mode_elem is not None:
        turn_mode_elem.text = 'toPointAndStopWithDiscontinuityCurvature'
        print("✅ globalWaypointTurnMode を適切に設定")
    
    # 10. ファイル保存
    tree.write(output_path, encoding='utf-8', xml_declaration=True)
    
    print(f"\n=== 変換完了 ===")
    print(f"✅ 処理済みウェイポイント数: {updated_count}")
    print(f"✅ 保存先: {output_path}")
    
    # 11. 使用方法の案内
    print("\n📋 DJI Pilot 2 でのインポート手順:")
    print("  1. KMLファイルをDJI Pilot 2にインポート")
    print("  2. ミッション設定で高度モードを確認")
    print("  3. 「絶対高度(ASL)」が選択されていることを確認")
    print("  4. GPS測位モードが有効になっていることを確認")
    print("  5. 離陸前にホームポイント高度を確認")
    
    print("\n⚠️  注意事項:")
    print("  - GPS精度は±3〜5mです")
    print("  - 気圧変化により高度誤差が発生する場合があります")
    print("  - 重要なミッションではRTK使用を推奨します")
    
    return updated_count

# 座標補正機能付きバージョン
def convert_with_gps_correction(input_path, output_path, takeoff_asl, 
                              ref_lat=None, ref_lon=None, 
                              today_lat=None, today_lon=None):
    """
    GPS補正機能付きのASL変換プログラム
    """
    # まずASL変換を実行
    convert_to_asl_gps_compatible(input_path, output_path, takeoff_asl)
    
    # GPS補正が指定されている場合は追加で実行
    if all([ref_lat, ref_lon, today_lat, today_lon]):
        print("\n=== GPS座標補正処理 ===")
        
        # オフセット計算
        delta_lat = today_lat - ref_lat
        delta_lon = today_lon - ref_lon
        
        print(f"緯度オフセット: {delta_lat:.10f}°")
        print(f"経度オフセット: {delta_lon:.10f}°")
        
        # 座標補正を実行
        tree = ET.parse(output_path)
        root = tree.getroot()
        
        corrected_count = 0
        for placemark in root.findall('.//kml:Placemark', NS):
            coord_elem = placemark.find('.//kml:Point/kml:coordinates', NS)
            if coord_elem is not None and coord_elem.text:
                parts = coord_elem.text.strip().split(',')
                if len(parts) >= 3:
                    lon, lat, alt = float(parts[0]), float(parts[1]), float(parts[2])
                    
                    # 座標補正
                    new_lon = lon + delta_lon
                    new_lat = lat + delta_lat
                    
                    coord_elem.text = f"{new_lon:.6f},{new_lat:.6f},{alt:.6f}"
                    corrected_count += 1
        
        # 更新されたファイルを保存
        tree.write(output_path, encoding='utf-8', xml_declaration=True)
        print(f"✅ GPS座標補正完了: {corrected_count} 個のウェイポイント")

if __name__ == '__main__':
    # ファイルパスの設定
    input_kml = r"C:\Users\keita\Documents\local\M30_GPS\1Q30m\wpmz\template.kml"
    output_kml = r"C:\Users\keita\Documents\local\M30_GPS\1Q30m\wpmz\template_GPS_ASL.kml"
    
    # 基本的なASL変換
    print("=== GPS使用時 ASL変換プログラム ===")
    convert_to_asl_gps_compatible(input_kml, output_kml, TAKEOFF_ASL_M)
    
    # GPS補正も必要な場合は以下のコメントを外して使用
    # convert_with_gps_correction(
    #     input_kml, output_kml, TAKEOFF_ASL_M,
    #     ref_lat=36.0737551548803,    # 想定の基準離陸点
    #     ref_lon=136.556923378546,
    #     today_lat=36.073747,         # 実際の今日のGPS
    #     today_lon=136.556924
    # )
