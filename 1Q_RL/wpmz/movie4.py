from lxml import etree

# --- 設定パラメータ（用途コメント付き） ---
BASE_HEIGHT = 613.5            # 高度変換時の基準高さ[m]
NEW_HEIGHT_MODE = 'als'        # 高度基準（'als':絶対高度基準, 'relativeToStartPoint'等）
AUTO_FLIGHT_SPEED = 15         # グローバル自動飛行速度[m/s]
TAKEOFF_SECURITY_HEIGHT = 20   # 安全離陸高度[m]
TURN_MODE = 'toPointAndStopWithDiscontinuityCurvature'  # 旋回モード
HOVER_SECONDS = 2              # 各WPでのホバリング秒数
GIMBAL_PITCH = -90             # ジンバルピッチ角度[°]（真下）
YAW_ANGLE = 87.37              # 機体ヨー角度[°]（常時固定）
VIDEO_TYPE = 'zoom'            # 動画種別（'zoom', 'wide', 'ir'）
ZOOM_FACTOR = 5                # ズーム倍率（zoom時のみ有効）
ENABLE_ZOOM = 1                # ズームカメラを使うか（0:使わない, 1:使う）
ENABLE_WIDE = 1                # ワイドカメラを使うか（0:使わない, 1:使う）
ENABLE_IR = 1                  # IRカメラを使うか（0:使わない, 1:使う）
PAYLOAD_POSITION_INDEX = '0'   # ペイロード位置インデックス

# 入出力ファイルパス
INPUT_KML = r"C:\Users\keita\Documents\local\M30_GPS\1Q_RL\wpmz\template.kml"
OUTPUT_KML = r"C:\Users\keita\Documents\local\M30_GPS\1Q_RL\wpmz\template_movie.kml"

def convert_kml(input_kml, output_kml):
    ns = {
        'kml': 'http://www.opengis.net/kml/2.2',
        'wpml': 'http://www.dji.com/wpmz/1.0.6'
    }

    tree = etree.parse(input_kml)
    root = tree.getroot()
    folder = root.find('.//kml:Folder', ns)

    # 高度基準の変更（既にalsならスキップ）
    height_mode_elem = root.find('.//wpml:waylineCoordinateSysParam/wpml:heightMode', ns)
    if height_mode_elem is not None:
        if height_mode_elem.text != NEW_HEIGHT_MODE:
            height_mode_elem.text = NEW_HEIGHT_MODE
            # 各WPの高度値を変換
            for placemark in root.findall('.//kml:Placemark', ns):
                for tag in ['wpml:height', 'wpml:ellipsoidHeight']:
                    h_elem = placemark.find(tag, ns)
                    if h_elem is not None:
                        try:
                            old_height = float(h_elem.text)
                            h_elem.text = str(old_height + BASE_HEIGHT)
                        except Exception:
                            pass

    # 自動飛行速度の変更（グローバル・各WPとも）
    auto_speed_elem = folder.find('wpml:autoFlightSpeed', ns)
    if auto_speed_elem is not None:
        auto_speed_elem.text = str(AUTO_FLIGHT_SPEED)
    for placemark in root.findall('.//kml:Placemark', ns):
        speed_elem = placemark.find('wpml:waypointSpeed', ns)
        if speed_elem is not None:
            speed_elem.text = str(AUTO_FLIGHT_SPEED)

    # 安全離陸高度
    takeoff_height_elem = root.find('.//wpml:takeOffSecurityHeight', ns)
    if takeoff_height_elem is not None:
        takeoff_height_elem.text = str(TAKEOFF_SECURITY_HEIGHT)

    # 旋回モード
    turn_mode_elem = folder.find('wpml:globalWaypointTurnMode', ns)
    if turn_mode_elem is not None:
        turn_mode_elem.text = TURN_MODE

    # ヘディングモード: followWayline かつ常に87.37°
    global_heading = folder.find('wpml:globalWaypointHeadingParam', ns)
    if global_heading is not None:
        mode_elem = global_heading.find('wpml:waypointHeadingMode', ns)
        angle_elem = global_heading.find('wpml:waypointHeadingAngle', ns)
        if mode_elem is not None:
            mode_elem.text = 'followWayline'
        if angle_elem is not None:
            angle_elem.text = str(YAW_ANGLE)
    for placemark in root.findall('.//kml:Placemark', ns):
        heading_param = placemark.find('wpml:waypointHeadingParam', ns)
        if heading_param is not None:
            mode_elem = heading_param.find('wpml:waypointHeadingMode', ns)
            angle_elem = heading_param.find('wpml:waypointHeadingAngle', ns)
            if mode_elem is not None:
                mode_elem.text = 'followWayline'
            if angle_elem is not None:
                angle_elem.text = str(YAW_ANGLE)
        else:
            heading_param = etree.SubElement(placemark, '{http://www.dji.com/wpmz/1.0.6}waypointHeadingParam')
            etree.SubElement(heading_param, '{http://www.dji.com/wpmz/1.0.6}waypointHeadingMode').text = 'followWayline'
            etree.SubElement(heading_param, '{http://www.dji.com/wpmz/1.0.6}waypointHeadingAngle').text = str(YAW_ANGLE)
            etree.SubElement(heading_param, '{http://www.dji.com/wpmz/1.0.6}waypointPoiPoint').text = '0.000000,0.000000,0.000000'
            etree.SubElement(heading_param, '{http://www.dji.com/wpmz/1.0.6}waypointHeadingPoiIndex').text = '0'

    placemarks = root.findall('.//kml:Placemark', ns)
    for i, placemark in enumerate(placemarks):
        action_group = placemark.find('.//wpml:actionGroup', ns)

        # actionGroup がなければ作成
        if action_group is None:
            action_group = etree.SubElement(placemark, '{http://www.dji.com/wpmz/1.0.6}actionGroup')
            etree.SubElement(action_group, '{http://www.dji.com/wpmz/1.0.6}actionGroupId').text = str(i)
            etree.SubElement(action_group, '{http://www.dji.com/wpmz/1.0.6}actionGroupStartIndex').text = str(i)
            etree.SubElement(action_group, '{http://www.dji.com/wpmz/1.0.6}actionGroupEndIndex').text = str(i)
            etree.SubElement(action_group, '{http://www.dji.com/wpmz/1.0.6}actionGroupMode').text = 'sequence'
            trigger = etree.SubElement(action_group, '{http://www.dji.com/wpmz/1.0.6}actionTrigger')
            etree.SubElement(trigger, '{http://www.dji.com/wpmz/1.0.6}actionTriggerType').text = 'reachPoint'

        actions = action_group.findall('wpml:action', ns)
        new_actions = []

        # rotateYawアクションを削除
        for action in actions:
            func = action.find('wpml:actionActuatorFunc', ns)
            if func is not None and func.text == 'rotateYaw':
                continue  # 削除
            new_actions.append(action)

        # 動画開始（最初のみ、各カメラ種別ごとに追加）
        if i == 0:
            if ENABLE_ZOOM:
                video_start = etree.Element('{http://www.dji.com/wpmz/1.0.6}action')
                etree.SubElement(video_start, '{http://www.dji.com/wpmz/1.0.6}actionId').text = '998'
                etree.SubElement(video_start, '{http://www.dji.com/wpmz/1.0.6}actionActuatorFunc').text = 'startRecordVideo'
                param = etree.SubElement(video_start, '{http://www.dji.com/wpmz/1.0.6}actionActuatorFuncParam')
                etree.SubElement(param, '{http://www.dji.com/wpmz/1.0.6}payloadPositionIndex').text = PAYLOAD_POSITION_INDEX
                etree.SubElement(param, '{http://www.dji.com/wpmz/1.0.6}videoType').text = 'zoom'
                etree.SubElement(param, '{http://www.dji.com/wpmz/1.0.6}zoomFactor').text = str(ZOOM_FACTOR)
                new_actions.append(video_start)
            if ENABLE_WIDE:
                video_start = etree.Element('{http://www.dji.com/wpmz/1.0.6}action')
                etree.SubElement(video_start, '{http://www.dji.com/wpmz/1.0.6}actionId').text = '9981'
                etree.SubElement(video_start, '{http://www.dji.com/wpmz/1.0.6}actionActuatorFunc').text = 'startRecordVideo'
                param = etree.SubElement(video_start, '{http://www.dji.com/wpmz/1.0.6}actionActuatorFuncParam')
                etree.SubElement(param, '{http://www.dji.com/wpmz/1.0.6}payloadPositionIndex').text = PAYLOAD_POSITION_INDEX
                etree.SubElement(param, '{http://www.dji.com/wpmz/1.0.6}videoType').text = 'wide'
                new_actions.append(video_start)
            if ENABLE_IR:
                video_start = etree.Element('{http://www.dji.com/wpmz/1.0.6}action')
                etree.SubElement(video_start, '{http://www.dji.com/wpmz/1.0.6}actionId').text = '9982'
                etree.SubElement(video_start, '{http://www.dji.com/wpmz/1.0.6}actionActuatorFunc').text = 'startRecordVideo'
                param = etree.SubElement(video_start, '{http://www.dji.com/wpmz/1.0.6}actionActuatorFuncParam')
                etree.SubElement(param, '{http://www.dji.com/wpmz/1.0.6}payloadPositionIndex').text = PAYLOAD_POSITION_INDEX
                etree.SubElement(param, '{http://www.dji.com/wpmz/1.0.6}videoType').text = 'ir'
                new_actions.append(video_start)

        # ジンバルを動かしてからホバリング
        gimbal_action = etree.Element('{http://www.dji.com/wpmz/1.0.6}action')
        etree.SubElement(gimbal_action, '{http://www.dji.com/wpmz/1.0.6}actionId').text = '1001'
        etree.SubElement(gimbal_action, '{http://www.dji.com/wpmz/1.0.6}actionActuatorFunc').text = 'gimbalRotate'
        param = etree.SubElement(gimbal_action, '{http://www.dji.com/wpmz/1.0.6}actionActuatorFuncParam')
        etree.SubElement(param, '{http://www.dji.com/wpmz/1.0.6}gimbalRotateMode').text = 'absoluteAngle'
        etree.SubElement(param, '{http://www.dji.com/wpmz/1.0.6}gimbalPitchRotateEnable').text = '1'
        etree.SubElement(param, '{http://www.dji.com/wpmz/1.0.6}gimbalPitchRotateAngle').text = str(GIMBAL_PITCH)
        etree.SubElement(param, '{http://www.dji.com/wpmz/1.0.6}gimbalRollRotateEnable').text = '0'
        etree.SubElement(param, '{http://www.dji.com/wpmz/1.0.6}gimbalRollRotateAngle').text = '0'
        etree.SubElement(param, '{http://www.dji.com/wpmz/1.0.6}gimbalYawRotateEnable').text = '0'
        etree.SubElement(param, '{http://www.dji.com/wpmz/1.0.6}gimbalYawRotateAngle').text = '0'
        etree.SubElement(param, '{http://www.dji.com/wpmz/1.0.6}gimbalRotateTimeEnable').text = '0'
        etree.SubElement(param, '{http://www.dji.com/wpmz/1.0.6}gimbalRotateTime').text = '0'
        etree.SubElement(param, '{http://www.dji.com/wpmz/1.0.6}payloadPositionIndex').text = PAYLOAD_POSITION_INDEX
        new_actions.append(gimbal_action)

        # ホバリング追加（各ポイント共通）
        hover_action = etree.Element('{http://www.dji.com/wpmz/1.0.6}action')
        etree.SubElement(hover_action, '{http://www.dji.com/wpmz/1.0.6}actionId').text = '999'
        etree.SubElement(hover_action, '{http://www.dji.com/wpmz/1.0.6}actionActuatorFunc').text = 'stayForSeconds'
        param = etree.SubElement(hover_action, '{http://www.dji.com/wpmz/1.0.6}actionActuatorFuncParam')
        etree.SubElement(param, '{http://www.dji.com/wpmz/1.0.6}stayTime').text = str(HOVER_SECONDS)
        etree.SubElement(param, '{http://www.dji.com/wpmz/1.0.6}gimbalPitchRotateAngle').text = str(GIMBAL_PITCH)
        etree.SubElement(param, '{http://www.dji.com/wpmz/1.0.6}aircraftHeading').text = str(YAW_ANGLE)
        new_actions.append(hover_action)

        # 動画停止（最後のみ、各カメラ種別ごとに追加）
        if i == len(placemarks) - 1:
            if ENABLE_ZOOM:
                video_stop = etree.Element('{http://www.dji.com/wpmz/1.0.6}action')
                etree.SubElement(video_stop, '{http://www.dji.com/wpmz/1.0.6}actionId').text = '997'
                etree.SubElement(video_stop, '{http://www.dji.com/wpmz/1.0.6}actionActuatorFunc').text = 'stopRecordVideo'
                param = etree.SubElement(video_stop, '{http://www.dji.com/wpmz/1.0.6}actionActuatorFuncParam')
                etree.SubElement(param, '{http://www.dji.com/wpmz/1.0.6}payloadPositionIndex').text = PAYLOAD_POSITION_INDEX
                etree.SubElement(param, '{http://www.dji.com/wpmz/1.0.6}videoType').text = 'zoom'
                etree.SubElement(param, '{http://www.dji.com/wpmz/1.0.6}zoomFactor').text = str(ZOOM_FACTOR)
                new_actions.append(video_stop)
            if ENABLE_WIDE:
                video_stop = etree.Element('{http://www.dji.com/wpmz/1.0.6}action')
                etree.SubElement(video_stop, '{http://www.dji.com/wpmz/1.0.6}actionId').text = '9971'
                etree.SubElement(video_stop, '{http://www.dji.com/wpmz/1.0.6}actionActuatorFunc').text = 'stopRecordVideo'
                param = etree.SubElement(video_stop, '{http://www.dji.com/wpmz/1.0.6}actionActuatorFuncParam')
                etree.SubElement(param, '{http://www.dji.com/wpmz/1.0.6}payloadPositionIndex').text = PAYLOAD_POSITION_INDEX
                etree.SubElement(param, '{http://www.dji.com/wpmz/1.0.6}videoType').text = 'wide'
                new_actions.append(video_stop)
            if ENABLE_IR:
                video_stop = etree.Element('{http://www.dji.com/wpmz/1.0.6}action')
                etree.SubElement(video_stop, '{http://www.dji.com/wpmz/1.0.6}actionId').text = '9972'
                etree.SubElement(video_stop, '{http://www.dji.com/wpmz/1.0.6}actionActuatorFunc').text = 'stopRecordVideo'
                param = etree.SubElement(video_stop, '{http://www.dji.com/wpmz/1.0.6}actionActuatorFuncParam')
                etree.SubElement(param, '{http://www.dji.com/wpmz/1.0.6}payloadPositionIndex').text = PAYLOAD_POSITION_INDEX
                etree.SubElement(param, '{http://www.dji.com/wpmz/1.0.6}videoType').text = 'ir'
                new_actions.append(video_stop)

        # 古いアクションを削除して新しいものに置換
        for old_action in actions:
            action_group.remove(old_action)
        for new_action in new_actions:
            action_group.append(new_action)

    # 出力
    tree.write(output_kml, encoding='utf-8', pretty_print=True, xml_declaration=True)
    print(f"[✔] 処理完了: {input_kml} → {output_kml}")

# --- 実行 ---
if __name__ == '__main__':
    convert_kml(INPUT_KML, OUTPUT_KML)
