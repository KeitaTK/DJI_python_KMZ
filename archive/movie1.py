import xml.etree.ElementTree as ET

# --- 設定 ---
INPUT_WPML = r"C:\Users\keita\Documents\local\M30_GPS\1Q_RL\wpmz\waylines.wpml"
OUTPUT_WPML = r"C:\Users\keita\Documents\local\M30_GPS\1Q_RL\wpmz\waylines_modified.wpml"
HOVER_SECONDS = 10

# WPML 名前空間
NS = {
    'wpml': 'http://www.dji.com/wpmz/1.0.6'
}
ET.register_namespace('wpml', NS['wpml'])

def modify_wpml(input_path, output_path, hover_time):
    tree = ET.parse(input_path)
    root = tree.getroot()

    placemarks = root.findall('.//wpml:Placemark', NS)

    for idx, placemark in enumerate(placemarks):
        action_groups = placemark.findall('wpml:actionGroup', NS)
        for group in action_groups:
            actions = group.findall('wpml:action', NS)
            to_remove = []
            for action in actions:
                func = action.find('wpml:actionActuatorFunc', NS)
                if func is not None and func.text == 'orientedShoot':
                    to_remove.append(action)
            # 削除
            for a in to_remove:
                group.remove(a)

            # 追加: stay (hover)
            new_action = ET.Element(f"{{{NS['wpml']}}}action")
            ET.SubElement(new_action, f"{{{NS['wpml']}}}actionId").text = '999'
            ET.SubElement(new_action, f"{{{NS['wpml']}}}actionActuatorFunc").text = 'stay'
            param = ET.SubElement(new_action, f"{{{NS['wpml']}}}actionActuatorFuncParam")
            ET.SubElement(param, f"{{{NS['wpml']}}}stayTime").text = str(hover_time)
            group.append(new_action)

        # 最初・最後で録画アクション
        if idx == 0 or idx == len(placemarks) - 1:
            group = placemark.find('wpml:actionGroup', NS)
            if group is not None:
                new_action = ET.Element(f"{{{NS['wpml']}}}action")
                ET.SubElement(new_action, f"{{{NS['wpml']}}}actionId").text = '998' if idx == 0 else '997'
                ET.SubElement(new_action, f"{{{NS['wpml']}}}actionActuatorFunc").text = 'recordVideo'
                param = ET.SubElement(new_action, f"{{{NS['wpml']}}}actionActuatorFuncParam")
                ET.SubElement(param, f"{{{NS['wpml']}}}recordVideoOperationType").text = 'start' if idx == 0 else 'stop'
                group.append(new_action)

    tree.write(output_path, encoding='utf-8', xml_declaration=True)
    print(f'✅ WPMLファイル変換完了: {output_path}')

# 実行（コメントアウトを外すと実行されます）
modify_wpml(INPUT_WPML, OUTPUT_WPML, HOVER_SECONDS)
