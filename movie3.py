from lxml import etree

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

        # 写真アクションを削除
        for action in actions:
            func = action.find('wpml:actionActuatorFunc', ns)
            if func is not None and func.text == 'orientedShoot':
                continue  # 削除
            new_actions.append(action)

        # 動画開始（最初のみ）
        if i == 0:
            video_start = etree.Element('{http://www.dji.com/wpmz/1.0.6}action')
            etree.SubElement(video_start, '{http://www.dji.com/wpmz/1.0.6}actionId').text = '998'
            etree.SubElement(video_start, '{http://www.dji.com/wpmz/1.0.6}actionActuatorFunc').text = 'startRecordVideo'
            param = etree.SubElement(video_start, '{http://www.dji.com/wpmz/1.0.6}actionActuatorFuncParam')
            etree.SubElement(param, '{http://www.dji.com/wpmz/1.0.6}payloadPositionIndex').text = '0'
            etree.SubElement(param, '{http://www.dji.com/wpmz/1.0.6}videoType').text = 'zoom'
            new_actions.insert(0, video_start)

        # ホバリング追加（各ポイント共通）
        hover_action = etree.Element('{http://www.dji.com/wpmz/1.0.6}action')
        etree.SubElement(hover_action, '{http://www.dji.com/wpmz/1.0.6}actionId').text = '999'
        etree.SubElement(hover_action, '{http://www.dji.com/wpmz/1.0.6}actionActuatorFunc').text = 'stayForSeconds'
        param = etree.SubElement(hover_action, '{http://www.dji.com/wpmz/1.0.6}actionActuatorFuncParam')
        etree.SubElement(param, '{http://www.dji.com/wpmz/1.0.6}stayTime').text = '10.0'
        new_actions.append(hover_action)

        # 動画停止（最後のみ）
        if i == len(placemarks) - 1:
            video_stop = etree.Element('{http://www.dji.com/wpmz/1.0.6}action')
            etree.SubElement(video_stop, '{http://www.dji.com/wpmz/1.0.6}actionId').text = '997'
            etree.SubElement(video_stop, '{http://www.dji.com/wpmz/1.0.6}actionActuatorFunc').text = 'stopRecordVideo'
            param = etree.SubElement(video_stop, '{http://www.dji.com/wpmz/1.0.6}actionActuatorFuncParam')
            etree.SubElement(param, '{http://www.dji.com/wpmz/1.0.6}payloadPositionIndex').text = '0'
            etree.SubElement(param, '{http://www.dji.com/wpmz/1.0.6}videoType').text = 'zoom'
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
