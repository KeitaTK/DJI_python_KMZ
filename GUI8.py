import os
import shutil
import zipfile
import glob
import tkinter as tk
from tkinter import messagebox
from tkinterdnd2 import TkinterDnD, DND_FILES
from lxml import etree

# --- GUI フォント設定 ---
LABEL_FONT = ("Meiryo", 12)

# --- movie4.py 由来の設定パラメータ ---
BASE_HEIGHT = 613.5            # 高度補正[m]
NEW_HEIGHT_MODE = 'als'        # 'als' or 'relativeToStartPoint'
AUTO_FLIGHT_SPEED = 15         # 自動飛行速度[m/s]
TAKEOFF_SECURITY_HEIGHT = 20   # 安全離陸高度[m]
TURN_MODE = 'toPointAndStopWithDiscontinuityCurvature'
HOVER_SECONDS = 2              # ホバリング[秒]
GIMBAL_PITCH = -90             # ジンバルピッチ[°]
YAW_ANGLE = 87.37              # 機体ヨー角度[°]
ZOOM_FACTOR = 5                # ズーム倍率
ENABLE_ZOOM = True
ENABLE_WIDE = False
ENABLE_IR = False
PAYLOAD_POSITION_INDEX = '0'   # ペイロード位置インデックス

def convert_and_extract_kmz(kmz_path):
    """
    KMZ → ZIP 展開。展開先は同フォルダ内の Converted フォルダ。
    """
    if not kmz_path.lower().endswith(".kmz"):
        return False, "拡張子が .kmz ではありません"
    base_dir = os.path.dirname(kmz_path)
    output_dir = os.path.join(base_dir, "Converted")
    zip_path = os.path.splitext(kmz_path)[0] + ".zip"

    try:
        # 一時 ZIP を作成して展開フォルダを初期化
        shutil.copy(kmz_path, zip_path)
        if os.path.isdir(output_dir):
            shutil.rmtree(output_dir)
        os.makedirs(output_dir)
        with zipfile.ZipFile(zip_path, "r") as z:
            z.extractall(output_dir)
        os.remove(zip_path)
        return True, output_dir
    except Exception as e:
        return False, f"展開失敗: {e}"

def repack_to_kmz_only(folder_path, original_kmz_path):
    """
    Converted フォルダ内を ZIP → 新 KMZ に変換し、
    展開フォルダ内は再パック後の KMZ のみを残す。
    """
    new_name = os.path.splitext(os.path.basename(original_kmz_path))[0] + "_movie.kmz"
    new_kmz_path = os.path.join(folder_path, new_name)
    temp_zip = new_kmz_path.replace(".kmz", ".zip")

    # ZIP 圧縮
    with zipfile.ZipFile(temp_zip, "w", zipfile.ZIP_DEFLATED) as z:
        for root_dir, _, files in os.walk(folder_path):
            for f in files:
                full = os.path.join(root_dir, f)
                rel = os.path.relpath(full, folder_path)
                z.write(full, rel)

    # ZIP → KMZ リネーム
    if os.path.exists(new_kmz_path):
        os.remove(new_kmz_path)
    os.rename(temp_zip, new_kmz_path)

    # フォルダ内を一掃し、KMZ のみ残す
    for item in os.listdir(folder_path):
        path = os.path.join(folder_path, item)
        if not item.endswith(".kmz"):
            if os.path.isdir(path):
                shutil.rmtree(path)
            else:
                os.remove(path)
    return new_kmz_path

def convert_kml_logic(kml_input_path, kml_output_path):
    """
    lxml を用いて KML を解析・編集し、
    新しい KML として出力後、元の KML を削除。
    """
    ns = {
        'kml': 'http://www.opengis.net/kml/2.2',
        'wpml': 'http://www.dji.com/wpmz/1.0.6'
    }
    tree = etree.parse(kml_input_path)
    root = tree.getroot()
    folder = root.find('.//kml:Folder', ns)

    # 高度モード変更
    hm = root.find('.//wpml:waylineCoordinateSysParam/wpml:heightMode', ns)
    if hm is not None and hm.text != NEW_HEIGHT_MODE:
        hm.text = NEW_HEIGHT_MODE

    # 各 Placemark の編集
    for pm in root.findall('.//kml:Placemark', ns):
        # 高度補正
        for tag in ('wpml:height', 'wpml:ellipsoidHeight'):
            he = pm.find(tag, ns)
            if he is not None:
                try:
                    he.text = str(float(he.text) + BASE_HEIGHT)
                except:
                    pass
        # 速度設定
        sp = pm.find('wpml:waypointSpeed', ns)
        if sp is not None:
            sp.text = str(AUTO_FLIGHT_SPEED)
        # ヘディング設定
        hp = pm.find('wpml:waypointHeadingParam', ns)
        if hp is None:
            hp = etree.SubElement(pm, '{http://www.dji.com/wpmz/1.0.6}waypointHeadingParam')
        for tag, val in [
            ('waypointHeadingMode', 'followWayline'),
            ('waypointHeadingAngle', str(YAW_ANGLE)),
            ('waypointPoiPoint', '0,0,0'),
            ('waypointHeadingPoiIndex', '0')
        ]:
            e = hp.find(f'wpml:{tag}', ns)
            if e is None:
                etree.SubElement(hp, f'{{http://www.dji.com/wpmz/1.0.6}}{tag}').text = val
            else:
                e.text = val

        # アクション編集
        ag = pm.find('.//wpml:actionGroup', ns)
        if ag is None:
            ag = etree.SubElement(pm, '{http://www.dji.com/wpmz/1.0.6}actionGroup')
        # 既存アクションを削除
        for old in ag.findall('wpml:action', ns):
            ag.remove(old)
        new_actions = []

        # 動画撮影開始（最初のWPのみ）
        idx_elem = pm.find('wpml:actionGroup/actionGroupId', ns)
        idx = int(idx_elem.text) if idx_elem is not None else 0
        if idx == 0:
            def add_video(id_offset, vtype):
                act = etree.Element('{http://www.dji.com/wpmz/1.0.6}action')
                etree.SubElement(act, '{http://www.dji.com/wpmz/1.0.6}actionId').text = str(998 + id_offset)
                etree.SubElement(act, '{http://www.dji.com/wpmz/1.0.6}actionActuatorFunc').text = 'startRecordVideo'
                p = etree.SubElement(act, '{http://www.dji.com/wpmz/1.0.6}actionActuatorFuncParam')
                etree.SubElement(p, '{http://www.dji.com/wpmz/1.0.6}payloadPositionIndex').text = PAYLOAD_POSITION_INDEX
                etree.SubElement(p, '{http://www.dji.com/wpmz/1.0.6}videoType').text = vtype
                if vtype == 'zoom':
                    etree.SubElement(p, '{http://www.dji.com/wpmz/1.0.6}zoomFactor').text = str(ZOOM_FACTOR)
                new_actions.append(act)
            if ENABLE_ZOOM: add_video(0, 'zoom')
            if ENABLE_WIDE: add_video(1, 'wide')
            if ENABLE_IR:   add_video(2, 'ir')

        # ジンバル制御
        ga = etree.Element('{http://www.dji.com/wpmz/1.0.6}action')
        etree.SubElement(ga, '{http://www.dji.com/wpmz/1.0.6}actionId').text = '1001'
        etree.SubElement(ga, '{http://www.dji.com/wpmz/1.0.6}actionActuatorFunc').text = 'gimbalRotate'
        gp = etree.SubElement(ga, '{http://www.dji.com/wpmz/1.0.6}actionActuatorFuncParam')
        for tag, val in [
            ('gimbalRotateMode', 'absoluteAngle'),
            ('gimbalPitchRotateEnable', '1'),
            ('gimbalPitchRotateAngle', str(GIMBAL_PITCH)),
            ('gimbalRollRotateEnable', '0'),
            ('gimbalRollRotateAngle', '0'),
            ('gimbalYawRotateEnable', '0'),
            ('gimbalYawRotateAngle', '0'),
            ('gimbalRotateTimeEnable', '0'),
            ('gimbalRotateTime', '0'),
            ('payloadPositionIndex', PAYLOAD_POSITION_INDEX)
        ]:
            etree.SubElement(gp, f'{{http://www.dji.com/wpmz/1.0.6}}{tag}').text = val
        new_actions.append(ga)

        # ホバリング
        ha = etree.Element('{http://www.dji.com/wpmz/1.0.6}action')
        etree.SubElement(ha, '{http://www.dji.com/wpmz/1.0.6}actionId').text = '999'
        etree.SubElement(ha, '{http://www.dji.com/wpmz/1.0.6}actionActuatorFunc').text = 'stayForSeconds'
        hp2 = etree.SubElement(ha, '{http://www.dji.com/wpmz/1.0.6}actionActuatorFuncParam')
        etree.SubElement(hp2, '{http://www.dji.com/wpmz/1.0.6}stayTime').text = str(HOVER_SECONDS)
        etree.SubElement(hp2, '{http://www.dji.com/wpmz/1.0.6}gimbalPitchRotateAngle').text = str(GIMBAL_PITCH)
        etree.SubElement(hp2, '{http://www.dji.com/wpmz/1.0.6}aircraftHeading').text = str(YAW_ANGLE)
        new_actions.append(ha)

        # 新アクションを登録
        for act in new_actions:
            ag.append(act)

    # グローバル設定編集
    gs = folder.find('wpml:autoFlightSpeed', ns)
    if gs is not None:
        gs.text = str(AUTO_FLIGHT_SPEED)
    tkh = root.find('.//wpml:takeOffSecurityHeight', ns)
    if tkh is not None:
        tkh.text = str(TAKEOFF_SECURITY_HEIGHT)
    tm = folder.find('wpml:globalWaypointTurnMode', ns)
    if tm is not None:
        tm.text = TURN_MODE
    gh = folder.find('wpml:globalWaypointHeadingParam', ns)
    if gh is not None:
        for tag, val in [('waypointHeadingMode', 'followWayline'), ('waypointHeadingAngle', str(YAW_ANGLE))]:
            e = gh.find(f'wpml:{tag}', ns)
            if e is not None:
                e.text = val

    # 変更後の KML を保存
    tree.write(kml_output_path, encoding='utf-8', pretty_print=True, xml_declaration=True)
    # 元の KML を削除
    os.remove(kml_input_path)

def process_kmz_and_edit_kml(kmz_path):
    ok, res = convert_and_extract_kmz(kmz_path)
    if not ok:
        return False, res
    work_dir = res

    # テンプレート KML を検索
    kmls = glob.glob(os.path.join(work_dir, "**", "wpmz", "template.kml"), recursive=True)
    if not kmls:
        shutil.rmtree(work_dir)
        return False, "template.kml が見つかりません。"

    # 各 KML を変換
    for src in kmls:
        dst = src.replace("template.kml", "template_movie.kml")
        convert_kml_logic(src, dst)

    # 再パックして KMZ のみ出力
    new_kmz = repack_to_kmz_only(work_dir, kmz_path)
    return True, f"変換完了: {new_kmz}"

def drop_file(event):
    paths = event.data.strip().split()
    if len(paths) != 1:
        messagebox.showwarning("エラー", "ファイルは1つだけドロップしてください。", parent=root)
        return
    path = paths[0].strip("{}")
    label.config(text=f"ファイル：{os.path.basename(path)}")
    if path.lower().endswith(".kmz"):
        ok, msg = process_kmz_and_edit_kml(path)
        if ok:
            messagebox.showinfo("完了", msg, parent=root)
        else:
            messagebox.showerror("エラー", msg, parent=root)
    else:
        messagebox.showinfo("情報", ".kmz のみ対応", parent=root)

# --- GUI 初期化 ---
root = TkinterDnD.Tk()
root.title("KMZ 展開・movie4 変換・再パックツール")
root.geometry("500x300")

label = tk.Label(
    root,
    text="ここに .kmz ファイルをドラッグ＆ドロップ",
    bg="lightgray",
    width=60,
    height=10,
    anchor="nw",
    justify="left",
    font=LABEL_FONT
)
label.pack(padx=20, pady=40, fill="both", expand=True)

label.drop_target_register(DND_FILES)
label.dnd_bind("<<Drop>>", drop_file)

root.mainloop()
