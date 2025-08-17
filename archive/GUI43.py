#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
GUI43_imageFormat.py

ATL→ASL 変換＋撮影制御ツール v4.3 (imageFormat方式)
ヨー軸固定＋写真時ジンバル前方固定対応
"""

import os, shutil, zipfile, threading, tkinter as tk
from tkinter import ttk, messagebox, scrolledtext
from tkinterdnd2 import TkinterDnD, DND_FILES
from lxml import etree
from datetime import datetime
import pyperclip

# -------------------------------------------------------------------
# 定数
# -------------------------------------------------------------------
NS = {
    "kml": "http://www.opengis.net/kml/2.2",
    "wpml": "http://www.dji.com/wpmz/1.0.6"
}
HEIGHT_OPTIONS = {
    "613.5 – 事務所前": 613.5,
    "962.02 – 烏帽子": 962.02,
    "その他 – 手動入力": "custom"
}
YAW_OPTIONS = {
    "1Q: 88.00°": 88.00,
    "2Q: 96.92°": 96.92,
    "4Q: 87.31°": 87.31,
    "手動入力": "custom"
}
IMAGE_FORMAT_OPTIONS = ["zoom", "wide", "ir"]
REFERENCE_POINTS = {
    "本部": (136.5559522506280, 36.0729517605894, 612.2),
    "烏帽子": (136.560000000000, 36.075000000000, 962.02),
}
DEVIATION_THRESHOLD = {"lat": 0.00018, "lng": 0.00022, "alt": 20.0}

# -------------------------------------------------------------------
# ユーティリティ関数
# -------------------------------------------------------------------
def extract_kmz(path, work_dir="_kmz_work"):
    if os.path.exists(work_dir):
        shutil.rmtree(work_dir)
    os.makedirs(work_dir)
    with zipfile.ZipFile(path, "r") as zf:
        zf.extractall(work_dir)
    return work_dir

def prepare_output_dirs(input_kmz, offset):
    base = os.path.splitext(os.path.basename(input_kmz))[0]
    sign = "+" if offset >= 0 else "-"
    root = os.path.dirname(input_kmz)
    out_root = os.path.join(root, f"{base}_asl{sign}{abs(offset)}")
    if os.path.exists(out_root):
        shutil.rmtree(out_root)
    os.makedirs(out_root)
    wpmz_dir = os.path.join(out_root, "wpmz")
    os.makedirs(wpmz_dir)
    return out_root, wpmz_dir

def repackage_to_kmz(out_root, input_kmz):
    base = os.path.splitext(os.path.basename(input_kmz))[0]
    out_kmz = os.path.join(os.path.dirname(out_root), f"{base}_Converted.kmz")
    tmp = out_kmz + ".zip"
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zf:
        for root, _, files in os.walk(out_root):
            for f in files:
                full = os.path.join(root, f)
                rel = os.path.relpath(full, out_root)
                zf.write(full, rel)
    if os.path.exists(out_kmz):
        os.remove(out_kmz)
    os.rename(tmp, out_kmz)
    return out_kmz

def calculate_deviation(reference_coords, current_coords):
    ref_lng, ref_lat, ref_alt = reference_coords
    cur_lng, cur_lat, cur_alt = current_coords
    return cur_lng - ref_lng, cur_lat - ref_lat, cur_alt - ref_alt

def check_deviation_safety(deviation):
    dev_lng, dev_lat, dev_alt = deviation
    if (abs(dev_lng) > DEVIATION_THRESHOLD["lng"] or
        abs(dev_lat) > DEVIATION_THRESHOLD["lat"] or
        abs(dev_alt) > DEVIATION_THRESHOLD["alt"]):
        return False, (f"偏差が閾値を超えています:\n"
                       f"経度偏差: {dev_lng:.8f}°\n"
                       f"緯度偏差: {dev_lat:.8f}°\n"
                       f"標高偏差: {dev_alt:.2f}m")
    return True, None

def log_conversion_details(log, params, placemark_count):
    log.insert(tk.END, "="*60 + "\n")
    log.insert(tk.END, "変換設定詳細:\n")
    log.insert(tk.END, "="*60 + "\n")
    log.insert(tk.END,
               f"ATL→ASL変換: {'有効' if params.get('do_asl') else '無効'}" +
               (f" (オフセット: {params.get('offset')}m)" if params.get('do_asl') else "") + "\n")
    mode = "写真撮影" if params.get("do_photo") else "動画撮影" if params.get("do_video") else "設定なし"
    log.insert(tk.END, f"撮影モード: {mode}\n")
    imgfmts = ",".join(params.get("image_formats", []))
    log.insert(tk.END, f"imageFormat: {imgfmts or '既存設定を維持'}\n")
    if params.get("gimbal_pitch_ctrl"):
        log.insert(tk.END, f"ジンバルピッチ制御: 有効 (モード: {params.get('gimbal_pitch_mode')}, "
                            f"角度: {params.get('gimbal_pitch_angle')}°)\n")
    else:
        log.insert(tk.END, "ジンバルピッチ制御: 無効\n")
    if params.get("yaw_angle") is not None:
        log.insert(tk.END, f"ヨー固定: 有効 ({params.get('yaw_angle')}°)\n")
    else:
        log.insert(tk.END, "ヨー固定: 無効\n")
    if params.get("coordinate_deviation"):
        dev = params.get("coordinate_deviation")
        log.insert(tk.END, f"偏差補正: 有効 (経度: {dev[0]:.8f}°, 緯度: {dev[1]:.8f}°, 標高: {dev[2]:.2f}m)\n")
    else:
        log.insert(tk.END, "偏差補正: 無効\n")
    log.insert(tk.END, f"\n処理対象ウェイポイント数: {placemark_count}\n")
    log.insert(tk.END, "="*60 + "\n\n")

# -------------------------------------------------------------------
# アクション生成関数
# -------------------------------------------------------------------
def create_yaw_rotate_action(yaw_angle):
    action = etree.Element("{http://www.dji.com/wpmz/1.0.6}action")
    etree.SubElement(action, "{http://www.dji.com/wpmz/1.0.6}actionId").text = "0"
    etree.SubElement(action, "{http://www.dji.com/wpmz/1.0.6}actionActuatorFunc").text = "rotateYaw"
    param = etree.SubElement(action, "{http://www.dji.com/wpmz/1.0.6}actionActuatorFuncParam")
    etree.SubElement(param, "{http://www.dji.com/wpmz/1.0.6}aircraftHeading").text = str(yaw_angle)
    etree.SubElement(param, "{http://www.dji.com/wpmz/1.0.6}aircraftPathMode").text = "counterClockwise"
    return action

def create_gimbal_rotate_action(angle):
    action = etree.Element("{http://www.dji.com/wpmz/1.0.6}action")
    etree.SubElement(action, "{http://www.dji.com/wpmz/1.0.6}actionId").text = "0"
    etree.SubElement(action, "{http://www.dji.com/wpmz/1.0.6}actionActuatorFunc").text = "gimbalRotate"
    param = etree.SubElement(action, "{http://www.dji.com/wpmz/1.0.6}actionActuatorFuncParam")
    etree.SubElement(param, "{http://www.dji.com/wpmz/1.0.6}gimbalRotateMode").text = "absoluteAngle"
    etree.SubElement(param, "{http://www.dji.com/wpmz/1.0.6}gimbalPitchRotateEnable").text = "1"
    etree.SubElement(param, "{http://www.dji.com/wpmz/1.0.6}gimbalPitchRotateAngle").text = str(angle)
    etree.SubElement(param, "{http://www.dji.com/wpmz/1.0.6}gimbalRollRotateEnable").text = "0"
    etree.SubElement(param, "{http://www.dji.com/wpmz/1.0.6}gimbalRollRotateAngle").text = "0"
    etree.SubElement(param, "{http://www.dji.com/wpmz/1.0.6}gimbalYawRotateEnable").text = "0"
    etree.SubElement(param, "{http://www.dji.com/wpmz/1.0.6}gimbalYawRotateAngle").text = "0"
    etree.SubElement(param, "{http://www.dji.com/wpmz/1.0.6}gimbalRotateTimeEnable").text = "0"
    etree.SubElement(param, "{http://www.dji.com/wpmz/1.0.6}gimbalRotateTime").text = "0"
    etree.SubElement(param, "{http://www.dji.com/wpmz/1.0.6}payloadPositionIndex").text = "0"
    return action

def create_start_record_action():
    action = etree.Element("{http://www.dji.com/wpmz/1.0.6}action")
    etree.SubElement(action, "{http://www.dji.com/wpmz/1.0.6}actionId").text = "0"
    etree.SubElement(action, "{http://www.dji.com/wpmz/1.0.6}actionActuatorFunc").text = "startRecord"
    etree.SubElement(action, "{http://www.dji.com/wpmz/1.0.6}actionActuatorFuncParam")
    return action

def create_stop_record_action():
    action = etree.Element("{http://www.dji.com/wpmz/1.0.6}action")
    etree.SubElement(action, "{http://www.dji.com/wpmz/1.0.6}actionId").text = "0"
    etree.SubElement(action, "{http://www.dji.com/wpmz/1.0.6}actionActuatorFunc").text = "stopRecord"
    etree.SubElement(action, "{http://www.dji.com/wpmz/1.0.6}actionActuatorFuncParam")
    return action

def update_oriented_shoot_gimbal(shoot_action, gimbal_angle):
    param = shoot_action.find("wpml:actionActuatorFuncParam", NS)
    if param is not None:
        pitch_elem = param.find("wpml:gimbalPitchRotateAngle", NS)
        if pitch_elem is not None:
            pitch_elem.text = str(gimbal_angle)

def update_payload_param(tree, image_formats):
    if not image_formats:
        return
    fmt_str = ",".join(image_formats)
    payload = tree.find(".//wpml:payloadParam", NS)
    if payload is None:
        fld = tree.find(".//kml:Folder", NS)
        if fld is not None:
            payload = etree.SubElement(fld, "{http://www.dji.com/wpmz/1.0.6}payloadParam")
    if payload is not None:
        img = payload.find("wpml:imageFormat", NS)
        if img is None:
            img = etree.SubElement(payload, "{http://www.dji.com/wpmz/1.0.6}imageFormat")
        img.text = fmt_str

# -------------------------------------------------------------------
# アクション再編成
# -------------------------------------------------------------------
def reorganize_actions(action_group, params, log,
                       waypoint_index, is_first_waypoint, is_last_waypoint):
    # 分類
    existing = {"rotateYaw": [], "gimbalRotate": [], "orientedShoot": [], 
                "startRecord": [], "stopRecord": [], "other": []}
    for act in action_group.findall("wpml:action", NS):
        fn = act.find("wpml:actionActuatorFunc", NS)
        if fn is not None and fn.text in existing:
            existing[fn.text].append(act)
        else:
            existing["other"].append(act)
    # 削除
    for act in action_group.findall("wpml:action", NS):
        action_group.remove(act)
    new_actions = []
    action_log = []

    # 1. ヨー回転
    yaw_ang = params.get("yaw_angle")
    if yaw_ang is not None:
        if existing["rotateYaw"]:
            ya = existing["rotateYaw"][0]
            p = ya.find("wpml:actionActuatorFuncParam", NS)
            if p is not None:
                he = p.find("wpml:aircraftHeading", NS)
                if he is not None:
                    he.text = str(yaw_ang)
        else:
            ya = create_yaw_rotate_action(yaw_ang)
            new_actions.append(ya)
        action_log.append(f"ヨー回転: {yaw_ang}°")

    # 2. 動画処理
    if params.get("do_video"):
        if is_first_waypoint:
            # ジンバルピッチ
            if params.get("gimbal_pitch_ctrl") and params.get("gimbal_pitch_angle") is not None:
                ga = params["gimbal_pitch_angle"]
                if existing["gimbalRotate"]:
                    gact = existing["gimbalRotate"][0]
                    p = gact.find("wpml:actionActuatorFuncParam", NS)
                    if p is not None:
                        pe = p.find("wpml:gimbalPitchRotateAngle", NS)
                        if pe is not None:
                            pe.text = str(ga)
                else:
                    gact = create_gimbal_rotate_action(ga)
                    new_actions.append(gact)
                action_log.append(f"ジンバルピッチ: {ga}°")
            # 録画開始
            s = existing["startRecord"][0] if existing["startRecord"] else create_start_record_action()
            new_actions.append(s)
            action_log.append("録画開始")
        if is_last_waypoint:
            st = existing["stopRecord"][0] if existing["stopRecord"] else create_stop_record_action()
            new_actions.append(st)
            action_log.append("録画停止")

    # 3. 写真撮影
    elif params.get("do_photo"):
        # ジンバルピッチコントロール
        if params.get("gimbal_pitch_ctrl") and params.get("gimbal_pitch_angle") is not None:
            ga = params["gimbal_pitch_angle"]
            if existing["gimbalRotate"]:
                gact = existing["gimbalRotate"][0]
                p = gact.find("wpml:actionActuatorFuncParam", NS)
                if p is not None:
                    pe = p.find("wpml:gimbalPitchRotateAngle", NS)
                    if pe is not None:
                        pe.text = str(ga)
            else:
                gact = create_gimbal_rotate_action(ga)
                new_actions.append(gact)
            action_log.append(f"ジンバルピッチ: {ga}°")

        # 追加: ヨー固定かつ写真モードならジンバルヨー前方固定
        if params.get("yaw_angle") is not None and params.get("gimbal_yaw_fixed"):
            if existing["gimbalRotate"]:
                yact = existing["gimbalRotate"][0]
                p = yact.find("wpml:actionActuatorFuncParam", NS)
                etree.SubElement(p, "{http://www.dji.com/wpmz/1.0.6}gimbalYawRotateEnable").text = "1"
                etree.SubElement(p, "{http://www.dji.com/wpmz/1.0.6}gimbalYawRotateAngle").text = "0"
            else:
                yact = etree.Element("{http://www.dji.com/wpmz/1.0.6}action")
                etree.SubElement(yact, "{http://www.dji.com/wpmz/1.0.6}actionId").text = "0"
                etree.SubElement(yact, "{http://www.dji.com/wpmz/1.0.6}actionActuatorFunc").text = "gimbalRotate"
                p = etree.SubElement(yact, "{http://www.dji.com/wpmz/1.0.6}actionActuatorFuncParam")
                etree.SubElement(p, "{http://www.dji.com/wpmz/1.0.6}gimbalRotateMode").text = "absoluteAngle"
                etree.SubElement(p, "{http://www.dji.com/wpmz/1.0.6}gimbalPitchRotateEnable").text = "0"
                etree.SubElement(p, "{http://www.dji.com/wpmz/1.0.6}gimbalRollRotateEnable").text = "0"
                etree.SubElement(p, "{http://www.dji.com/wpmz/1.0.6}gimbalYawRotateEnable").text = "1"
                etree.SubElement(p, "{http://www.dji.com/wpmz/1.0.6}gimbalYawRotateAngle").text = "0"
                etree.SubElement(p, "{http://www.dji.com/wpmz/1.0.6}payloadPositionIndex").text = "0"
                idx = len(new_actions) - len(existing["orientedShoot"]) - len(existing["other"])
                new_actions.insert(idx, yact)
            action_log.append("ジンバルヨー固定: 前方(0°)")

        # orientedShoot
        for shoot in existing["orientedShoot"]:
            if params.get("gimbal_pitch_ctrl") and params.get("gimbal_pitch_angle") is not None:
                update_oriented_shoot_gimbal(shoot, params["gimbal_pitch_angle"])
                action_log.append("写真撮影(角度統一)")
            else:
                action_log.append("写真撮影")
            new_actions.append(shoot)

    # 4. その他
    for oth in existing["other"]:
        new_actions.append(oth)
        action_log.append("その他アクション")

    # 再配置 & ID再設定
    for i, act in enumerate(new_actions):
        aid = act.find("wpml:actionId", NS)
        if aid is not None:
            aid.text = str(i)
        action_group.append(act)

    start_i = action_group.find("wpml:actionGroupStartIndex", NS)
    end_i = action_group.find("wpml:actionGroupEndIndex", NS)
    if new_actions and start_i is not None and end_i is not None:
        si = int(start_i.text or 0)
        end_i.text = str(si + len(new_actions) - 1)

    if action_log:
        log.insert(tk.END, f" ウェイポイント{waypoint_index}: {' → '.join(action_log)}\n")

# -------------------------------------------------------------------
# KML変換処理本体
# -------------------------------------------------------------------
def convert_kml(tree, params, log):
    offset = params.get("offset", 0.0)
    deviation = params.get("coordinate_deviation")
    do_asl = params.get("do_asl", False)

    hmode = tree.find(
        "./kml:Document/kml:Folder/wpml:waylineCoordinateSysParam/wpml:heightMode",
        NS
    )
    height_mode = hmode.text if hmode is not None else None

    gh = tree.find(".//wpml:globalHeight", NS)
    orig_gh = float(gh.text) if gh is not None else 0.0

    if params.get("image_formats"):
        update_payload_param(tree, params["image_formats"])
        log.insert(tk.END,
                   f"imageFormat更新: {','.join(params['image_formats'])}\n")

    log.insert(tk.END,
               f"\n変換処理開始: globalHeight={orig_gh}m, heightMode={height_mode}\n")

    pms = tree.findall(".//kml:Placemark", NS)
    total = len(pms)
    log.insert(tk.END, "\nウェイポイント処理:\n")

    for i, pm in enumerate(pms):
        wp_idx = int(pm.find("wpml:index", NS).text)
        is_first = (i == 0)
        is_last = (i == total - 1)

        # 座標・高度
        c = pm.find(".//kml:coordinates", NS)
        if not c or not c.text:
            continue
        parts = list(map(float, c.text.strip().split(",")))
        lng, lat = parts[0], parts[1]
        alt = parts[2] if len(parts) > 2 else float(pm.find("wpml:height", NS).text or 0.0)
        orig_alt = alt

        if deviation:
            dl, da, dz = deviation
            lng += dl; lat += da; alt += dz
        if do_asl and height_mode == "relativeToStartPoint":
            alt += offset

        c.text = f"{lng},{lat},{alt}"
        for tag in ("height", "ellipsoidHeight"):
            e = pm.find(f"wpml:{tag}", NS)
            if e is not None:
                e.text = str(alt)

        ag = pm.find(".//wpml:actionGroup", NS)
        if ag is not None:
            reorganize_actions(ag, params, log, wp_idx, is_first, is_last)

        # ExtendedData
        update_extended_data(pm, params)

        log.insert(tk.END,
                   f" ウェイポイント{wp_idx}: 高度 {orig_alt:.1f}→{alt:.1f}m\n")

    # globalHeight & heightMode 更新
    new_gh = orig_gh
    if do_asl and height_mode == "relativeToStartPoint":
        new_gh += offset
    if deviation:
        new_gh += deviation[2]
    if gh is not None:
        gh.text = str(new_gh)
    for hm in tree.findall(".//wpml:heightMode", NS):
        if do_asl and height_mode == "relativeToStartPoint":
            hm.text = "EGM96"

    log.insert(tk.END,
               f"\nglobalHeight更新: {orig_gh}→{new_gh}m\n処理完了\n")

# -------------------------------------------------------------------
# ExtendedData 更新
# -------------------------------------------------------------------
def update_extended_data(placemark, params):
    ed = placemark.find(".//kml:ExtendedData", NS)
    if ed is None:
        ed = etree.SubElement(placemark, "{http://www.opengis.net/kml/2.2}ExtendedData")
    for d in ed.findall("kml:Data", NS):
        ed.remove(d)
    mode = "photo" if params.get("do_photo") else "video" if params.get("do_video") else ""
    if mode:
        md = etree.SubElement(ed, "{http://www.opengis.net/kml/2.2}Data", name="mode")
        md.text = mode
    gb = etree.SubElement(ed, "{http://www.opengis.net/kml/2.2}Data", name="gimbal")
    gb.text = str(params.get("do_gimbal", True))
    ya = params.get("yaw_angle")
    if ya is not None:
        ye = etree.SubElement(ed, "{http://www.opengis.net/kml/2.2}Data", name="yaw")
        ye.text = str(ya)
    ht = params.get("hover_time", 0)
    if ht > 0:
        ho = etree.SubElement(ed, "{http://www.opengis.net/kml/2.2}Data", name="hover_time")
        ho.text = str(ht)

# -------------------------------------------------------------------
# KMZ 全体処理
# -------------------------------------------------------------------
def process_kmz(path, log, params):
    try:
        log.insert(tk.END, f"処理開始: {os.path.basename(path)}...\n")
        wd = extract_kmz(path)
        kml = os.path.join(wd, "wpmz", "template.kml")
        if not os.path.exists(kml):
            kml = os.path.join(wd, "template.kml")
        tree = etree.parse(kml, etree.XMLParser(remove_blank_text=True))
        pcs = len(tree.findall(".//kml:Placemark", NS))
        log_conversion_details(log, params, pcs)
        out_root, outdir = prepare_output_dirs(path, params["offset"])
        actual = convert_kml(tree, params, log)
        out_kml = os.path.join(outdir, os.path.basename(kml))
        tree.write(out_kml, encoding="utf-8", pretty_print=True, xml_declaration=True)
        res = os.path.join(os.path.dirname(kml), "res")
        if os.path.isdir(res):
            shutil.copytree(res, os.path.join(outdir, "res"))
        out = repackage_to_kmz(out_root, path)
        log.insert(tk.END, f"\n変換完了: {out}\n")
        messagebox.showinfo("完了", f"変換完了:\n{out}")
    except Exception as e:
        log.insert(tk.END, f"エラー: {e}\n")
        messagebox.showerror("エラー", str(e))
    finally:
        if os.path.exists("_kmz_work"):
            shutil.rmtree("_kmz_work")

# -------------------------------------------------------------------
# UI
# -------------------------------------------------------------------
class AppUI(ttk.Frame):
    def __init__(self, master, controller):
        super().__init__(master, padding=10)
        self.controller = controller
        self._create_vars()
        self._create_widgets()
        self._grid_widgets()

    def _create_vars(self):
        self.height_choice_var = tk.StringVar()
        self.height_entry_var = tk.StringVar()
        self.asl_var = tk.BooleanVar(value=False)
        self.speed_var = tk.IntVar(value=15)
        self.photo_var = tk.BooleanVar(value=False)
        self.video_var = tk.BooleanVar(value=False)
        self.video_suffix_var = tk.StringVar(value="video_01")
        self.image_format_vars = {fmt: tk.BooleanVar(value=False) for fmt in IMAGE_FORMAT_OPTIONS}
        self.gimbal_var = tk.BooleanVar(value=True)
        self.gim_pitch_ctrl_var = tk.BooleanVar(value=False)
        self.gim_pitch_choice_var = tk.StringVar()
        self.gim_pitch_entry_var = tk.StringVar()
        self.yaw_fix_var = tk.BooleanVar(value=False)
        self.yaw_choice_var = tk.StringVar()
        self.yaw_entry_var = tk.StringVar()
        self.hover_var = tk.BooleanVar(value=False)
        self.hover_time_var = tk.StringVar(value="2")
        self.deviation_var = tk.BooleanVar(value=False)
        self.ref_point_var = tk.StringVar(value="本部")
        self.today_lng_var = tk.StringVar(value="136.555")
        self.today_lat_var = tk.StringVar(value="36.072")
        self.today_alt_var = tk.StringVar(value="0")

    def _create_widgets(self):
        # ASL
        self.asl_check = ttk.Checkbutton(self, text="ATL→ASL変換",
                                         variable=self.asl_var,
                                         command=self.controller.update_ui_states)
        self.height_label = ttk.Label(self, text="基準高度:")
        self.height_combo = ttk.Combobox(self, textvariable=self.height_choice_var,
                                         values=list(HEIGHT_OPTIONS.keys()),
                                         state="readonly", width=20)
        self.height_combo.bind("<<ComboboxSelected>>", self.controller.update_ui_states)
        self.height_entry = ttk.Entry(self, textvariable=self.height_entry_var, width=10)

        # 速度
        self.speed_label = ttk.Label(self, text="速度 (1–15 m/s):")
        self.speed_spinbox = ttk.Spinbox(self, from_=1, to=15,
                                         textvariable=self.speed_var, width=5)

        # 撮影設定
        self.photo_check = ttk.Checkbutton(self, text="写真撮影",
                                           variable=self.photo_var,
                                           command=self.controller.update_ui_states)
        self.video_check = ttk.Checkbutton(self, text="動画撮影",
                                           variable=self.video_var,
                                           command=self.controller.update_ui_states)
        self.video_suffix_label = ttk.Label(self, text="動画ファイル名:")
        self.video_suffix_entry = ttk.Entry(self,
                                            textvariable=self.video_suffix_var, width=20)

        # imageFormat
        self.image_format_label = ttk.Label(self, text="imageFormat:")
        self.image_format_checks = {
            fmt: ttk.Checkbutton(self, text=fmt,
                                 variable=self.image_format_vars[fmt])
            for fmt in IMAGE_FORMAT_OPTIONS
        }

        # ジンバル制御
        self.gimbal_check = ttk.Checkbutton(self, text="ジンバル制御",
                                            variable=self.gimbal_var)

        # ジンバルピッチ制御
        self.gim_pitch_ctrl_check = ttk.Checkbutton(
            self, text="ジンバルピッチ制御",
            variable=self.gim_pitch_ctrl_var,
            command=self.controller.update_ui_states
        )
        self.gim_pitch_combo = ttk.Combobox(self,
                                            textvariable=self.gim_pitch_choice_var,
                                            values=["真下 (-90°)", "前 (0°)", "手動入力", "フリー"],
                                            state="readonly", width=15)
        self.gim_pitch_combo.bind("<<ComboboxSelected>>", self.controller.update_ui_states)
        self.gim_pitch_entry = ttk.Entry(self, textvariable=self.gim_pitch_entry_var, width=8)

        # ヨー固定
        self.yaw_fix_check = ttk.Checkbutton(self, text="ヨー固定",
                                             variable=self.yaw_fix_var,
                                             command=self.controller.update_ui_states)
        self.yaw_combo = ttk.Combobox(self, textvariable=self.yaw_choice_var,
                                      values=list(YAW_OPTIONS.keys()),
                                      state="readonly", width=15)
        self.yaw_combo.bind("<<ComboboxSelected>>", self.controller.update_ui_states)
        self.yaw_entry = ttk.Entry(self, textvariable=self.yaw_entry_var, width=8)

        # ホバリング
        self.hover_check = ttk.Checkbutton(self, text="ホバリング",
                                           variable=self.hover_var,
                                           command=self.controller.update_ui_states)
        self.hover_time_label = ttk.Label(self, text="ホバリング時間 (秒):")
        self.hover_time_entry = ttk.Entry(self, textvariable=self.hover_time_var, width=8)

        # 偏差補正
        self.deviation_check = ttk.Checkbutton(self, text="偏差補正",
                                               variable=self.deviation_var,
                                               command=self.controller.update_ui_states)
        self.ref_point_label = ttk.Label(self, text="基準位置:")
        self.ref_point_combo = ttk.Combobox(self, textvariable=self.ref_point_var,
                                            values=list(REFERENCE_POINTS.keys()),
                                            state="readonly", width=10)
        self.today_coords_label = ttk.Label(self, text="本日の値 (経度,緯度,標高):")
        self.today_lng_entry = ttk.Entry(self, textvariable=self.today_lng_var, width=12)
        self.today_lat_entry = ttk.Entry(self, textvariable=self.today_lat_var, width=10)
        self.today_alt_entry = ttk.Entry(self, textvariable=self.today_alt_var, width=8)
        self.copy_button = ttk.Button(self, text="コピー",
                                      command=self.controller.copy_reference_data, width=8)

    def _grid_widgets(self):
        # 行・列配置
        self.asl_check.grid(row=0, column=0, sticky="w", pady=5)
        self.height_label.grid(row=0, column=1, sticky="w")
        self.height_combo.grid(row=0, column=2, sticky="w")
        self.speed_label.grid(row=1, column=0, sticky="w", pady=5)
        self.speed_spinbox.grid(row=1, column=1, sticky="w")
        self.photo_check.grid(row=2, column=0, sticky="w")
        self.video_check.grid(row=2, column=1, sticky="w")
        self.video_suffix_label.grid(row=2, column=2, sticky="e", padx=(10,2))
        self.video_suffix_entry.grid(row=2, column=3, sticky="w")
        self.image_format_label.grid(row=3, column=0, sticky="w")
        for i, fmt in enumerate(IMAGE_FORMAT_OPTIONS):
            self.image_format_checks[fmt].grid(row=3, column=1+i, sticky="w", padx=5)
        self.gimbal_check.grid(row=4, column=0, sticky="w", pady=5)
        self.gim_pitch_ctrl_check.grid(row=5, column=0, sticky="w")
        self.gim_pitch_combo.grid(row=5, column=1, sticky="w", padx=5)
        self.gim_pitch_entry.grid(row=5, column=2, sticky="w")
        self.yaw_fix_check.grid(row=6, column=0, sticky="w")
        self.yaw_combo.grid(row=6, column=1, sticky="w", padx=5)
        self.yaw_entry.grid(row=6, column=2, sticky="w")
        self.hover_check.grid(row=7, column=0, sticky="w")
        self.hover_time_label.grid(row=7, column=1, sticky="e")
        self.hover_time_entry.grid(row=7, column=2, sticky="w")
        self.deviation_check.grid(row=8, column=0, sticky="w")
        self.ref_point_label.grid(row=8, column=1, sticky="e")
        self.ref_point_combo.grid(row=8, column=2, sticky="w")
        self.today_coords_label.grid(row=9, column=0, sticky="w")
        self.today_lng_entry.grid(row=9, column=1, sticky="w")
        self.today_lat_entry.grid(row=9, column=2, sticky="w")
        self.today_alt_entry.grid(row=9, column=3, sticky="w")
        self.copy_button.grid(row=9, column=4, sticky="w")

# -------------------------------------------------------------------
# Controller
# -------------------------------------------------------------------
class KmlConverterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("ATL→ASL 変換＋撮影制御ツール v4.3")
        self.root.geometry("820x800")
        self.ui = AppUI(self.root, self)
        self.ui.pack(fill="x", pady=(0,10))
        self._create_dnd_and_log_area()
        self.update_ui_states()

    def _create_dnd_and_log_area(self):
        frm = ttk.Frame(self.root, padding=10)
        frm.pack(fill="both", expand=True)
        self.drop_label = tk.Label(frm, text=".kmz をここにドロップ",
                                    bg="lightgray", width=70, height=5, relief=tk.RIDGE)
        self.drop_label.pack(pady=12, fill="x")
        self.drop_label.drop_target_register(DND_FILES)
        self.drop_label.dnd_bind("<<Drop>>", self.on_drop)
        log_frame = ttk.LabelFrame(frm, text="変換ログ")
        log_frame.pack(fill="both", expand=True)
        self.log_text = scrolledtext.ScrolledText(log_frame, height=20)
        self.log_text.pack(fill="both", expand=True)

    def update_ui_states(self, event=None):
        ui = self.ui
        # ASL
        if ui.asl_var.get():
            ui.height_label.grid()
            ui.height_combo.grid()
        else:
            ui.height_label.grid_remove()
            ui.height_combo.grid_remove()
        # 撮影モード排他
        if ui.photo_var.get():
            ui.video_var.set(False)
        if ui.video_var.get():
            ui.photo_var.set(False)
        # ジンバル制御
        ui.gimbal_check.config(state="disabled" if (ui.photo_var.get() or ui.video_var.get()) else "normal")
        # ジンバルピッチ表示
        if ui.gim_pitch_ctrl_var.get():
            ui.gim_pitch_combo.grid()
            ui.gim_pitch_entry.grid()
        else:
            ui.gim_pitch_combo.grid_remove()
            ui.gim_pitch_entry.grid_remove()
        # ヨー固定表示
        if ui.yaw_fix_var.get():
            ui.yaw_combo.grid()
            ui.yaw_entry.grid()
        else:
            ui.yaw_combo.grid_remove()
            ui.yaw_entry.grid_remove()
        # ホバリング表示
        if ui.hover_var.get():
            ui.hover_time_label.grid()
            ui.hover_time_entry.grid()
        else:
            ui.hover_time_label.grid_remove()
            ui.hover_time_entry.grid_remove()
        # 偏差補正表示
        if ui.deviation_var.get():
            ui.ref_point_label.grid()
            ui.ref_point_combo.grid()
            ui.today_coords_label.grid()
            ui.today_lng_entry.grid()
            ui.today_lat_entry.grid()
            ui.today_alt_entry.grid()
            ui.copy_button.grid()
        else:
            ui.ref_point_label.grid_remove()
            ui.ref_point_combo.grid_remove()
            ui.today_coords_label.grid_remove()
            ui.today_lng_entry.grid_remove()
            ui.today_lat_entry.grid_remove()
            ui.today_alt_entry.grid_remove()
            ui.copy_button.grid_remove()

    def on_drop(self, event):
        path = event.data.strip("{}")
        if not path.lower().endswith(".kmz"):
            messagebox.showwarning("警告", ".kmzファイルのみ対応")
            return
        try:
            params = self._get_params()
        except ValueError as e:
            messagebox.showerror("入力エラー", str(e))
            return
        threading.Thread(target=process_kmz,
                         args=(path, self.log_text, params),
                         daemon=True).start()

    def copy_reference_data(self):
        try:
            ref = self.ui.ref_point_var.get()
            coords = REFERENCE_POINTS[ref]
            now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            txt = f"{now},{coords[0]},{coords[1]},{coords[2]}"
            pyperclip.copy(txt)
            messagebox.showinfo("コピー完了", f"コピーしました:\n{txt}")
        except Exception as e:
            messagebox.showerror("エラー", str(e))

    def _get_params(self):
        ui = self.ui
        p = {}
        # ASL
        if ui.asl_var.get():
            v = HEIGHT_OPTIONS.get(ui.height_choice_var.get())
            if v == "custom":
                p["offset"] = float(ui.height_entry_var.get())
            else:
                p["offset"] = float(v)
        else:
            p["offset"] = 0.0
        p["do_asl"] = ui.asl_var.get()
        # imageFormat
        p["image_formats"] = [f for f, var in ui.image_format_vars.items() if var.get()]
        # ジンバルピッチ
        p["gimbal_pitch_ctrl"] = ui.gim_pitch_ctrl_var.get()
        if p["gimbal_pitch_ctrl"]:
            choice = ui.gim_pitch_choice_var.get()
            if choice.startswith("真下"):
                p["gimbal_pitch_angle"] = -90.0
                p["gimbal_pitch_mode"] = "-90"
            elif choice.startswith("前"):
                p["gimbal_pitch_angle"] = 0.0
                p["gimbal_pitch_mode"] = "0"
            elif choice == "手動入力":
                p["gimbal_pitch_angle"] = float(ui.gim_pitch_entry_var.get())
                p["gimbal_pitch_mode"] = "manual"
            else:
                p["gimbal_pitch_angle"] = None
                p["gimbal_pitch_mode"] = "free"
        else:
            p["gimbal_pitch_angle"] = None
            p["gimbal_pitch_mode"] = "free"
        # ヨー固定
        p["yaw_angle"] = None
        if ui.yaw_fix_var.get():
            yv = YAW_OPTIONS.get(ui.yaw_choice_var.get())
            if yv == "custom":
                p["yaw_angle"] = float(ui.yaw_entry_var.get())
            else:
                p["yaw_angle"] = float(yv)
            p["gimbal_yaw_fixed"] = ui.photo_var.get()
        else:
            p["gimbal_yaw_fixed"] = False
        # ホバリング
        p["hover_time"] = float(ui.hover_time_var.get()) if ui.hover_var.get() else 0.0
        # 偏差補正
        p["coordinate_deviation"] = None
        if ui.deviation_var.get():
            ref = REFERENCE_POINTS[ui.ref_point_var.get()]
            cur = (float(ui.today_lng_var.get()),
                   float(ui.today_lat_var.get()),
                   float(ui.today_alt_var.get()))
            dev = calculate_deviation(ref, cur)
            ok, msg = check_deviation_safety(dev)
            if not ok:
                raise ValueError(msg)
            p["coordinate_deviation"] = dev
        # その他
        p["do_photo"] = ui.photo_var.get()
        p["do_video"] = ui.video_var.get()
        p["video_suffix"] = ui.video_suffix_var.get()
        p["do_gimbal"] = ui.gimbal_var.get()
        p["speed"] = max(1, min(15, ui.speed_var.get()))
        return p

# -------------------------------------------------------------------
# main
# -------------------------------------------------------------------
def main():
    root = TkinterDnD.Tk()
    app = KmlConverterApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()
