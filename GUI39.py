#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
convert_height_gui_asl_mod.py  (GUI38 からの拡張版)

主な追加／変更点
1) ジンバルピッチ制御
   ・チェックボックス「ジンバルピッチ制御」を追加
   ・選択肢：真下(-90°)、前(0°)、手動入力、フリー
   ・手動入力選択時のみ角度入力欄を表示
   ・有効時は全ウェイポイントの写真撮影前に gimbalRotate を挿入し、
     既存ピッチ角を統一（-90/0/任意角）。フリー時は既存操作を維持

2) センサー選択の差分更新
   ・GUI のチェック状態と既存 KML 内 ExtendedData/sensors を比較し、
     過不足のみ追加／削除。チェック無し時は既存値を保持

3) 既存機能（ATL→ASL 変換、偏差補正、速度・ヨー固定など）はすべて維持
"""

import os
import shutil
import zipfile
import threading
import tkinter as tk
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

SENSOR_MODES = ["Wide", "Zoom", "IR"]

REFERENCE_POINTS = {
    "本部": (136.5559522506280, 36.0729517605894, 612.2),
    "烏帽子": (136.560000000000, 36.075000000000, 962.02)
}

DEVIATION_THRESHOLD = {"lat": 0.00018, "lng": 0.00022, "alt": 20.0}

# -------------------------------------------------------------------
# Model
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
    if (
        abs(dev_lng) > DEVIATION_THRESHOLD["lng"]
        or abs(dev_lat) > DEVIATION_THRESHOLD["lat"]
        or abs(dev_alt) > DEVIATION_THRESHOLD["alt"]
    ):
        return False, (
            f"偏差が閾値を超えています:\n"
            f"経度偏差: {dev_lng:.8f}°\n"
            f"緯度偏差: {dev_lat:.8f}°\n"
            f"標高偏差: {dev_alt:.2f}m"
        )
    return True, None


# -------------------------------------------------------------------
# KML 変換処理本体
# -------------------------------------------------------------------
def insert_gimbal_rotate(action_group, angle):
    """
    写真撮影直前に gimbalRotate アクションを挿入
    """
    actions = action_group.findall("wpml:action", NS)
    # 既存 actionId の最大値を取得
    max_id = max(
        (int(a.find("wpml:actionId", NS).text) for a in actions if a.find("wpml:actionId", NS) is not None),
        default=-1,
    )
    new_id = max_id + 1

    # gimbalRotate アクション生成
    act = etree.Element("{http://www.dji.com/wpmz/1.0.6}action")
    etree.SubElement(act, "{http://www.dji.com/wpmz/1.0.6}actionId").text = str(new_id)
    etree.SubElement(act, "{http://www.dji.com/wpmz/1.0.6}actionActuatorFunc").text = "gimbalRotate"

    param = etree.SubElement(act, "{http://www.dji.com/wpmz/1.0.6}actionActuatorFuncParam")
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

    # orientedShoot の直前に挿入
    inserted = False
    for idx, a in enumerate(actions):
        func = a.find("wpml:actionActuatorFunc", NS)
        if func is not None and func.text == "orientedShoot":
            action_group.insert(idx, act)
            inserted = True
            break
    if not inserted:
        action_group.append(act)


def unify_oriented_shoot_pitch(action_group, angle):
    """
    orientedShoot の Pitch を統一
    """
    for a in action_group.findall("wpml:action", NS):
        func = a.find("wpml:actionActuatorFunc", NS)
        if func is not None and func.text == "orientedShoot":
            param = a.find("wpml:actionActuatorFuncParam", NS)
            if param is None:
                param = etree.SubElement(a, "{http://www.dji.com/wpmz/1.0.6}actionActuatorFuncParam")
            pitch_el = param.find("wpml:gimbalPitchRotateAngle", NS)
            if pitch_el is None:
                pitch_el = etree.SubElement(param, "{http://www.dji.com/wpmz/1.0.6}gimbalPitchRotateAngle")
            pitch_el.text = str(angle)


def convert_kml(tree, params):
    offset = params.get("offset", 0.0)
    deviation = params.get("coordinate_deviation")  # None または (dlng, dlat, dalt)
    do_asl = params.get("do_asl", False)

    do_photo = params.get("do_photo", False)
    do_video = params.get("do_video", False)

    yaw_angle = params.get("yaw_angle")
    hover_time = params.get("hover_time", 0)
    sensor_modes = params.get("sensor_modes", [])

    # --- 新規追加パラメータ ---
    gimbal_pitch_ctrl = params.get("gimbal_pitch_ctrl", False)
    gimbal_pitch_mode = params.get("gimbal_pitch_mode")  # '-90' | '0' | 'manual' | 'free'
    gimbal_pitch_angle = params.get("gimbal_pitch_angle")  # float or None

    hmode_elem = tree.find("./kml:Document/kml:Folder/wpml:waylineCoordinateSysParam/wpml:heightMode", NS)
    height_mode = hmode_elem.text if hmode_elem is not None else None

    for pm in tree.findall(".//kml:Placemark", NS):
        # ------------------------------------------------------------------
        # 位置情報補正（既存機能）
        # ------------------------------------------------------------------
        coords_elem = pm.find(".//kml:coordinates", NS)
        if coords_elem is None or not coords_elem.text:
            continue
        coords_text = coords_elem.text.strip().split(",")
        try:
            if len(coords_text) == 2:
                height_el = pm.find("wpml:height", NS)
                alt = float(height_el.text) if height_el is not None else 0.0
                lng, lat = map(float, coords_text)
            elif len(coords_text) == 3:
                lng, lat, alt = map(float, coords_text)
            else:
                continue
        except (ValueError, IndexError):
            continue

        if deviation is not None:
            dlng, dlat, dalt = deviation
            lng += dlng
            lat += dlat
            alt += dalt

        if do_asl and height_mode == "relativeToStartPoint":
            alt += offset

        coords_elem.text = f"{lng},{lat},{alt}"

        for tag in ("height", "ellipsoidHeight"):
            el = pm.find(f"wpml:{tag}", NS)
            if el is not None:
                el.text = f"{alt}"

        # ------------------------------------------------------------------
        # ExtendedData 更新
        # ------------------------------------------------------------------
        ed = pm.find(".//kml:ExtendedData", NS)
        if ed is None:
            ed = etree.SubElement(pm, "{http://www.opengis.net/kml/2.2}ExtendedData")

        # 撮影モード
        mode = "photo" if do_photo else "video" if do_video else ""
        existing_mode = ed.find(f".//kml:Data[@name='mode']", NS)
        if mode:
            if existing_mode is not None:
                existing_mode.text = mode
            else:
                etree.SubElement(ed, "{http://www.opengis.net/kml/2.2}Data", name="mode").text = mode
        else:
            if existing_mode is not None:
                ed.remove(existing_mode)

        # ヨー角
        if yaw_angle is not None:
            etree.SubElement(ed, "{http://www.opengis.net/kml/2.2}Data", name="yaw").text = str(yaw_angle)

        # ホバリング
        if hover_time > 0:
            etree.SubElement(ed, "{http://www.opengis.net/kml/2.2}Data", name="hover_time").text = str(hover_time)

        # センサー選択（差分更新）
        sensors_data = ed.find(f".//kml:Data[@name='sensors']", NS)
        if sensor_modes:
            new_text = ",".join(sensor_modes)
            if sensors_data is not None:
                sensors_data.text = new_text
            else:
                etree.SubElement(ed, "{http://www.opengis.net/kml/2.2}Data", name="sensors").text = new_text
        # チェック無しの場合は既存値を保持（変更なし）

        # ジンバル制御フラグ（既存）
        etree.SubElement(ed, "{http://www.opengis.net/kml/2.2}Data", name="gimbal").text = str(
            params.get("do_gimbal", True)
        )

        # ------------------------------------------------------------------
        # ジンバルピッチ統一処理
        # ------------------------------------------------------------------
        if gimbal_pitch_ctrl and gimbal_pitch_mode != "free" and gimbal_pitch_angle is not None:
            ag = pm.find(".//wpml:actionGroup", NS)
            if ag is not None:
                # Pitch 統一
                unify_oriented_shoot_pitch(ag, gimbal_pitch_angle)
                # 撮影前 gimbalRotate 挿入
                insert_gimbal_rotate(ag, gimbal_pitch_angle)

    # globalHeight & heightMode 更新（既存機能）
    gh = tree.find(".//wpml:globalHeight", NS)
    if gh is not None:
        base = float(gh.text or 0)
        new = base
        if do_asl and height_mode == "relativeToStartPoint":
            new += offset
        if deviation is not None:
            new += deviation[2]
        gh.text = str(new)

    if do_asl and height_mode == "relativeToStartPoint":
        for hm in tree.findall(".//wpml:heightMode", NS):
            hm.text = "EGM96"


def process_kmz(path, log, params):
    try:
        log.insert(tk.END, f"処理開始: {os.path.basename(path)}...\n")
        work_dir = extract_kmz(path)

        kml_path = os.path.join(work_dir, "wpmz", "template.kml")
        if not os.path.exists(kml_path):
            kml_path = os.path.join(work_dir, "template.kml")
        if not os.path.exists(kml_path):
            raise FileNotFoundError("template.kml が見つかりませんでした。")

        out_root, outdir = prepare_output_dirs(path, params["offset"])
        log.insert(tk.END, f"変換中: {os.path.basename(kml_path)}...\n")

        parser = etree.XMLParser(remove_blank_text=True)
        tree = etree.parse(kml_path, parser)
        convert_kml(tree, params)

        out_kml = os.path.join(outdir, os.path.basename(kml_path))
        tree.write(out_kml, encoding="utf-8", pretty_print=True, xml_declaration=True)

        # リソースフォルダのコピー
        res_src = os.path.join(os.path.dirname(kml_path), "res")
        if os.path.isdir(res_src):
            shutil.copytree(res_src, os.path.join(outdir, "res"))

        out_kmz = repackage_to_kmz(out_root, path)
        log.insert(tk.END, f"変換完了: {out_kmz}\n\n")
        messagebox.showinfo("完了", f"変換が正常に完了しました:\n{out_kmz}")
    except Exception as e:
        log.insert(tk.END, f"エラー: {e}\n\n")
        messagebox.showerror("エラー", f"処理中にエラーが発生しました:\n{e}")
    finally:
        if os.path.exists("_kmz_work"):
            shutil.rmtree("_kmz_work")

# -------------------------------------------------------------------
# View
# -------------------------------------------------------------------
class AppUI(ttk.Frame):
    def __init__(self, master, controller):
        super().__init__(master, padding=10)
        self.controller = controller
        self._create_vars()
        self._create_widgets()
        self._grid_widgets()

    # 変数定義 --------------------------------------------------------
    def _create_vars(self):
        self.height_choice_var = tk.StringVar()
        self.height_entry_var = tk.StringVar()
        self.asl_var = tk.BooleanVar(value=False)

        self.speed_var = tk.IntVar(value=15)

        self.photo_var = tk.BooleanVar(value=False)
        self.video_var = tk.BooleanVar(value=False)
        self.video_suffix_var = tk.StringVar(value="video_01")

        self.sensor_vars = {m: tk.BooleanVar(value=False) for m in SENSOR_MODES}

        self.gimbal_var = tk.BooleanVar(value=True)

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

        # --- 新規: ジンバルピッチ制御 ---
        self.gim_pitch_ctrl_var = tk.BooleanVar(value=False)
        self.gim_pitch_choice_var = tk.StringVar()
        self.gim_pitch_entry_var = tk.StringVar()

    # ウィジェット作成 ------------------------------------------------
    def _create_widgets(self):
        # ASL
        self.asl_check = ttk.Checkbutton(
            self,
            text="ATL→ASL変換",
            variable=self.asl_var,
            command=self.controller.update_ui_states,
        )
        # 基準高度
        self.height_label = ttk.Label(self, text="基準高度:")
        self.height_combo = ttk.Combobox(
            self,
            textvariable=self.height_choice_var,
            values=list(HEIGHT_OPTIONS.keys()),
            state="readonly",
            width=20,
        )
        self.height_combo.bind("<<ComboboxSelected>>", self.controller.update_ui_states)
        self.height_entry = ttk.Entry(self, textvariable=self.height_entry_var, width=10)

        # 速度
        self.speed_label = ttk.Label(self, text="速度 (1–15 m/s):")
        self.speed_spinbox = ttk.Spinbox(self, from_=1, to=15, textvariable=self.speed_var, width=5)

        # 撮影設定
        self.photo_check = ttk.Checkbutton(
            self, text="写真撮影", variable=self.photo_var, command=self.controller.update_ui_states
        )
        self.video_check = ttk.Checkbutton(
            self, text="動画撮影", variable=self.video_var, command=self.controller.update_ui_states
        )
        self.video_suffix_label = ttk.Label(self, text="動画ファイル名:")
        self.video_suffix_entry = ttk.Entry(self, textvariable=self.video_suffix_var, width=20)

        # センサー
        self.sensor_label = ttk.Label(self, text="センサー選択:")
        self.sensor_checks = {m: ttk.Checkbutton(self, text=m, variable=v) for m, v in self.sensor_vars.items()}

        # ジンバル有無
        self.gimbal_check = ttk.Checkbutton(self, text="ジンバル制御", variable=self.gimbal_var)

        # ジンバルピッチ制御（新規）
        self.gim_pitch_ctrl_check = ttk.Checkbutton(
            self, text="ジンバルピッチ制御", variable=self.gim_pitch_ctrl_var, command=self.controller.update_ui_states
        )
        self.gim_pitch_combo = ttk.Combobox(
            self,
            textvariable=self.gim_pitch_choice_var,
            values=["真下 (-90°)", "前 (0°)", "手動入力", "フリー"],
            state="readonly",
            width=15,
        )
        self.gim_pitch_combo.bind("<<ComboboxSelected>>", self.controller.update_ui_states)
        self.gim_pitch_entry = ttk.Entry(self, textvariable=self.gim_pitch_entry_var, width=8)

        # ヨー固定
        self.yaw_fix_check = ttk.Checkbutton(
            self, text="ヨー固定", variable=self.yaw_fix_var, command=self.controller.update_ui_states
        )
        self.yaw_combo = ttk.Combobox(
            self, textvariable=self.yaw_choice_var, values=list(YAW_OPTIONS.keys()), state="readonly", width=15
        )
        self.yaw_combo.bind("<<ComboboxSelected>>", self.controller.update_ui_states)
        self.yaw_entry = ttk.Entry(self, textvariable=self.yaw_entry_var, width=8)

        # ホバリング
        self.hover_check = ttk.Checkbutton(
            self, text="ホバリング", variable=self.hover_var, command=self.controller.update_ui_states
        )
        self.hover_time_label = ttk.Label(self, text="ホバリング時間 (秒):")
        self.hover_time_entry = ttk.Entry(self, textvariable=self.hover_time_var, width=8)

        # 偏差補正
        self.deviation_check = ttk.Checkbutton(
            self, text="偏差補正", variable=self.deviation_var, command=self.controller.update_ui_states
        )
        self.ref_point_label = ttk.Label(self, text="基準位置:")
        self.ref_point_combo = ttk.Combobox(
            self, textvariable=self.ref_point_var, values=list(REFERENCE_POINTS.keys()), state="readonly", width=10
        )
        self.today_coords_label = ttk.Label(self, text="本日の値 (経度,緯度,標高):")
        self.today_lng_entry = ttk.Entry(self, textvariable=self.today_lng_var, width=12)
        self.today_lat_entry = ttk.Entry(self, textvariable=self.today_lat_var, width=10)
        self.today_alt_entry = ttk.Entry(self, textvariable=self.today_alt_var, width=8)
        self.copy_button = ttk.Button(self, text="コピー", command=self.controller.copy_reference_data, width=8)

    # レイアウト ------------------------------------------------------
    def _grid_widgets(self):
        # 1行目
        self.asl_check.grid(row=0, column=0, sticky="w", pady=5)

        # 速度
        self.speed_label.grid(row=1, column=0, sticky="w", pady=5)
        self.speed_spinbox.grid(row=1, column=1, columnspan=2, sticky="w")

        # 撮影
        self.photo_check.grid(row=2, column=0, sticky="w")
        self.video_check.grid(row=2, column=1, sticky="w")

        # センサー
        self.sensor_label.grid(row=3, column=0, sticky="w")
        for i, m in enumerate(SENSOR_MODES):
            self.sensor_checks[m].grid(row=3, column=1 + i, sticky="w")

        # ジンバル制御
        self.gimbal_check.grid(row=4, column=0, sticky="w", pady=5)

        # ジンバルピッチ制御
        self.gim_pitch_ctrl_check.grid(row=5, column=0, sticky="w")


# -------------------------------------------------------------------
# Controller
# -------------------------------------------------------------------
class KmlConverterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("ATL→ASL 変換＋撮影制御ツール (Enhanced)")
        self.root.geometry("820x780")
        self.ui = AppUI(self.root, self)
        self.ui.pack(fill="x", pady=(0, 10))
        self._create_dnd_and_log_area()
        self.update_ui_states()

    def _create_dnd_and_log_area(self):
        frm = ttk.Frame(self.root, padding=10)
        frm.pack(fill="both", expand=True)

        self.drop_label = tk.Label(frm, text=".kmz をここにドロップ", bg="lightgray", width=70, height=5, relief=tk.RIDGE)
        self.drop_label.pack(pady=12, fill="x")
        self.drop_label.drop_target_register(DND_FILES)
        self.drop_label.dnd_bind("<<Drop>>", self.on_drop)

        log_frame = ttk.LabelFrame(frm, text="ログ")
        log_frame.pack(fill="both", expand=True)
        self.log_text = scrolledtext.ScrolledText(log_frame, height=18)
        self.log_text.pack(fill="both", expand=True)

    # --------------------------------------------------------------
    # UI ロジック更新
    # --------------------------------------------------------------
    def update_ui_states(self, event=None):
        ui = self.ui

        # --- ASL ---
        if ui.asl_var.get():
            ui.height_label.grid(row=0, column=2, sticky="w")
            ui.height_combo.grid(row=0, column=3, padx=5, columnspan=2, sticky="w")

            if HEIGHT_OPTIONS.get(ui.height_choice_var.get()) == "custom":
                ui.height_entry.config(state="normal")
                ui.height_entry.grid(row=0, column=5, padx=5)
            else:
                ui.height_entry.config(state="disabled")
                ui.height_entry.grid_forget()
        else:
            ui.height_label.grid_forget()
            ui.height_combo.grid_forget()
            ui.height_entry.grid_forget()

        # 撮影モード排他
        if ui.photo_var.get():
            ui.video_var.set(False)
        if ui.video_var.get():
            ui.photo_var.set(False)

        # ジンバル制御チェック
        ui.gimbal_check.config(state="disabled" if ui.photo_var.get() or ui.video_var.get() else "normal")
        if ui.photo_var.get() or ui.video_var.get():
            ui.gimbal_var.set(True)

        # 動画ファイル名
        if ui.video_var.get():
            ui.video_suffix_label.grid(row=2, column=2, sticky="e", padx=(10, 2))
            ui.video_suffix_entry.grid(row=2, column=3, sticky="w")
        else:
            ui.video_suffix_label.grid_forget()
            ui.video_suffix_entry.grid_forget()

        # ジンバルピッチ制御
        if ui.gim_pitch_ctrl_var.get():
            ui.gim_pitch_combo.grid(row=5, column=1, padx=5, sticky="w")
            # デフォルト値
            if not ui.gim_pitch_choice_var.get():
                ui.gim_pitch_choice_var.set("真下 (-90°)")

            if ui.gim_pitch_choice_var.get() == "手動入力":
                ui.gim_pitch_entry.config(state="normal")
                ui.gim_pitch_entry.grid(row=5, column=2)
            else:
                ui.gim_pitch_entry.config(state="disabled")
                ui.gim_pitch_entry.grid_forget()
        else:
            ui.gim_pitch_combo.grid_forget()
            ui.gim_pitch_entry.grid_forget()

        # ヨー固定
        if ui.yaw_fix_var.get():
            ui.yaw_combo.grid(row=6, column=1, padx=5, columnspan=2, sticky="w")
            if not ui.yaw_choice_var.get():
                ui.yaw_choice_var.set(next(iter(YAW_OPTIONS.keys())))
            if ui.yaw_choice_var.get() == "手動入力":
                ui.yaw_entry.config(state="normal")
                ui.yaw_entry.grid(row=6, column=3)
            else:
                ui.yaw_entry.config(state="disabled")
                ui.yaw_entry.grid_forget()
        else:
            ui.yaw_combo.grid_forget()
            ui.yaw_entry.grid_forget()

        # ホバリング
        if ui.hover_var.get():
            ui.hover_time_label.grid(row=7, column=1, sticky="e", padx=(10, 2))
            ui.hover_time_entry.grid(row=7, column=2, sticky="w")
        else:
            ui.hover_time_label.grid_forget()
            ui.hover_time_entry.grid_forget()

        # 偏差補正
        if ui.deviation_var.get():
            ui.ref_point_label.grid(row=8, column=1, sticky="e", padx=(10, 2))
            ui.ref_point_combo.grid(row=8, column=2, sticky="w")
            ui.today_coords_label.grid(row=9, column=0, sticky="w", pady=5)
            ui.today_lng_entry.grid(row=9, column=1, padx=2, sticky="w")
            ui.today_lat_entry.grid(row=9, column=2, padx=2, sticky="w")
            ui.today_alt_entry.grid(row=9, column=3, padx=2, sticky="w")
            ui.copy_button.grid(row=9, column=4, padx=5, sticky="w")
        else:
            ui.ref_point_label.grid_forget()
            ui.ref_point_combo.grid_forget()
            ui.today_coords_label.grid_forget()
            ui.today_lng_entry.grid_forget()
            ui.today_lat_entry.grid_forget()
            ui.today_alt_entry.grid_forget()
            ui.copy_button.grid_forget()

    # --------------------------------------------------------------
    # ファイルドロップ
    # --------------------------------------------------------------
    def on_drop(self, event):
        path = event.data.strip("{}")
        if not path.lower().endswith(".kmz"):
            messagebox.showwarning("警告", ".kmz ファイルのみ対応しています。")
            return

        try:
            params = self._get_params()
        except ValueError as e:
            messagebox.showerror("入力エラー", str(e))
            return

        threading.Thread(target=process_kmz, args=(path, self.log_text, params), daemon=True).start()

    def copy_reference_data(self):
        try:
            ref = self.ui.ref_point_var.get()
            coords = REFERENCE_POINTS[ref]
            now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            txt = f"{now},{coords[0]},{coords[1]},{coords[2]}"
            pyperclip.copy(txt)
            messagebox.showinfo("コピー完了", f"クリップボードにコピーしました:\n{txt}")
        except Exception as e:
            messagebox.showerror("エラー", f"コピーに失敗しました: {e}")

    # --------------------------------------------------------------
    # パラメータ取得
    # --------------------------------------------------------------
    def _get_params(self):
        ui = self.ui
        p = {}

        # --- ASL ---
        if ui.asl_var.get():
            v = HEIGHT_OPTIONS.get(ui.height_choice_var.get())
            if v == "custom":
                try:
                    p["offset"] = float(ui.height_entry_var.get())
                except ValueError:
                    raise ValueError("基準高度（手動入力）には数値を入力してください。")
            else:
                p["offset"] = float(v)
        else:
            p["offset"] = 0.0

        p["do_asl"] = ui.asl_var.get()

        # --- ジンバルピッチ制御 ---
        p["gimbal_pitch_ctrl"] = ui.gim_pitch_ctrl_var.get()
        if ui.gim_pitch_ctrl_var.get():
            choice = ui.gim_pitch_choice_var.get()
            if choice.startswith("真下"):
                p["gimbal_pitch_mode"] = "-90"
                p["gimbal_pitch_angle"] = -90.0
            elif choice.startswith("前"):
                p["gimbal_pitch_mode"] = "0"
                p["gimbal_pitch_angle"] = 0.0
            elif choice == "手動入力":
                p["gimbal_pitch_mode"] = "manual"
                try:
                    p["gimbal_pitch_angle"] = float(ui.gim_pitch_entry_var.get())
                except ValueError:
                    raise ValueError("ジンバルピッチ角（手動入力）には数値を入力してください。")
            else:  # フリー
                p["gimbal_pitch_mode"] = "free"
                p["gimbal_pitch_angle"] = None
        else:
            p["gimbal_pitch_ctrl"] = False
            p["gimbal_pitch_mode"] = "free"
            p["gimbal_pitch_angle"] = None

        # --- ヨー固定 ---
        p["yaw_angle"] = None
        if ui.yaw_fix_var.get():
            y = YAW_OPTIONS.get(ui.yaw_choice_var.get())
            if y == "custom":
                try:
                    p["yaw_angle"] = float(ui.yaw_entry_var.get())
                except ValueError:
                    raise ValueError("ヨー角（手動入力）には数値を入力してください。")
            else:
                p["yaw_angle"] = float(y)

        # --- ホバリング ---
        p["hover_time"] = 0
        if ui.hover_var.get():
            try:
                p["hover_time"] = max(0, float(ui.hover_time_var.get()))
            except ValueError:
                p["hover_time"] = 2.0

        # --- 偏差補正 ---
        p["coordinate_deviation"] = None
        if ui.deviation_var.get():
            try:
                ref = REFERENCE_POINTS[ui.ref_point_var.get()]
                cur = (
                    float(ui.today_lng_var.get()),
                    float(ui.today_lat_var.get()),
                    float(ui.today_alt_var.get()),
                )
                dev = calculate_deviation(ref, cur)
                ok, msg = check_deviation_safety(dev)
                if not ok:
                    raise ValueError(msg)
                p["coordinate_deviation"] = dev
            except ValueError as e:
                raise ValueError(f"偏差補正の値が不正です: {e}")

        # --- その他 ---
        p["do_photo"] = ui.photo_var.get()
        p["do_video"] = ui.video_var.get()
        p["video_suffix"] = ui.video_suffix_var.get()
        p["do_gimbal"] = ui.gimbal_var.get()
        p["speed"] = max(1, min(15, ui.speed_var.get()))
        p["sensor_modes"] = [m for m, var in ui.sensor_vars.items() if var.get()]

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
