#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
GUI42.py  –  ATL→ASL 変換＋撮影制御ツール（imageFormat 方式）
2025-07-31
----------------------------------------------------------------
・wpml:payloadParam/imageFormat=zoom,wide 形式でカメラ選択
・Zoom / Wide / IR を複数チェックボックスで選択可
・写真＝orientedShoot、動画＝startRecord/stopRecord
・最初 WP で gimbal → recordStart、最後 WP で recordStop
・ヨー固定は rotateYaw を継続使用
----------------------------------------------------------------
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
NS = {"kml": "http://www.opengis.net/kml/2.2",
      "wpml": "http://www.dji.com/wpmz/1.0.6"}

HEIGHT_OPTIONS = {"613.5 – 事務所前": 613.5,
                  "962.02 – 烏帽子": 962.02,
                  "その他 – 手動入力": "custom"}

YAW_OPTIONS = {"1Q: 88.00°": 88.00,
               "2Q: 96.92°": 96.92,
               "4Q: 87.31°": 87.31,
               "手動入力": "custom"}

IMAGE_FORMAT_CHOICES = ["zoom", "wide", "ir"]

REFERENCE_POINTS = {"本部": (136.55595225, 36.07295176, 612.2),
                    "烏帽子": (136.56000000, 36.07500000, 962.02)}

DEVIATION_THRESHOLD = {"lat": 0.00018, "lng": 0.00022, "alt": 20.0}

# -------------------------------------------------------------------
# ユーティリティ
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

def calculate_deviation(ref, cur):
    rlng, rlat, ralt = ref
    clng, clat, calt = cur
    return clng - rlng, clat - rlat, calt - ralt

def check_deviation_safety(dev):
    dlng, dlat, dalt = dev
    if (abs(dlng) > DEVIATION_THRESHOLD["lng"] or
        abs(dlat) > DEVIATION_THRESHOLD["lat"] or
        abs(dalt) > DEVIATION_THRESHOLD["alt"]):
        return False, (f"偏差が閾値を超えています:\n"
                       f"経度偏差: {dlng:.8f}°\n"
                       f"緯度偏差: {dlat:.8f}°\n"
                       f"標高偏差: {dalt:.2f}m")
    return True, None

# -------------------------------------------------------------------
# アクション生成
# -------------------------------------------------------------------
def new_action(func_name):
    act = etree.Element("{wpml}action".format(**NS))
    etree.SubElement(act, "{wpml}actionId".format(**NS)).text = "0"
    etree.SubElement(act, "{wpml}actionActuatorFunc".format(**NS)).text = func_name
    param = etree.SubElement(act, "{wpml}actionActuatorFuncParam".format(**NS))
    return act, param

def create_rotate_yaw(yaw):
    act, param = new_action("rotateYaw")
    etree.SubElement(param, "{wpml}aircraftHeading".format(**NS)).text = str(yaw)
    etree.SubElement(param, "{wpml}aircraftPathMode".format(**NS)).text = "counterClockwise"
    return act

def create_gimbal_rotate(angle):
    act, param = new_action("gimbalRotate")
    etree.SubElement(param, "{wpml}gimbalRotateMode".format(**NS)).text = "absoluteAngle"
    etree.SubElement(param, "{wpml}gimbalPitchRotateEnable".format(**NS)).text = "1"
    etree.SubElement(param, "{wpml}gimbalPitchRotateAngle".format(**NS)).text = str(angle)
    etree.SubElement(param, "{wpml}gimbalRollRotateEnable".format(**NS)).text = "0"
    etree.SubElement(param, "{wpml}gimbalYawRotateEnable".format(**NS)).text = "0"
    return act

def create_start_record():
    act, _ = new_action("startRecord")
    return act

def create_stop_record():
    act, _ = new_action("stopRecord")
    return act

# -------------------------------------------------------------------
# KML 操作
# -------------------------------------------------------------------
def update_payload_image_format(tree, imgfmt):
    elem = tree.find(".//wpml:payloadParam/wpml:imageFormat", NS)
    if elem is None:
        pp = tree.find(".//wpml:payloadParam", NS)
        elem = etree.SubElement(pp, "{wpml}imageFormat".format(**NS))
    elem.text = imgfmt

def organised_actions(group, idx, first_idx, last_idx, params, log):
    # 既存アクション抽出
    acts = list(group.findall("wpml:action", NS))
    # 全削除
    for a in acts:
        group.remove(a)

    is_photo = params["mode"] == "photo"
    seq = []

    # ヨー固定
    if params["yaw_angle"] is not None:
        seq.append(create_rotate_yaw(params["yaw_angle"]))
    # ジンバル
    if params["gimbal_pitch_enable"] and params["gimbal_pitch_angle"] is not None:
        seq.append(create_gimbal_rotate(params["gimbal_pitch_angle"]))

    if is_photo:
        # 既存 orientedShoot を再利用（写真時）
        for a in acts:
            fn = a.find("wpml:actionActuatorFunc", NS).text
            if fn == "orientedShoot":
                # ジンバル角統一
                if params["gimbal_pitch_enable"]:
                    p = a.find("wpml:actionActuatorFuncParam", NS)
                    p.find("wpml:gimbalPitchRotateAngle", NS).text = str(params["gimbal_pitch_angle"])
                seq.append(a)
    else:
        # 動画時
        if idx == first_idx:
            seq.append(create_start_record())
        if idx == last_idx:
            seq.append(create_stop_record())

    # 再配置 ＆ actionId 再振り
    for i, a in enumerate(seq):
        a.find("wpml:actionId", NS).text = str(i)
        group.append(a)

    log.append(f"WP{idx}: " + " → ".join(a.find("wpml:actionActuatorFunc", NS).text for a in seq))

def convert_kml(tree, params, log_buf):
    placemarks = tree.findall(".//kml:Placemark", NS)
    first_idx = int(placemarks[0].find("wpml:index", NS).text)
    last_idx  = int(placemarks[-1].find("wpml:index", NS).text)

    # imageFormat 更新
    update_payload_image_format(tree, params["image_format"])

    for pm in placemarks:
        idx = int(pm.find("wpml:index", NS).text)
        group = pm.find(".//wpml:actionGroup", NS)
        if group is not None:
            organised_actions(group, idx, first_idx, last_idx, params, log_buf)

# -------------------------------------------------------------------
# UI
# -------------------------------------------------------------------
class AppUI(ttk.Frame):
    def __init__(self, master, controller):
        super().__init__(master, padding=10)
        self.controller = controller
        self._vars()
        self._widgets()
        self._layout()

    def _vars(self):
        self.photo_var = tk.BooleanVar()
        self.video_var = tk.BooleanVar()
        # imageFormat チェックボックス
        self.img_vars = {k: tk.BooleanVar() for k in IMAGE_FORMAT_CHOICES}

        self.gimbal_pitch_enable = tk.BooleanVar()
        self.pitch_choice = tk.StringVar(value="真下 (-90°)")
        self.pitch_entry  = tk.StringVar()

        self.yaw_enable = tk.BooleanVar()
        self.yaw_choice = tk.StringVar(value="1Q: 88.00°")
        self.yaw_entry  = tk.StringVar()

    def _widgets(self):
        self.photo_chk = ttk.Checkbutton(self, text="写真", variable=self.photo_var,
                                         command=self.controller.sync_mode)
        self.video_chk = ttk.Checkbutton(self, text="動画", variable=self.video_var,
                                         command=self.controller.sync_mode)

        self.img_checks = [ttk.Checkbutton(self, text=txt.upper(), variable=self.img_vars[txt])
                           for txt in IMAGE_FORMAT_CHOICES]

        self.gim_chk = ttk.Checkbutton(self, text="ジンバルピッチ制御",
                                       variable=self.gimbal_pitch_enable,
                                       command=self.controller.update_ui)

        self.pitch_cmb = ttk.Combobox(self, textvariable=self.pitch_choice,
                                      values=["真下 (-90°)", "前 (0°)", "手動入力"], width=12,
                                      state="readonly")
        self.pitch_cmb.bind("<<ComboboxSelected>>", self.controller.update_ui)
        self.pitch_ent = ttk.Entry(self, textvariable=self.pitch_entry, width=6)

        self.yaw_chk = ttk.Checkbutton(self, text="ヨー固定",
                                       variable=self.yaw_enable,
                                       command=self.controller.update_ui)
        self.yaw_cmb = ttk.Combobox(self, textvariable=self.yaw_choice,
                                    values=list(YAW_OPTIONS.keys()), width=12,
                                    state="readonly")
        self.yaw_cmb.bind("<<ComboboxSelected>>", self.controller.update_ui)
        self.yaw_ent = ttk.Entry(self, textvariable=self.yaw_entry, width=6)

    def _layout(self):
        self.photo_chk.grid(row=0, column=0, sticky="w", padx=2)
        self.video_chk.grid(row=0, column=1, sticky="w", padx=2)
        ttk.Label(self, text="imageFormat:").grid(row=1, column=0, sticky="w")
        for col, chk in enumerate(self.img_checks, start=1):
            chk.grid(row=1, column=col, sticky="w", padx=2)

        self.gim_chk.grid(row=2, column=0, sticky="w", pady=(8,0))
        self.pitch_cmb.grid(row=2, column=1, sticky="w", padx=4)
        self.pitch_ent.grid(row=2, column=2, sticky="w")

        self.yaw_chk.grid(row=3, column=0, sticky="w", pady=(8,0))
        self.yaw_cmb.grid(row=3, column=1, sticky="w", padx=4)
        self.yaw_ent.grid(row=3, column=2, sticky="w")

# -------------------------------------------------------------------
# Controller
# -------------------------------------------------------------------
class AppCtl:
    def __init__(self, root):
        self.root = root
        self.ui   = AppUI(root, self)
        self.ui.pack(anchor="w")
        self.log = scrolledtext.ScrolledText(root, height=20)
        self.log.pack(fill="both", expand=True, pady=6)
        self.drop = tk.Label(root, text=".kmz をドロップ", bg="lightgray",
                             width=60, height=3, relief=tk.RIDGE)
        self.drop.pack(fill="x", pady=4)
        self.drop.drop_target_register(DND_FILES)
        self.drop.dnd_bind("<<Drop>>", self.on_drop)
        self.update_ui()

    # --- UI sync ---------------------------------------------------
    def sync_mode(self):
        # 排他
        if self.ui.photo_var.get():
            self.ui.video_var.set(False)
        elif self.ui.video_var.get():
            self.ui.photo_var.set(False)
        self.update_ui()

    def update_ui(self, *_):
        # pitch 手動入力表示
        pc = self.ui.pitch_choice.get()
        if self.ui.gimbal_pitch_enable.get() and pc == "手動入力":
            self.ui.pitch_ent.config(state="normal")
        else:
            self.ui.pitch_ent.config(state="disabled")
        # yaw 手動
        yc = self.ui.yaw_choice.get()
        if self.ui.yaw_enable.get() and yc == "手動入力":
            self.ui.yaw_ent.config(state="normal")
        else:
            self.ui.yaw_ent.config(state="disabled")

    # --- Drop ------------------------------------------------------
    def on_drop(self, ev):
        path = ev.data.strip("{}")
        if not path.lower().endswith(".kmz"):
            messagebox.showwarning("注意", ".kmz のみサポート")
            return
        try:
            params = self.collect_params()
        except ValueError as e:
            messagebox.showerror("入力エラー", str(e))
            return
        threading.Thread(target=self.process, args=(path, params), daemon=True).start()

    # --- Param collection -----------------------------------------
    def collect_params(self):
        p = {}
        # mode
        if self.ui.photo_var.get():
            p["mode"] = "photo"
        elif self.ui.video_var.get():
            p["mode"] = "video"
        else:
            raise ValueError("写真か動画のいずれかを選択してください。")

        # imageFormat
        sel = [k for k,v in self.ui.img_vars.items() if v.get()]
        if not sel:
            raise ValueError("少なくとも1つの imageFormat を選択してください。")
        p["image_format"] = ",".join(sel)

        # gimbal
        p["gimbal_pitch_enable"] = self.ui.gimbal_pitch_enable.get()
        if p["gimbal_pitch_enable"]:
            if self.ui.pitch_choice.get().startswith("真下"):
                p["gimbal_pitch_angle"] = -90.0
            elif self.ui.pitch_choice.get().startswith("前"):
                p["gimbal_pitch_angle"] = 0.0
            elif self.ui.pitch_choice.get() == "手動入力":
                try:
                    p["gimbal_pitch_angle"] = float(self.ui.pitch_entry.get())
                except ValueError:
                    raise ValueError("ピッチ角に数値を入力してください。")
        else:
            p["gimbal_pitch_angle"] = None

        # yaw
        p["yaw_angle"] = None
        if self.ui.yaw_enable.get():
            yc = self.ui.yaw_choice.get()
            if YAW_OPTIONS.get(yc) == "custom":
                try:
                    p["yaw_angle"] = float(self.ui.yaw_entry.get())
                except ValueError:
                    raise ValueError("ヨー角に数値を入力してください。")
            else:
                p["yaw_angle"] = float(YAW_OPTIONS[yc])
        return p

    # --- Core processing ------------------------------------------
    def process(self, kmz_path, params):
        self.log.insert(tk.END, f"処理開始: {os.path.basename(kmz_path)}\n")
        work = extract_kmz(kmz_path)
        kml_p = os.path.join(work, "wpmz", "template.kml")
        if not os.path.exists(kml_p):
            kml_p = os.path.join(work, "template.kml")
        tree = etree.parse(kml_p)

        buf = []
        convert_kml(tree, params, buf)
        for line in buf:
            self.log.insert(tk.END, line + "\n")

        out_root, wpmz_dir = prepare_output_dirs(kmz_path, 0.0)
        out_kml = os.path.join(wpmz_dir, "template.kml")
        tree.write(out_kml, encoding="utf-8", pretty_print=True, xml_declaration=True)

        # res フォルダコピー
        res_src = os.path.join(os.path.dirname(kml_p), "res")
        if os.path.isdir(res_src):
            shutil.copytree(res_src, os.path.join(wpmz_dir, "res"))

        kmz_out = repackage_to_kmz(out_root, kmz_path)
        self.log.insert(tk.END, f"出力: {kmz_out}\n\n")
        messagebox.showinfo("完了", f"変換完了:\n{kmz_out}")
        shutil.rmtree("_kmz_work")

# -------------------------------------------------------------------
def main():
    root = TkinterDnD.Tk()
    root.title("ATL→ASL 変換＋撮影制御ツール v4 (imageFormat)")
    root.geometry("780x700")
    AppCtl(root)
    root.mainloop()

if __name__ == "__main__":
    main()
