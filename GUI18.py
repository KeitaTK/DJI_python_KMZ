#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
convert_height_gui_asl.py (最終版)
---------------------------------
KMZ 内の WPML:
・高度をオフセット付きで変換 (relativeToStartPoint → EGM96)
・速度をグローバル／各ウェイポイントともに個別指定
・ウェイポイント速度範囲1–15 m/s
・写真撮影オプション
・撮影オフ時のみジンバル制御オン／オフ
・撮影オン時はジンバル制御有効
・ヨー角固定オプション（既定 or 手動入力）
依存ライブラリ:
pip install lxml tkinterdnd2
"""

import os
import shutil
import zipfile
import glob
import threading
import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext
from tkinterdnd2 import TkinterDnD, DND_FILES
from lxml import etree

# WPML 名前空間
NS = {
    "kml": "http://www.opengis.net/kml/2.2",
    "wpml": "http://www.dji.com/wpmz/1.0.6"
}

# GUI 用定数
HEIGHT_OPTIONS = {
    "613.5 – 事務所前": 613.5,
    "962.02 – 烏帽子": 962.02,
    "その他 – 手動入力": "custom"
}
YAW_FIX_OPTIONS = {
    "1Q: 87.37°": 87.37,
    "2Q: 96.92°": 96.92,
    "4Q: 87.31°": 87.31,
    "その他 – 手動入力": "custom"
}

def extract_kmz(kmz_path: str, work_dir: str) -> None:
    if os.path.exists(work_dir):
        shutil.rmtree(work_dir)
    os.makedirs(work_dir)
    with zipfile.ZipFile(kmz_path, "r") as zf:
        zf.extractall(work_dir)

def prepare_output_dirs(input_kmz: str, offset: float):
    base = os.path.splitext(os.path.basename(input_kmz))[0]
    sign = "+" if offset >= 0 else "-"
    out_base = f"{base}_asl{sign}{abs(offset)}"
    inp_dir = os.path.dirname(input_kmz)
    base_out = os.path.join(inp_dir, out_base)
    if os.path.exists(base_out):
        shutil.rmtree(base_out)
    os.makedirs(base_out)
    wpmz_dir = os.path.join(base_out, "wpmz")
    os.makedirs(wpmz_dir)
    return base_out, wpmz_dir

def repackage_to_kmz(base_out: str, input_kmz: str) -> str:
    inp_dir = os.path.dirname(input_kmz)
    name = os.path.splitext(os.path.basename(input_kmz))[0] + "_Converted.kmz"
    out_kmz = os.path.join(inp_dir, name)
    tmp = out_kmz + ".zip"
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zf:
        for root, _, files in os.walk(base_out):
            for f in files:
                full = os.path.join(root, f)
                rel = os.path.relpath(full, base_out)
                zf.write(full, rel)
    if os.path.exists(out_kmz):
        os.remove(out_kmz)
    os.rename(tmp, out_kmz)
    return out_kmz

def convert_heights_and_mode(tree: etree._ElementTree,
                             offset: float,
                             do_capture: bool,
                             do_gimbal: bool,
                             fix_yaw: bool,
                             yaw_angle: float,
                             speed: int) -> None:
    # 高度オフセット + heightMode→EGM96
    for pm in tree.findall(".//kml:Placemark", NS):
        for tag in ("height", "ellipsoidHeight"):
            e = pm.find(f"wpml:{tag}", NS)
            if e is not None:
                try:
                    e.text = str(float(e.text) + offset)
                except:
                    pass
    gh = tree.find(".//wpml:globalHeight", NS)
    if gh is not None:
        try:
            gh.text = str(float(gh.text) + offset)
        except:
            pass
    for hm in tree.findall(".//wpml:heightMode", NS):
        hm.text = "EGM96"

    # 速度設定：missionConfig/globalTransitionalSpeed
    gts = tree.find(".//wpml:missionConfig/wpml:globalTransitionalSpeed", NS)
    if gts is not None:
        gts.text = str(speed)
    # Folder/autoFlightSpeed
    afs = tree.find(".//kml:Folder/wpml:autoFlightSpeed", NS)
    if afs is not None:
        afs.text = str(speed)
    else:
        folder = tree.find(".//kml:Folder", NS)
        if folder is not None:
            new_afs = etree.SubElement(folder, f"{{{NS['wpml']}}}autoFlightSpeed")
            new_afs.text = str(speed)

    # 各ウェイポイントの速度 + useGlobalSpeed=0
    for pm in tree.findall(".//kml:Placemark", NS):
        ws = pm.find("wpml:waypointSpeed", NS)
        if ws is not None:
            ws.text = str(speed)
        else:
            new_ws = etree.SubElement(pm, f"{{{NS['wpml']}}}waypointSpeed")
            new_ws.text = str(speed)
        ugs = pm.find("wpml:useGlobalSpeed", NS)
        if ugs is not None:
            ugs.text = "0"
        else:
            new_ugs = etree.SubElement(pm, f"{{{NS['wpml']}}}useGlobalSpeed")
            new_ugs.text = "0"

    # アクション処理：orientedShoot/gimbalRotate 削除
    for ag in tree.findall(".//wpml:actionGroup", NS):
        for act in ag.findall("wpml:action", NS):
            f = act.find("wpml:actionActuatorFunc", NS)
            if f is not None and f.text == "orientedShoot" and not do_capture:
                ag.remove(act)
            if f is not None and f.text == "gimbalRotate" and not do_capture and not do_gimbal:
                ag.remove(act)

    # ヨー角固定処理
    if fix_yaw:
        gw = tree.find(".//wpml:globalWaypointHeadingParam", NS)
        if gw is not None:
            gw.find("wpml:waypointHeadingMode", NS).text = "fixed"
            gw.find("wpml:waypointHeadingAngle", NS).text = str(yaw_angle)
        for wp in tree.findall(".//wpml:waypointHeadingParam", NS):
            wp.find("wpml:waypointHeadingMode", NS).text = "fixed"
            wp.find("wpml:waypointHeadingAngle", NS).text = str(yaw_angle)

    # 空の actionGroup を削除
    for ag in tree.findall(".//wpml:actionGroup", NS):
        if not ag.findall("wpml:action", NS):
            parent = ag.getparent()
            parent.remove(ag)

def process_kmz(input_kmz: str,
                offset: float,
                do_capture: bool,
                do_gimbal: bool,
                fix_yaw: bool,
                yaw_angle: float,
                speed: int,
                log: tk.Text) -> None:
    work_dir = "_kmz_work"
    try:
        log.insert(tk.END, f"Extracting {os.path.basename(input_kmz)}...\n")
        extract_kmz(input_kmz, work_dir)

        kmls = glob.glob(os.path.join(work_dir, "**", "template.kml"), recursive=True)
        if not kmls:
            raise FileNotFoundError("template.kml が見つかりません")

        base_out, wpmz_dir = prepare_output_dirs(input_kmz, offset)
        for kml in kmls:
            log.insert(tk.END, f"Converting {os.path.basename(kml)}...\n")
            tree = etree.parse(kml)
            convert_heights_and_mode(tree, offset,
                                     do_capture, do_gimbal,
                                     fix_yaw, yaw_angle,
                                     speed)
            out_kml = os.path.join(wpmz_dir, os.path.basename(kml))
            tree.write(out_kml, encoding="utf-8", pretty_print=True, xml_declaration=True)

        # res フォルダコピー
        for src in glob.glob(os.path.join(work_dir, "**", "res"), recursive=True):
            dst = os.path.join(wpmz_dir, "res")
            if os.path.exists(dst):
                shutil.rmtree(dst)
            shutil.copytree(src, dst)
            log.insert(tk.END, "Copied res folder\n")

        # waylines.wpml コピー
        for src in glob.glob(os.path.join(work_dir, "**", "waylines.wpml"), recursive=True):
            dst = os.path.join(wpmz_dir, "waylines.wpml")
            shutil.copy2(src, dst)
            log.insert(tk.END, "Copied waylines.wpml\n")

        out_kmz = repackage_to_kmz(base_out, input_kmz)
        log.insert(tk.END, f"Saved: {out_kmz}\nFinished\n\n")
        messagebox.showinfo("完了", f"変換完了:\n{out_kmz}")

    except Exception as e:
        messagebox.showerror("エラー", str(e))
        log.insert(tk.END, f"Error: {e}\n\n")
    finally:
        if os.path.exists(work_dir):
            shutil.rmtree(work_dir)

class HeightSelector(ttk.Frame):
    def __init__(self, master):
        super().__init__(master)
        # 基準高度オフセット
        ttk.Label(self, text="基準高度オフセット:").grid(row=0, column=0, sticky="w")
        self.combo = ttk.Combobox(self, values=list(HEIGHT_OPTIONS.keys()), state="readonly", width=20)
        self.combo.set(next(iter(HEIGHT_OPTIONS)))
        self.combo.grid(row=0, column=1, padx=5)
        self.entry = ttk.Entry(self, width=10, state="disabled")
        self.entry.grid(row=0, column=2, padx=5)
        self.combo.bind("<<ComboboxSelected>>", lambda _: self.on_select())

        # 速度オプション
        ttk.Label(self, text="速度 (1–15 m/s):").grid(row=1, column=0, sticky="w", pady=(5,0))
        self.speed_var = tk.IntVar(value=15)
        self.speed_spin = ttk.Spinbox(self, from_=1, to=15, textvariable=self.speed_var, width=5)
        self.speed_spin.grid(row=1, column=1, sticky="w", pady=(5,0))

        # 撮影オプション
        self.capture_var = tk.BooleanVar(value=False)
        self.capture_chk = ttk.Checkbutton(self, text="ウェイポイントで写真撮影を行う",
                                           variable=self.capture_var,
                                           command=self.update_controls)
        self.capture_chk.grid(row=2, column=0, columnspan=3, sticky="w", pady=(5,0))

        # ジンバルオプション
        self.gimbal_var = tk.BooleanVar(value=True)
        self.gimbal_chk = ttk.Checkbutton(self, text="ジンバル操作を保持する",
                                          variable=self.gimbal_var)
        self.gimbal_chk.grid(row=3, column=0, columnspan=3, sticky="w", pady=(5,0))

        # ヨー固定オプション
        self.yaw_fix_var = tk.BooleanVar(value=False)
        self.yaw_fix_chk = ttk.Checkbutton(self, text="ヨー角を固定する",
                                           variable=self.yaw_fix_var,
                                           command=self.update_yaw_controls)
        self.yaw_fix_chk.grid(row=4, column=0, columnspan=3, sticky="w", pady=(5,0))
        self.yaw_combo = ttk.Combobox(self, values=list(YAW_FIX_OPTIONS.keys()),
                                      state="readonly", width=20)
        self.yaw_entry = ttk.Entry(self, width=10, state="disabled")
        self.yaw_combo.bind("<<ComboboxSelected>>", lambda _: self.update_yaw_controls())

        self.update_controls()

    def on_select(self):
        choice = self.combo.get()
        if HEIGHT_OPTIONS[choice] == "custom":
            self.entry.config(state="normal")
            self.entry.delete(0, tk.END)
            self.entry.focus()
        else:
            self.entry.config(state="disabled")
            self.entry.delete(0, tk.END)

    def update_controls(self):
        if self.capture_var.get():
            self.gimbal_chk.state(['disabled'])
            self.gimbal_var.set(True)
        else:
            self.gimbal_chk.state(['!disabled'])
        self.update_yaw_controls()

    def update_yaw_controls(self):
        if self.yaw_fix_var.get():
            self.yaw_combo.grid(row=5, column=0, columnspan=2, padx=5, pady=(5,0), sticky="w")
            self.yaw_entry.grid(row=5, column=2, padx=5, pady=(5,0))
            if not self.yaw_combo.get():
                self.yaw_combo.set(next(iter(YAW_FIX_OPTIONS)))
        else:
            self.yaw_combo.grid_forget()
            self.yaw_entry.grid_forget()
        if self.yaw_combo.get() == "その他 – 手動入力":
            self.yaw_entry.config(state="normal")
            self.yaw_entry.delete(0, tk.END)
            self.yaw_entry.focus()
        else:
            self.yaw_entry.config(state="disabled")
            self.yaw_entry.delete(0, tk.END)

    def get_offset(self) -> float:
        choice = self.combo.get()
        val = HEIGHT_OPTIONS[choice]
        if val == "custom":
            try:
                return float(self.entry.get())
            except:
                return 0.0
        return float(val)

    def get_speed(self) -> int:
        return max(1, min(15, self.speed_var.get()))

    def get_capture_option(self) -> bool:
        return self.capture_var.get()

    def get_gimbal_option(self) -> bool:
        return self.gimbal_var.get()

    def get_yaw_fix_option(self) -> bool:
        return self.yaw_fix_var.get()

    def get_yaw_angle(self) -> float:
        choice = self.yaw_combo.get()
        val = YAW_FIX_OPTIONS.get(choice)
        if val == "custom":
            try:
                return float(self.yaw_entry.get())
            except:
                return 0.0
        return float(val)

def main():
    root = TkinterDnD.Tk()
    root.title("ATL → ASL 高度・速度変換＆制御ツール")
    root.geometry("700x580")

    frame = ttk.Frame(root, padding=10)
    frame.pack(fill="both", expand=True)

    selector = HeightSelector(frame)
    selector.grid(row=0, column=0, columnspan=2, sticky="w", pady=(0,10))

    drop_lbl = tk.Label(frame, text="ここに .kmz をドラッグ＆ドロップ",
                        bg="lightgray", width=60, height=5, relief=tk.RIDGE)
    drop_lbl.grid(row=1, column=0, columnspan=2, pady=12, sticky="nsew")
    drop_lbl.drop_target_register(DND_FILES)

    def on_drop(event):
        path = event.data.strip("{}")
        if not path.lower().endswith(".kmz"):
            messagebox.showwarning("警告", "KMZファイルのみ対応")
            return
        offset = selector.get_offset()
        speed = selector.get_speed()
        do_capture = selector.get_capture_option()
        do_gimbal = selector.get_gimbal_option()
        fix_yaw = selector.get_yaw_fix_option()
        yaw_angle = selector.get_yaw_angle()
        threading.Thread(target=process_kmz,
                         args=(path, offset, do_capture, do_gimbal,
                               fix_yaw, yaw_angle, speed, log_text),
                         daemon=True).start()

    drop_lbl.dnd_bind("<<Drop>>", on_drop)

    log_frame = ttk.LabelFrame(frame, text="ログ")
    log_frame.grid(row=2, column=0, columnspan=2, sticky="nsew")
    frame.rowconfigure(2, weight=1)
    frame.columnconfigure(1, weight=1)
    log_text = scrolledtext.ScrolledText(log_frame, height=14, state="normal")
    log_text.pack(fill="both", expand=True)

    root.mainloop()

if __name__ == "__main__":
    main()