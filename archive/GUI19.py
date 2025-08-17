#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
convert_height_gui_asl.py
（動画撮影＋マルチセンサ選択機能搭載版）

— 機能 —
・高度オフセット＋EGM96モード変換
・速度設定（1–15 m/s）
・写真撮影オプション
・動画撮影オプション（最初で開始、最後で停止）
・写真／動画排他制御
・撮影オフ時のみジンバル制御可能
・撮影オン時はジンバル制御固定
・ヨー角固定オプション
・撮影モード選択：ワイド／ズーム／IR センサー
  （各チェックポイントで選択・制御、複数選択可）
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

# WPML 名前空間定義
NS = {
    "kml":  "http://www.opengis.net/kml/2.2",
    "wpml": "http://www.dji.com/wpmz/1.0.6"
}

# GUI定数
HEIGHT_OPTIONS = {
    "613.5 – 事務所前": 613.5,
    "962.02 – 烏帽子":   962.02,
    "その他 – 手動入力": "custom"
}
YAW_OPTIONS = {
    "固定なし": None,
    "1Q: 87.37°": 87.37,
    "2Q: 96.92°": 96.92,
    "4Q: 87.31°": 87.31,
    "手動入力":    "custom"
}
SENSOR_MODES = ["Wide", "Zoom", "IR"]


def extract_kmz(path, work_dir="_kmz_work"):
    if os.path.exists(work_dir):
        shutil.rmtree(work_dir)
    os.makedirs(work_dir)
    with zipfile.ZipFile(path, "r") as zf:
        zf.extractall(work_dir)
    return work_dir


def prepare_output_dirs(input_kmz: str, offset: float):
    base = os.path.splitext(os.path.basename(input_kmz))[0]
    sign = "+" if offset >= 0 else "-"
    mid_dir = f"{base}_asl{sign}{abs(offset)}"
    # 中間フォルダをドラッグ元と同じ階層に作成
    parent = os.path.dirname(input_kmz)
    out_root = os.path.join(parent, mid_dir)
    if os.path.exists(out_root):
        shutil.rmtree(out_root)
    os.makedirs(out_root)
    wpmz_dir = os.path.join(out_root, "wpmz")
    os.makedirs(wpmz_dir)
    return out_root, wpmz_dir


def repackage_to_kmz(out_root: str, input_kmz: str) -> str:
    """
    中間フォルダと同じ階層（ドラッグ元と同じ場所）に
    最終 KMZ ファイルを出力して返します。
    """
    base_name = os.path.splitext(os.path.basename(input_kmz))[0] + "_Converted.kmz"
    # out_root は中間フォルダのパス => その1つ上がドラッグ元と同じ階層
    target_dir = os.path.dirname(out_root)
    out_kmz = os.path.join(target_dir, base_name)
    tmp_zip = out_kmz + ".zip"

    with zipfile.ZipFile(tmp_zip, "w", zipfile.ZIP_DEFLATED) as zf:
        for root, _, files in os.walk(out_root):
            for f in files:
                full_path = os.path.join(root, f)
                rel_path = os.path.relpath(full_path, out_root)
                zf.write(full_path, rel_path)

    if os.path.exists(out_kmz):
        os.remove(out_kmz)
    os.rename(tmp_zip, out_kmz)
    return out_kmz


def convert_kml(tree: etree._ElementTree,
                offset: float,
                do_photo: bool,
                do_video: bool,
                do_gimbal: bool,
                yaw_fix: bool,
                yaw_angle,
                speed: int,
                sensor_modes: list):
    # 高度オフセット＋EGM96
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

    # 速度設定
    gts = tree.find(".//wpml:missionConfig/wpml:globalTransitionalSpeed", NS)
    if gts is not None:
        gts.text = str(speed)
    afs = tree.find(".//kml:Folder/wpml:autoFlightSpeed", NS)
    if afs is not None:
        afs.text = str(speed)
    else:
        fld = tree.find(".//kml:Folder", NS)
        etree.SubElement(fld, f"{{{NS['wpml']}}}autoFlightSpeed").text = str(speed)
    for pm in tree.findall(".//kml:Placemark", NS):
        ws = pm.find("wpml:waypointSpeed", NS)
        if ws is not None:
            ws.text = str(speed)
        else:
            new = etree.SubElement(pm, f"{{{NS['wpml']}}}waypointSpeed"); new.text = str(speed)
        ugs = pm.find("wpml:useGlobalSpeed", NS)
        if ugs is not None:
            ugs.text = "0"
        else:
            new = etree.SubElement(pm, f"{{{NS['wpml']}}}useGlobalSpeed"); new.text = "0"

    # 既存アクション整理（写真／ジンバル削除）
    for ag in tree.findall(".//wpml:actionGroup", NS):
        for act in ag.findall("wpml:action", NS):
            f = act.find("wpml:actionActuatorFunc", NS)
            if f is not None:
                if f.text == "orientedShoot" and not do_photo:
                    ag.remove(act)
                if f.text == "gimbalRotate" and not do_photo and not do_gimbal:
                    ag.remove(act)

    # 動画開始／停止アクション（最初と最後のWPにのみ追加）
    wps = tree.findall(".//kml:Placemark", NS)
    if do_video and wps:
        first = wps[0]
        ag1 = etree.SubElement(first, f"{{{NS['wpml']}}}actionGroup")
        trig1 = etree.SubElement(ag1, f"{{{NS['wpml']}}}actionTrigger")
        etree.SubElement(trig1, f"{{{NS['wpml']}}}actionTriggerType").text = "reachPoint"
        etree.SubElement(ag1, f"{{{NS['wpml']}}}actionActuatorFunc").text = "recordStart"
        last = wps[-1]
        ag2 = etree.SubElement(last, f"{{{NS['wpml']}}}actionGroup")
        trig2 = etree.SubElement(ag2, f"{{{NS['wpml']}}}actionTrigger")
        etree.SubElement(trig2, f"{{{NS['wpml']}}}actionTriggerType").text = "reachPoint"
        etree.SubElement(ag2, f"{{{NS['wpml']}}}actionActuatorFunc").text = "recordStop"

    # センサー選択＋撮影アクション（各WPごと）
    if sensor_modes:
        for pm in wps:
            ag = etree.SubElement(pm, f"{{{NS['wpml']}}}actionGroup")
            trig = etree.SubElement(ag, f"{{{NS['wpml']}}}actionTrigger")
            etree.SubElement(trig, f"{{{NS['wpml']}}}actionTriggerType").text = "reachPoint"
            for mode in sensor_modes:
                sel = etree.SubElement(ag, f"{{{NS['wpml']}}}actionActuatorFunc"); sel.text = f"select{mode}"
            if do_photo:
                shoot = etree.SubElement(ag, f"{{{NS['wpml']}}}actionActuatorFunc"); shoot.text = "orientedShoot"

    # ヨー固定
    if yaw_fix:
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
            ag.getparent().remove(ag)


def process_kmz(path, offset, do_photo, do_video, do_gimbal,
                yaw_fix, yaw_angle, speed, sensor_modes, log):
    try:
        log.insert(tk.END, f"Extracting {os.path.basename(path)}...\n")
        wd = extract_kmz(path)
        kmls = glob.glob(os.path.join(wd, "**", "template.kml"), recursive=True)
        base, outdir = prepare_output_dirs(path, offset)
        for kml in kmls:
            log.insert(tk.END, f"Converting {os.path.basename(kml)}...\n")
            tree = etree.parse(kml)
            convert_kml(tree, offset, do_photo, do_video, do_gimbal,
                        yaw_fix, yaw_angle, speed, sensor_modes)
            out = os.path.join(outdir, os.path.basename(kml))
            tree.write(out, encoding="utf-8", pretty_print=True, xml_declaration=True)
        for src in glob.glob(os.path.join(wd, "**", "res"), recursive=True):
            dst = os.path.join(outdir, "res")
            if os.path.exists(dst): shutil.rmtree(dst)
            shutil.copytree(src, dst)
        for src in glob.glob(os.path.join(wd, "**", "waylines.wpml"), recursive=True):
            shutil.copy2(src, outdir)
        outkmz = repackage_to_kmz(base, path)
        log.insert(tk.END, f"Saved: {outkmz}\nFinished\n\n")
        messagebox.showinfo("完了", f"変換完了:\n{outkmz}")
    except Exception as e:
        messagebox.showerror("エラー", str(e))
        log.insert(tk.END, f"Error: {e}\n\n")
    finally:
        if os.path.exists("_kmz_work"):
            shutil.rmtree("_kmz_work")


class HeightSelector(ttk.Frame):
    def __init__(self, master):
        super().__init__(master)
        ttk.Label(self, text="基準高度:").grid(row=0, column=0, sticky="w")
        self.hc = ttk.Combobox(self, values=list(HEIGHT_OPTIONS),
                               state="readonly", width=20)
        self.hc.set(next(iter(HEIGHT_OPTIONS)))
        self.hc.grid(row=0, column=1, padx=5)
        self.he = ttk.Entry(self, width=10, state="disabled")
        self.he.grid(row=0, column=2, padx=5)
        self.hc.bind("<<ComboboxSelected>>", lambda _: self.on_height())

        ttk.Label(self, text="速度 (1–15 m/s):").grid(row=1, column=0, sticky="w", pady=5)
        self.sp = tk.IntVar(value=15)
        ttk.Spinbox(self, from_=1, to=15, textvariable=self.sp, width=5).grid(row=1, column=1, sticky="w")

        self.ph = tk.BooleanVar(value=False)
        ttk.Checkbutton(self, text="写真撮影", variable=self.ph, command=self.update_ctrl).grid(row=2, column=0, sticky="w")
        self.vd = tk.BooleanVar(value=False)
        ttk.Checkbutton(self, text="動画撮影", variable=self.vd, command=self.update_ctrl).grid(row=2, column=1, sticky="w")

        ttk.Label(self, text="センサー選択:").grid(row=3, column=0, sticky="w")
        self.sm_vars = {mode: tk.BooleanVar(value=False) for mode in SENSOR_MODES}
        for i, mode in enumerate(SENSOR_MODES):
            ttk.Checkbutton(self, text=mode, variable=self.sm_vars[mode]).grid(row=3, column=1+i, sticky="w")

        self.gm = tk.BooleanVar(value=True)
        self.gc = ttk.Checkbutton(self, text="ジンバル制御", variable=self.gm)
        self.gc.grid(row=4, column=0, sticky="w", pady=5)

        self.yf = tk.BooleanVar(value=False)
        ttk.Checkbutton(self, text="ヨー固定", variable=self.yf, command=self.update_yaw).grid(row=5, column=0, sticky="w")
        self.yc = ttk.Combobox(self, values=list(YAW_OPTIONS), state="readonly", width=15)
        self.ye = ttk.Entry(self, width=8, state="disabled")

        self.update_ctrl()

    def on_height(self):
        if HEIGHT_OPTIONS[self.hc.get()] == "custom":
            self.he.config(state="normal"); self.he.delete(0, tk.END); self.he.focus()
        else:
            self.he.config(state="disabled"); self.he.delete(0, tk.END)

    def update_ctrl(self):
        if self.ph.get() or self.vd.get():
            self.gc.state(['disabled']); self.gm.set(True)
            if self.ph.get(): self.vd.set(False)
            if self.vd.get(): self.ph.set(False)
        else:
            self.gc.state(['!disabled'])
        self.update_yaw()

    def update_yaw(self):
        if self.yf.get():
            self.yc.grid(row=6, column=0, padx=5); self.ye.grid(row=6, column=1, padx=5)
            if not self.yc.get(): self.yc.set(next(iter(YAW_OPTIONS)))
        else:
            self.yc.grid_forget(); self.ye.grid_forget()
        if self.yc.get() == "手動入力":
            self.ye.config(state="normal"); self.ye.delete(0, tk.END); self.ye.focus()
        else:
            self.ye.config(state="disabled"); self.ye.delete(0, tk.END)

    def get_offset(self) -> float:
        v = HEIGHT_OPTIONS[self.hc.get()]
        return float(self.he.get()) if v == "custom" else float(v)

    def get_speed(self) -> int:
        return max(1, min(15, self.sp.get()))

    def get_photo(self) -> bool:
        return self.ph.get()

    def get_video(self) -> bool:
        return self.vd.get()

    def get_gimbal(self) -> bool:
        return self.gm.get()

    def get_yawfix(self) -> bool:
        return self.yf.get()

    def get_yawangle(self) -> float:
        val = YAW_OPTIONS.get(self.yc.get())
        if val == "custom":
            try:
                return float(self.ye.get())
            except:
                return 0.0
        return float(val) if val else 0.0

    def get_sensors(self) -> list:
        return [mode for mode, var in self.sm_vars.items() if var.get()]


def main():
    root = TkinterDnD.Tk()
    root.title("ATL→ASL 変換＋撮影制御ツール")
    root.geometry("750x650")
    frm = ttk.Frame(root, padding=10)
    frm.pack(fill="both", expand=True)

    sel = HeightSelector(frm)
    sel.grid(row=0, column=0, columnspan=4, sticky="w", pady=(0, 10))

    drop = tk.Label(frm, text=".kmzをここにドロップ", bg="lightgray",
                    width=70, height=5, relief=tk.RIDGE)
    drop.grid(row=1, column=0, columnspan=4, pady=12, sticky="nsew")
    drop.drop_target_register(DND_FILES)

    log_frame = ttk.LabelFrame(frm, text="ログ")
    log_frame.grid(row=2, column=0, columnspan=4, sticky="nsew")
    frm.rowconfigure(2, weight=1); frm.columnconfigure(3, weight=1)
    log = scrolledtext.ScrolledText(log_frame, height=16)
    log.pack(fill="both", expand=True)

    def on_drop(event):
        path = event.data.strip("{}")
        if not path.lower().endswith(".kmz"):
            messagebox.showwarning("警告", ".kmzファイルのみ対応")
            return
        args = (
            path,
            sel.get_offset(),
            sel.get_photo(),
            sel.get_video(),
            sel.get_gimbal(),
            sel.get_yawfix(),
            sel.get_yawangle(),
            sel.get_speed(),
            sel.get_sensors(),
            log
        )
        threading.Thread(target=process_kmz, args=args, daemon=True).start()

    drop.dnd_bind("<<Drop>>", on_drop)

    root.mainloop()


if __name__ == "__main__":
    main()
