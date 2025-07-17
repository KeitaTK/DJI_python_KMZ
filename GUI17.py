#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
convert_height_gui_asl.py (拡張版)
---------------------------------
KMZ 内の WPML 高度をオフセット付きで変換し、
relativeToStartPoint → EGM96 に置換。
・ウェイポイントでの写真撮影をオプション化
・撮影オフ時のみジンバル制御のオン／オフ選択可能
・撮影オン時はジンバル制御も実行
・ヨー角固定オプション追加：飛行中ずっと指定角度を保持
  └ 既定選択肢と「その他–手動入力」で任意角度指定可
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

# GUI 用：基準高度選択リスト
HEIGHT_OPTIONS = {
    "613.5 – 事務所前": 613.5,
    "962.02 – 烏帽子": 962.02,
    "その他 – 手動入力": "custom"
}
# ヨー固定選択リスト
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

def prepare_output_dirs(input_kmz: str, offset: float) -> (str, str):
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

def convert_heights_and_mode(tree: etree._ElementTree, offset: float,
                             do_capture: bool, do_gimbal: bool,
                             fix_yaw: bool, yaw_angle: float) -> None:
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

    # アクション処理
    # Placemarkごとの actionGroup をループ
    for ag in tree.findall(".//wpml:actionGroup", NS):
        # orientedShoot 削除
        for action in ag.findall("wpml:action", NS):
            f = action.find("wpml:actionActuatorFunc", NS)
            if f is not None and f.text == "orientedShoot" and not do_capture:
                ag.remove(action)
        # gimbalRotate 削除
        if not do_capture and not do_gimbal:
            for action in ag.findall("wpml:action", NS):
                f = action.find("wpml:actionActuatorFunc", NS)
                if f is not None and f.text == "gimbalRotate":
                    ag.remove(action)

    # ヨー固定処理（global & each WP）
    if fix_yaw:
        # グローバル設定
        gw = tree.find(".//wpml:globalWaypointHeadingParam", NS)
        if gw is not None:
            gw.find("wpml:waypointHeadingMode", NS).text = "fixed"
            gw.find("wpml:waypointHeadingAngle", NS).text = str(yaw_angle)
        # 各Waypoint
        for wp in tree.findall(".//wpml:waypointHeadingParam", NS):
            wp.find("wpml:waypointHeadingMode", NS).text = "fixed"
            wp.find("wpml:waypointHeadingAngle", NS).text = str(yaw_angle)

    # 空の actionGroup は削除
    for ag in tree.findall(".//wpml:actionGroup", NS):
        if not ag.findall("wpml:action", NS):
            parent = ag.getparent()
            parent.remove(ag)

def process_kmz(input_kmz: str, offset: float,
                do_capture: bool, do_gimbal: bool,
                fix_yaw: bool, yaw_angle: float,
                log: tk.Text) -> None:
    work = "_kmz_work"
    try:
        log.insert(tk.END, f"Extracting {os.path.basename(input_kmz)}...\n")
        extract_kmz(input_kmz, work)
        kmls = glob.glob(os.path.join(work, "**", "template.kml"), recursive=True)
        if not kmls:
            raise FileNotFoundError("template.kml が見つかりません")
        base_out, wpmz_dir = prepare_output_dirs(input_kmz, offset)
        for kml in kmls:
            log.insert(tk.END, f"Converting {os.path.basename(kml)}...\n")
            tree = etree.parse(kml)
            convert_heights_and_mode(tree, offset, do_capture, do_gimbal, fix_yaw, yaw_angle)
            out = os.path.join(wpmz_dir, os.path.basename(kml))
            tree.write(out, encoding="utf-8", pretty_print=True, xml_declaration=True)
        # res フォルダコピー
        for src in glob.glob(os.path.join(work, "**", "res"), recursive=True):
            dst = os.path.join(wpmz_dir, "res")
            if os.path.exists(dst):
                shutil.rmtree(dst)
            shutil.copytree(src, dst)
            log.insert(tk.END, "Copied res folder\n")
        # waylines.wpml コピー
        for src in glob.glob(os.path.join(work, "**", "waylines.wpml"), recursive=True):
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
        if os.path.exists(work):
            shutil.rmtree(work)

# --- GUI構築 ---
class HeightSelector(ttk.Frame):
    def __init__(self, master):
        super().__init__(master)
        # 基準高度
        ttk.Label(self, text="基準高度オフセット:").grid(row=0, column=0, sticky="w")
        self.combo = ttk.Combobox(self, values=list(HEIGHT_OPTIONS.keys()),
                                  state="readonly", width=20)
        self.combo.set(next(iter(HEIGHT_OPTIONS)))
        self.combo.grid(row=0, column=1, padx=5)
        self.entry = ttk.Entry(self, width=10, state="disabled")
        self.entry.grid(row=0, column=2, padx=5)
        self.combo.bind("<<ComboboxSelected>>", lambda _: self.on_select())

        # 写真撮影オプション
        self.capture_var = tk.BooleanVar(value=False)
        self.capture_chk = ttk.Checkbutton(self, text="ウェイポイントで写真撮影を行う",
                                           variable=self.capture_var,
                                           command=self.update_controls)
        self.capture_chk.grid(row=1, column=0, columnspan=3, sticky="w", pady=(5,0))

        # ジンバル操作オプション
        self.gimbal_var = tk.BooleanVar(value=True)
        self.gimbal_chk = ttk.Checkbutton(self, text="ジンバル操作を保持する",
                                          variable=self.gimbal_var)
        self.gimbal_chk.grid(row=2, column=0, columnspan=3, sticky="w", pady=(5,0))

        # ヨー固定オプション
        self.yaw_fix_var = tk.BooleanVar(value=False)
        self.yaw_fix_chk = ttk.Checkbutton(self, text="ヨー角を固定する",
                                           variable=self.yaw_fix_var,
                                           command=self.update_yaw_controls)
        self.yaw_fix_chk.grid(row=3, column=0, columnspan=3, sticky="w", pady=(5,0))

        # ヨー角選択（初期非表示）
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
        # 写真撮影オン時はジンバル不可操作
        if self.capture_var.get():
            self.gimbal_chk.state(["disabled"])
            self.gimbal_var.set(True)
        else:
            self.gimbal_chk.state(["!disabled"])
        # ヨー角UIは常に切り替え
        self.update_yaw_controls()

    def update_yaw_controls(self):
        if self.yaw_fix_var.get():
            self.yaw_combo.grid(row=4, column=0, columnspan=2, padx=5, pady=(5,0), sticky="w")
            self.yaw_entry.grid(row=4, column=2, padx=5, pady=(5,0))
            if not self.yaw_combo.get():
                self.yaw_combo.set(next(iter(YAW_FIX_OPTIONS)))
        else:
            self.yaw_combo.grid_forget()
            self.yaw_entry.grid_forget()
        # 手動入力切替
        if self.yaw_combo.get() == "その他 – 手動入力":
            self.yaw_entry.config(state="normal")
            self.yaw_entry.delete(0, tk.END)
            self.yaw_entry.focus()
        else:
            self.yaw_entry.config(state="disabled")
            self.yaw_entry.delete(0, tk.END)

    def get_offset(self) -> float:
        c = self.combo.get()
        val = HEIGHT_OPTIONS[c]
        if val == "custom":
            try:
                return float(self.entry.get())
            except:
                return 0.0
        return float(val)

    def get_capture_option(self) -> bool:
        return self.capture_var.get()

    def get_gimbal_option(self) -> bool:
        return self.gimbal_var.get()

    def get_yaw_fix_option(self) -> bool:
        return self.yaw_fix_var.get()

    def get_yaw_angle(self) -> float:
        c = self.yaw_combo.get()
        val = YAW_FIX_OPTIONS[c]
        if val == "custom":
            try:
                return float(self.yaw_entry.get())
            except:
                return 0.0
        return float(val)

def main():
    root = TkinterDnD.Tk()
    root.title("ATL → ASL 高度変換＆カメラ/ジンバル/ヨー制御ツール")
    root.geometry("700x480")

    frame = ttk.Frame(root, padding=10)
    frame.pack(fill="both", expand=True)

    height_selector = HeightSelector(frame)
    height_selector.grid(row=0, column=0, columnspan=2, sticky="w", pady=(0,10))

    drop_lbl = tk.Label(frame, text="ここに .kmz をドラッグ＆ドロップ",
                        bg="lightgray", width=60, height=5, relief=tk.RIDGE)
    drop_lbl.grid(row=1, column=0, columnspan=2, pady=12, sticky="nsew")
    drop_lbl.drop_target_register(DND_FILES)

    def on_drop(event):
        path = event.data.strip("{}")
        if not path.lower().endswith(".kmz"):
            messagebox.showwarning("警告", "KMZファイルのみ対応")
            return
        offset = height_selector.get_offset()
        do_capture = height_selector.get_capture_option()
        do_gimbal = height_selector.get_gimbal_option()
        fix_yaw = height_selector.get_yaw_fix_option()
        yaw_angle = height_selector.get_yaw_angle()
        threading.Thread(target=process_kmz,
                         args=(path, offset, do_capture, do_gimbal,
                               fix_yaw, yaw_angle, log_text),
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
