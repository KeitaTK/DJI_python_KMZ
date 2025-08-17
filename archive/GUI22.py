#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
convert_height_gui_asl.py (ver. GUI20-Fix-Final)
— 機能 —
・高度オフセット＋EGM96モード変換
・速度設定（1–15 m/s）
・写真撮影オプション
    ・<imageFormat>タグによる複数カメラ同時指定方式に対応
    ・★動画撮影選択時は写真撮影アクションを完全に削除
・動画撮影オプション（最初で開始、最後で停止）
    ・★動画撮影時はジンバルピッチを真下（-90度）に固定
・写真／動画排他制御
・撮影オフ時のみジンバル制御可能
・撮影オン時はジンバル制御固定（ただし動画時は-90度固定優先）
・ヨー角固定オプション
・撮影モード選択：ワイド／ズーム／IR センサー
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
    "kml": "http://www.opengis.net/kml/2.2",
    "wpml": "http://www.dji.com/wpmz/1.0.6"
}

# GUI定数
HEIGHT_OPTIONS = {
    "613.5 – 事務所前": 613.5,
    "962.02 – 烏帽子": 962.02,
    "その他 – 手動入力": "custom"
}

YAW_OPTIONS = {
    "固定なし": None,
    "1Q: 87.37°": 87.37,
    "2Q: 96.92°": 96.92,
    "4Q: 87.31°": 87.31,
    "手動入力": "custom"
}

SENSOR_MODES = ["Wide", "Zoom", "IR"]

def extract_kmz(path, work_dir="_kmz_work"):
    if os.path.exists(work_dir): shutil.rmtree(work_dir)
    os.makedirs(work_dir)
    with zipfile.ZipFile(path, "r") as zf: zf.extractall(work_dir)
    return work_dir

def prepare_output_dirs(input_kmz: str, offset: float):
    base = os.path.splitext(os.path.basename(input_kmz))[0]
    sign = "+" if offset >= 0 else "-"
    mid_dir = f"{base}_asl{sign}{abs(offset)}"
    parent = os.path.dirname(input_kmz)
    out_root = os.path.join(parent, mid_dir)
    if os.path.exists(out_root): shutil.rmtree(out_root)
    os.makedirs(out_root)
    wpmz_dir = os.path.join(out_root, "wpmz")
    os.makedirs(wpmz_dir)
    return out_root, wpmz_dir

def repackage_to_kmz(out_root: str, input_kmz: str) -> str:
    base_name = os.path.splitext(os.path.basename(input_kmz))[0] + "_Converted.kmz"
    target_dir = os.path.dirname(out_root)
    out_kmz = os.path.join(target_dir, base_name)
    tmp_zip = out_kmz + ".zip"
    with zipfile.ZipFile(tmp_zip, "w", zipfile.ZIP_DEFLATED) as zf:
        for root, _, files in os.walk(out_root):
            for f in files:
                full_path = os.path.join(root, f)
                rel_path = os.path.relpath(full_path, out_root)
                zf.write(full_path, rel_path)
    if os.path.exists(out_kmz): os.remove(out_kmz)
    os.rename(tmp_zip, out_kmz)
    return out_kmz

def convert_kml(tree: etree._ElementTree,
                offset: float,
                do_photo: bool,
                do_video: bool,
                video_suffix: str,
                do_gimbal: bool,
                yaw_fix: bool,
                yaw_angle,
                speed: int,
                sensor_modes: list):
    
    # --- ヘルパー関数 ---
    def _get_max_id(root, xpath_query):
        max_id = -1
        for elem in root.findall(xpath_query, NS):
            if elem.text and elem.text.isdigit(): max_id = max(max_id, int(elem.text))
        return max_id

    def _get_or_create_action_group(placemark, wp_index, current_max_group_id):
        action_group = placemark.find('wpml:actionGroup', NS)
        if action_group is None:
            action_group = etree.SubElement(placemark, f"{{{NS['wpml']}}}actionGroup")
            etree.SubElement(action_group, f"{{{NS['wpml']}}}actionGroupId").text = str(current_max_group_id + 1)
            etree.SubElement(action_group, f"{{{NS['wpml']}}}actionGroupStartIndex").text = str(wp_index)
            etree.SubElement(action_group, f"{{{NS['wpml']}}}actionGroupEndIndex").text = str(wp_index)
            etree.SubElement(action_group, f"{{{NS['wpml']}}}actionGroupMode").text = 'sequence'
            trigger = etree.SubElement(action_group, f"{{{NS['wpml']}}}actionTrigger")
            etree.SubElement(trigger, f"{{{NS['wpml']}}}actionTriggerType").text = 'reachPoint'
        max_action_id = -1
        for elem in action_group.findall('.//wpml:actionId', NS):
            if elem.text and elem.text.isdigit(): max_action_id = max(max_action_id, int(elem.text))
        return action_group, max_action_id + 1

    # --- KML変換処理 ---

    # 1. 高度オフセット＋EGM96
    for pm in tree.findall(".//kml:Placemark", NS):
        for tag in ("height", "ellipsoidHeight"):
            e = pm.find(f"wpml:{tag}", NS)
            if e is not None and e.text is not None:
                try: e.text = str(float(e.text) + offset)
                except (ValueError, TypeError): pass
    gh = tree.find(".//wpml:globalHeight", NS)
    if gh is not None and gh.text is not None:
        try: gh.text = str(float(gh.text) + offset)
        except (ValueError, TypeError): pass
    for hm in tree.findall(".//wpml:heightMode", NS): hm.text = "EGM96"

    # 2. 速度設定
    for speed_tag, parent_xpath in [
        ("globalTransitionalSpeed", ".//wpml:missionConfig"),
        ("autoFlightSpeed", ".//kml:Folder")
    ]:
        elem = tree.find(f"{parent_xpath}/wpml:{speed_tag}", NS)
        if elem is not None: elem.text = str(speed)
        else:
            parent = tree.find(parent_xpath, NS)
            if parent is not None: etree.SubElement(parent, f"{{{NS['wpml']}}}{speed_tag}").text = str(speed)
    for pm in tree.findall(".//kml:Placemark", NS):
        for tag, val in [("waypointSpeed", str(speed)), ("useGlobalSpeed", "0")]:
            elem = pm.find(f"wpml:{tag}", NS)
            if elem is not None: elem.text = val
            else: etree.SubElement(pm, f"{{{NS['wpml']}}}{tag}").text = val

    # 3. 既存アクションのクリーンアップ
    # 既存のカメラ・撮影関連のアクションを全て削除し、GUIの設定を正しく反映させる
    for ag in tree.findall(".//wpml:actionGroup", NS):
        for act in list(ag.findall("wpml:action", NS)):
            f = act.find("wpml:actionActuatorFunc", NS)
            if f is None or f.text is None: continue
            
            # 関連アクションを無条件にクリア
            if f.text in ("orientedShoot", "startRecord", "stopRecord", "gimbalRotate", 
                          "selectWide", "selectZoom", "selectIR"):
                ag.remove(act)
    
    # 既存のimageFormatタグを削除
    payload_param_elem = tree.find(".//wpml:payloadParam", NS)
    if payload_param_elem is not None:
        image_format_elem = payload_param_elem.find("wpml:imageFormat", NS)
        if image_format_elem is not None:
            payload_param_elem.remove(image_format_elem)
            
    # globalWaypointTurnMode/fixedYawAndPitchのジンバル設定をクリア
    global_turn_mode_elem = tree.find(".//wpml:globalWaypointTurnMode", NS)
    if global_turn_mode_elem is not None and global_turn_mode_elem.text == "fixedYawAndPitch":
        gimbal_pitch_elem = tree.find(".//wpml:gimbalPitch", NS)
        if gimbal_pitch_elem is not None:
            gimbal_pitch_elem.getparent().remove(gimbal_pitch_elem)

    # 4. 新しいカメラ・撮影アクションの追加
    
    # 4-1. カメラ選択 (imageFormat タグを使用)
    if sensor_modes:
        # payloadParamセクションがなければ作成
        if payload_param_elem is None:
            folder = tree.find(".//kml:Folder", NS)
            if folder is not None:
                payload_param_elem = etree.SubElement(folder, f"{{{NS['wpml']}}}payloadParam")
                etree.SubElement(payload_param_elem, f"{{{NS['wpml']}}}payloadPositionIndex").text = '0'
        
        # imageFormatタグを追加
        if payload_param_elem is not None:
            format_string = ",".join([mode.lower() for mode in sensor_modes])
            etree.SubElement(payload_param_elem, f"{{{NS['wpml']}}}imageFormat").text = format_string

    # 4-2. 写真撮影 (orientedShoot) - 動画撮影が選択されていない場合のみ追加
    if do_photo and not do_video: # ★修正: do_videoがTrueの場合は写真撮影しない
        placemarks = [p for p in tree.findall(".//kml:Placemark", NS) if p.find("wpml:index", NS) is not None]
        max_group_id = _get_max_id(tree.getroot(), ".//wpml:actionGroupId")
        for i, pm in enumerate(placemarks):
            wp_index = pm.find("wpml:index", NS).text
            group, next_action_id = _get_or_create_action_group(pm, wp_index, max_group_id + i)
            action = etree.SubElement(group, f"{{{NS['wpml']}}}action")
            etree.SubElement(action, f"{{{NS['wpml']}}}actionId").text = str(next_action_id)
            etree.SubElement(action, f"{{{NS['wpml']}}}actionActuatorFunc").text = "orientedShoot"
            etree.SubElement(action, f"{{{NS['wpml']}}}actionActuatorFuncParam")

    # 4-3. 動画撮影 (start/stopRecord) & ジンバル制御 (★修正: ジンバル制御を追加)
    if do_video:
        all_placemarks = sorted(
            [p for p in tree.findall(".//kml:Placemark", NS) if p.find("wpml:index", NS) is not None],
            key=lambda p: int(p.find("wpml:index", NS).text)
        )
        if all_placemarks:
            first_wp, last_wp = all_placemarks[0], all_placemarks[-1]
            max_group_id = _get_max_id(tree.getroot(), ".//wpml:actionGroupId")

            # ジンバルピッチを真下（-90度）に設定
            # globalWaypointTurnModeをfixedYawAndPitchに設定
            global_turn_mode_elem = tree.find(".//wpml:globalWaypointTurnMode", NS)
            if global_turn_mode_elem is not None:
                global_turn_mode_elem.text = "fixedYawAndPitch"
            else:
                # 存在しない場合はFolder要素に追加
                folder_elem = tree.find(".//kml:Folder", NS)
                if folder_elem is not None:
                    global_turn_mode_elem = etree.SubElement(folder_elem, f"{{{NS['wpml']}}}globalWaypointTurnMode")
                    global_turn_mode_elem.text = "fixedYawAndPitch"

            # gimbalPitchタグを追加または更新
            gimbal_pitch_elem = tree.find(".//wpml:gimbalPitch", NS)
            if gimbal_pitch_elem is None:
                # globalWaypointTurnModeの後に挿入
                if global_turn_mode_elem is not None:
                    gimbal_pitch_elem = etree.Element(f"{{{NS['wpml']}}}gimbalPitch")
                    global_turn_mode_elem.addnext(gimbal_pitch_elem)
                else: # Fallback: Folderの最後に追加
                    folder_elem = tree.find(".//kml:Folder", NS)
                    if folder_elem is not None:
                        gimbal_pitch_elem = etree.SubElement(folder_elem, f"{{{NS['wpml']}}}gimbalPitch")
            
            if gimbal_pitch_elem is not None:
                gimbal_pitch_elem.text = '-90.0' # 真下を向くよう設定

            # startRecordアクションを追加
            group_start, next_action_id_start = _get_or_create_action_group(first_wp, first_wp.find("wpml:index", NS).text, max_group_id)
            action_start = etree.SubElement(group_start, f"{{{NS['wpml']}}}action")
            etree.SubElement(action_start, f"{{{NS['wpml']}}}actionId").text = str(next_action_id_start)
            etree.SubElement(action_start, f"{{{NS['wpml']}}}actionActuatorFunc").text = 'startRecord'
            param_start = etree.SubElement(action_start, f"{{{NS['wpml']}}}actionActuatorFuncParam")
            etree.SubElement(param_start, f"{{{NS['wpml']}}}fileSuffix").text = video_suffix
            etree.SubElement(param_start, f"{{{NS['wpml']}}}payloadPositionIndex").text = '0'

            # stopRecordアクションを追加
            group_stop, next_action_id_stop = _get_or_create_action_group(last_wp, last_wp.find("wpml:index", NS).text, max_group_id + 1)
            action_stop = etree.SubElement(group_stop, f"{{{NS['wpml']}}}action")
            etree.SubElement(action_stop, f"{{{NS['wpml']}}}actionId").text = str(next_action_id_stop)
            etree.SubElement(action_stop, f"{{{NS['wpml']}}}actionActuatorFunc").text = 'stopRecord'
            param_stop = etree.SubElement(action_stop, f"{{{NS['wpml']}}}actionActuatorFuncParam")
            etree.SubElement(param_stop, f"{{{NS['wpml']}}}payloadPositionIndex").text = '0'

    # 5. ヨー固定
    if yaw_fix and yaw_angle is not None:
        for heading_param in tree.findall(".//wpml:waypointHeadingParam", NS) + tree.findall(".//wpml:globalWaypointHeadingParam", NS):
            mode_el = heading_param.find("wpml:waypointHeadingMode", NS)
            angle_el = heading_param.find("wpml:waypointHeadingAngle", NS)
            if mode_el is not None: mode_el.text = "fixed"
            if angle_el is not None: angle_el.text = str(yaw_angle)

    # 6. 空の actionGroup を削除
    for ag in tree.findall(".//wpml:actionGroup", NS):
        if not ag.findall("wpml:action", NS):
            ag.getparent().remove(ag)


def process_kmz(path, offset, do_photo, do_video, video_suffix, do_gimbal,
                yaw_fix, yaw_angle, speed, sensor_modes, log):
    try:
        log.insert(tk.END, f"Extracting {os.path.basename(path)}...\n")
        wd = extract_kmz(path)
        kmls = glob.glob(os.path.join(wd, "**", "template.kml"), recursive=True)
        if not kmls: raise FileNotFoundError("template.kmlが見つかりませんでした。")
        
        out_root, outdir = prepare_output_dirs(path, offset)

        for kml_path in kmls:
            log.insert(tk.END, f"Converting {os.path.basename(kml_path)}...\n")
            parser = etree.XMLParser(remove_blank_text=True)
            tree = etree.parse(kml_path, parser)
            
            convert_kml(tree, offset, do_photo, do_video, video_suffix, do_gimbal,
                        yaw_fix, yaw_angle, speed, sensor_modes)
            
            out_path = os.path.join(outdir, os.path.basename(kml_path))
            tree.write(out_path, encoding="utf-8", pretty_print=True, xml_declaration=True)

        for item_name in ["res", "waylines.wpml"]:
            src = glob.glob(os.path.join(wd, "**", item_name), recursive=True)
            if src:
                src_path = src[0]
                dst_path = os.path.join(outdir, os.path.basename(src_path))
                if os.path.isdir(src_path):
                    if os.path.exists(dst_path): shutil.rmtree(dst_path)
                    shutil.copytree(src_path, dst_path)
                else: shutil.copy2(src_path, dst_path)

        outkmz = repackage_to_kmz(out_root, path)
        log.insert(tk.END, f"Saved: {outkmz}\nFinished\n\n")
        messagebox.showinfo("完了", f"変換完了:\n{outkmz}")
    except Exception as e:
        messagebox.showerror("エラー", str(e))
        log.insert(tk.END, f"Error: {e}\n\n")
    finally:
        if os.path.exists("_kmz_work"): shutil.rmtree("_kmz_work")

class AppGUI(ttk.Frame):
    def __init__(self, master):
        super().__init__(master)
        
        ttk.Label(self, text="基準高度:").grid(row=0, column=0, sticky="w")
        self.hc = ttk.Combobox(self, values=list(HEIGHT_OPTIONS), state="readonly", width=20)
        self.hc.set(next(iter(HEIGHT_OPTIONS)))
        self.hc.grid(row=0, column=1, padx=5, columnspan=2, sticky="w")
        self.hc.bind("<<ComboboxSelected>>", self.on_height_change)
        self.he = ttk.Entry(self, width=10, state="disabled")
        self.he.grid(row=0, column=3, padx=5)

        ttk.Label(self, text="速度 (1–15 m/s):").grid(row=1, column=0, sticky="w", pady=5)
        self.sp = tk.IntVar(value=15)
        ttk.Spinbox(self, from_=1, to=15, textvariable=self.sp, width=5).grid(row=1, column=1, sticky="w", columnspan=2)

        self.ph = tk.BooleanVar(value=False)
        ttk.Checkbutton(self, text="写真撮影", variable=self.ph, command=self.update_ctrl).grid(row=2, column=0, sticky="w")
        self.vd = tk.BooleanVar(value=False)
        ttk.Checkbutton(self, text="動画撮影", variable=self.vd, command=self.update_ctrl).grid(row=2, column=1, sticky="w")
        
        self.vd_suffix_label = ttk.Label(self, text="動画ファイル名:")
        self.vd_suffix_var = tk.StringVar(value="video_01")
        self.vd_suffix_entry = ttk.Entry(self, textvariable=self.vd_suffix_var, width=20)
        
        ttk.Label(self, text="センサー選択:").grid(row=3, column=0, sticky="w")
        self.sm_vars = {mode: tk.BooleanVar(value=False) for mode in SENSOR_MODES}
        for i, mode in enumerate(SENSOR_MODES):
            ttk.Checkbutton(self, text=mode, variable=self.sm_vars[mode]).grid(row=3, column=1 + i, sticky="w")

        self.gm = tk.BooleanVar(value=True)
        self.gc = ttk.Checkbutton(self, text="ジンバル制御", variable=self.gm)
        self.gc.grid(row=4, column=0, sticky="w", pady=5)
        
        self.yf = tk.BooleanVar(value=False)
        ttk.Checkbutton(self, text="ヨー固定", variable=self.yf, command=self.update_yaw).grid(row=5, column=0, sticky="w")
        self.yc = ttk.Combobox(self, values=list(YAW_OPTIONS), state="readonly", width=15)
        self.yc.bind("<<ComboboxSelected>>", self.update_yaw)
        self.ye = ttk.Entry(self, width=8, state="disabled")
        
        self.update_ctrl()
        self.update_yaw()

    def on_height_change(self, event=None):
        if HEIGHT_OPTIONS.get(self.hc.get()) == "custom":
            self.he.config(state="normal"); self.he.delete(0, tk.END); self.he.focus()
        else: self.he.config(state="disabled"); self.he.delete(0, tk.END)

    def update_ctrl(self):
        # 写真と動画は排他
        if self.ph.get(): self.vd.set(False)
        if self.vd.get(): self.ph.set(False)
        
        # 撮影中はジンバル制御を固定（ただし、動画撮影時はジンバルピッチは-90度に設定される）
        if self.ph.get() or self.vd.get():
            self.gc.state(['disabled']); self.gm.set(True)
        else:
            self.gc.state(['!disabled'])

        # 動画撮影時のみサフィックス入力を表示
        if self.vd.get():
            self.vd_suffix_label.grid(row=2, column=2, sticky="e", padx=(10, 2))
            self.vd_suffix_entry.grid(row=2, column=3, sticky="w")
        else:
            self.vd_suffix_label.grid_forget(); self.vd_suffix_entry.grid_forget()

    def update_yaw(self, event=None):
        if self.yf.get():
            self.yc.grid(row=5, column=1, padx=5, columnspan=2, sticky="w")
            if not self.yc.get(): self.yc.set(next(iter(YAW_OPTIONS)))
            if self.yc.get() == "手動入力": self.ye.config(state="normal"); self.ye.grid(row=5, column=3)
            else: self.ye.config(state="disabled"); self.ye.grid_forget()
        else: self.yc.grid_forget(); self.ye.grid_forget()

    def get_params(self):
        offset = 0.0
        v = HEIGHT_OPTIONS.get(self.hc.get())
        if v == "custom":
            try: offset = float(self.he.get())
            except ValueError: pass
        elif v is not None: offset = float(v)
        
        yaw_angle = None
        if self.yf.get():
            y_val = YAW_OPTIONS.get(self.yc.get())
            if y_val == "custom":
                try: yaw_angle = float(self.ye.get())
                except (ValueError, TypeError): pass
            elif y_val is not None: yaw_angle = float(y_val)
        
        return {
            "offset": offset, "do_photo": self.ph.get(),
            "do_video": self.vd.get(), "video_suffix": self.vd_suffix_var.get(),
            "do_gimbal": self.gm.get(), "yaw_fix": self.yf.get(),
            "yaw_angle": yaw_angle, "speed": max(1, min(15, self.sp.get())),
            "sensor_modes": [mode for mode, var in self.sm_vars.items() if var.get()]
        }

def main():
    root = TkinterDnD.Tk()
    root.title("ATL→ASL 変換＋撮影制御ツール (ver. GUI20-Fix-Final)")
    root.geometry("750x650")

    frm = ttk.Frame(root, padding=10)
    frm.pack(fill="both", expand=True)

    app_gui = AppGUI(frm)
    app_gui.pack(fill="x", pady=(0, 10))

    drop_label = tk.Label(frm, text=".kmzをここにドロップ", bg="lightgray", width=70, height=5, relief=tk.RIDGE)
    drop_label.pack(pady=12, fill="x")
    drop_label.drop_target_register(DND_FILES)

    log_frame = ttk.LabelFrame(frm, text="ログ")
    log_frame.pack(fill="both", expand=True)
    log_text = scrolledtext.ScrolledText(log_frame, height=16)
    log_text.pack(fill="both", expand=True)

    def on_drop(event):
        path = event.data.strip('{}')
        if not path.lower().endswith(".kmz"):
            messagebox.showwarning("警告", ".kmzファイルのみ対応しています。")
            return
        
        params = app_gui.get_params()
        params["path"] = path
        params["log"] = log_text
        
        threading.Thread(target=process_kmz, kwargs=params, daemon=True).start()

    drop_label.dnd_bind("<<Drop>>", on_drop)
    root.mainloop()

if __name__ == "__main__":
    main()
