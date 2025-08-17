#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
convert_height_gui_asl.py (ver. GUI24)

変更点
・動画撮影開始前の gimbalRotate を DJI ファーム準拠の詳細タグで生成
・GUIタイトルを ver. GUI24 に更新
・stopRecord に正しい actionId 採番
・do_gimbal フラグを反映してジンバル操作をオプション化
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

# --- 定数 ------------------------------------------------------------------
NS = {
    "kml":  "http://www.opengis.net/kml/2.2",
    "wpml": "http://www.dji.com/wpmz/1.0.6"
}

HEIGHT_OPTIONS = {
    "613.5 – 事務所前": 613.5,
    "962.02 – 烏帽子":  962.02,
    "その他 – 手動入力": "custom"
}

YAW_OPTIONS = {
    "1Q: 87.37°":  87.37,
    "2Q: 96.92°":  96.92,
    "4Q: 87.31°":  87.31,
    "手動入力":    "custom"
}

SENSOR_MODES = ["Wide", "Zoom", "IR"]

# --- KMZ ユーティリティ -----------------------------------------------------
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
                if f.lower().endswith(".wpml"):
                    continue
                full = os.path.join(root, f)
                rel = os.path.relpath(full, out_root)
                zf.write(full, rel)
    if os.path.exists(out_kmz):
        os.remove(out_kmz)
    os.rename(tmp, out_kmz)
    return out_kmz

# --- KML 変換 ---------------------------------------------------------------
def convert_kml(tree, offset, do_photo, do_video, video_suffix,
                do_gimbal, yaw_fix, yaw_angle, speed, sensor_modes):

    def _max_id(root, xp):
        ids = [int(e.text) for e in root.findall(xp, NS)
               if e.text and e.text.isdigit()]
        return max(ids) if ids else -1

    def _grp(pm, idx, curr):
        ag = pm.find("wpml:actionGroup", NS)
        if ag is None:
            ag = etree.SubElement(pm, f"{{{NS['wpml']}}}actionGroup")
            for tag, val in [
                ("actionGroupId",      curr + 1),
                ("actionGroupStartIndex", idx),
                ("actionGroupEndIndex",   idx),
                ("actionGroupMode",   "sequence")
            ]:
                etree.SubElement(ag, f"{{{NS['wpml']}}}{tag}").text = str(val)
            trg = etree.SubElement(ag, f"{{{NS['wpml']}}}actionTrigger")
            etree.SubElement(trg, f"{{{NS['wpml']}}}actionTriggerType").text = "reachPoint"
        next_id = _max_id(ag, ".//wpml:actionId") + 1
        return ag, next_id

    # 1) 高度補正＋EGM96
    for pm in tree.findall(".//kml:Placemark", NS):
        for tag in ("height", "ellipsoidHeight"):
            el = pm.find(f"wpml:{tag}", NS)
            if el is not None and el.text:
                try:
                    el.text = str(float(el.text) + offset)
                except:
                    pass
    gh = tree.find(".//wpml:globalHeight", NS)
    if gh is not None and gh.text:
        try:
            gh.text = str(float(gh.text) + offset)
        except:
            pass
    for hm in tree.findall(".//wpml:heightMode", NS):
        hm.text = "EGM96"

    # 2) 速度設定
    for tag, path in [
        ("globalTransitionalSpeed", ".//wpml:missionConfig"),
        ("autoFlightSpeed",         ".//kml:Folder")
    ]:
        el = tree.find(f"{path}/wpml:{tag}", NS)
        if el is not None:
            el.text = str(speed)
        else:
            parent = tree.find(path, NS)
            if parent is not None:
                etree.SubElement(parent, f"{{{NS['wpml']}}}{tag}").text = str(speed)
    for pm in tree.findall(".//kml:Placemark", NS):
        for t, v in [("waypointSpeed", speed), ("useGlobalSpeed", 0)]:
            el = pm.find(f"wpml:{t}", NS)
            if el is not None:
                el.text = str(v)
            else:
                etree.SubElement(pm, f"{{{NS['wpml']}}}{t}").text = str(v)

    # 3) 既存アクション削除
    for ag in tree.findall(".//wpml:actionGroup", NS):
        for act in list(ag.findall("wpml:action", NS)):
            f = act.find("wpml:actionActuatorFunc", NS)
            if f is not None and f.text in (
                "orientedShoot", "startRecord", "stopRecord",
                "gimbalRotate", "selectWide", "selectZoom", "selectIR"
            ):
                ag.remove(act)
    pp = tree.find(".//wpml:payloadParam", NS)
    if pp is not None:
        img = pp.find("wpml:imageFormat", NS)
        if img is not None:
            pp.remove(img)

    # 4-1) センサー選択
    if sensor_modes:
        if pp is None:
            fld = tree.find(".//kml:Folder", NS)
            pp = etree.SubElement(fld, f"{{{NS['wpml']}}}payloadParam")
            etree.SubElement(pp, f"{{{NS['wpml']}}}payloadPositionIndex").text = "0"
        fmt = ",".join(m.lower() for m in sensor_modes)
        etree.SubElement(pp, f"{{{NS['wpml']}}}imageFormat").text = fmt

    # 4-2) 写真撮影
    if do_photo and not do_video:
        pms = [
            p for p in tree.findall(".//kml:Placemark", NS)
            if p.find("wpml:index", NS) is not None
        ]
        base = _max_id(tree.getroot(), ".//wpml:actionGroupId")
        for i, pm in enumerate(pms):
            idx = pm.find("wpml:index", NS).text
            grp, nid = _grp(pm, idx, base + i)
            act = etree.SubElement(grp, f"{{{NS['wpml']}}}action")
            etree.SubElement(act, f"{{{NS['wpml']}}}actionId").text = str(nid)
            etree.SubElement(act, f"{{{NS['wpml']}}}actionActuatorFunc").text = "orientedShoot"
            etree.SubElement(act, f"{{{NS['wpml']}}}actionActuatorFuncParam")

    # 4-3) 動画撮影
    if do_video:
        pms = sorted(
            [
                p for p in tree.findall(".//kml:Placemark", NS)
                if p.find("wpml:index", NS) is not None
            ],
            key=lambda p: int(p.find("wpml:index", NS).text)
        )
        if pms:
            first, last = pms[0], pms[-1]
            base = _max_id(tree.getroot(), ".//wpml:actionGroupId")

            # ジンバル回転（オプション化）
            if do_gimbal:
                grp, nid = _grp(first, first.find("wpml:index", NS).text, base)
                ag = etree.SubElement(grp, f"{{{NS['wpml']}}}action")
                etree.SubElement(ag, f"{{{NS['wpml']}}}actionId").text = str(nid)
                etree.SubElement(ag, f"{{{NS['wpml']}}}actionActuatorFunc").text = "gimbalRotate"
                param = etree.SubElement(ag, f"{{{NS['wpml']}}}actionActuatorFuncParam")
                add = lambda k, v: etree.SubElement(param, f"{{{NS['wpml']}}}{k}").__setattr__('text', str(v))
                add("gimbalRotateMode",        "absoluteAngle")
                add("gimbalPitchRotateEnable", 1)
                add("gimbalPitchRotateAngle",  -90)
                add("gimbalRollRotateEnable",  0)
                add("gimbalRollRotateAngle",   0)
                add("gimbalYawRotateEnable",   0)
                add("gimbalYawRotateAngle",    0)
                add("gimbalRotateTimeEnable",  0)
                add("gimbalRotateTime",        0)
                add("payloadPositionIndex",    0)

            # REC 開始
            rec_id = nid + 1 if do_gimbal else _grp(first, first.find("wpml:index", NS).text, base)[1]
            acs = etree.SubElement(grp, f"{{{NS['wpml']}}}action")
            etree.SubElement(acs, f"{{{NS['wpml']}}}actionId").text = str(rec_id)
            etree.SubElement(acs, f"{{{NS['wpml']}}}actionActuatorFunc").text = "startRecord"
            p = etree.SubElement(acs, f"{{{NS['wpml']}}}actionActuatorFuncParam")
            etree.SubElement(p, f"{{{NS['wpml']}}}fileSuffix").text = video_suffix
            etree.SubElement(p, f"{{{NS['wpml']}}}payloadPositionIndex").text = "0"

            # REC 停止（正しい ID 採番）
            grp2, stop_id = _grp(last, last.find("wpml:index", NS).text, base + 1)
            st = etree.SubElement(grp2, f"{{{NS['wpml']}}}action")
            etree.SubElement(st, f"{{{NS['wpml']}}}actionId").text = str(stop_id)
            etree.SubElement(st, f"{{{NS['wpml']}}}actionActuatorFunc").text = "stopRecord"
            sparam = etree.SubElement(st, f"{{{NS['wpml']}}}actionActuatorFuncParam")
            etree.SubElement(sparam, f"{{{NS['wpml']}}}payloadPositionIndex").text = "0"

    # 5) ヨー固定
    if yaw_fix and yaw_angle is not None:
        for hp in tree.findall(".//wpml:waypointHeadingParam", NS) + \
                  tree.findall(".//wpml:globalWaypointHeadingParam", NS):
            m = hp.find("wpml:waypointHeadingMode", NS)
            a = hp.find("wpml:waypointHeadingAngle", NS)
            if m is not None: m.text = "fixed"
            if a is not None: a.text = str(yaw_angle)

    # 6) 空の actionGroup 削除
    for ag in tree.findall(".//wpml:actionGroup", NS):
        if not ag.findall("wpml:action", NS):
            ag.getparent().remove(ag)

def process_kmz(path, offset, do_photo, do_video, video_suffix,
                do_gimbal, yaw_fix, yaw_angle, speed, sensor_modes, log):
    try:
        log.insert(tk.END, f"Extracting {os.path.basename(path)}...\n")
        wd = extract_kmz(path)
        kmls = glob.glob(os.path.join(wd, "**", "template.kml"), recursive=True)
        if not kmls:
            raise FileNotFoundError("template.kml が見つかりませんでした。")
        out_root, outdir = prepare_output_dirs(path, offset)
        for kml in kmls:
            log.insert(tk.END, f"Converting {os.path.basename(kml)}...\n")
            parser = etree.XMLParser(remove_blank_text=True)
            tree = etree.parse(kml, parser)
            convert_kml(tree, offset, do_photo, do_video, video_suffix,
                        do_gimbal, yaw_fix, yaw_angle, speed, sensor_modes)
            out_path = os.path.join(outdir, os.path.basename(kml))
            tree.write(out_path, encoding="utf-8",
                       pretty_print=True, xml_declaration=True)
        # "waylines.wpml" はコピーせず、"res" のみ
        for name in ["res"]:
            srcs = glob.glob(os.path.join(wd, "**", name), recursive=True)
            if srcs:
                src = srcs[0]
                dst = os.path.join(outdir, os.path.basename(src))
                if os.path.isdir(src):
                    if os.path.exists(dst):
                        shutil.rmtree(dst)
                    shutil.copytree(src, dst)
                else:
                    shutil.copy2(src, dst)
        out_kmz = repackage_to_kmz(out_root, path)
        log.insert(tk.END, f"Saved: {out_kmz}\nFinished\n\n")
        messagebox.showinfo("完了", f"変換完了:\n{out_kmz}")
    except Exception as e:
        messagebox.showerror("エラー", str(e))
        log.insert(tk.END, f"Error: {e}\n\n")
    finally:
        if os.path.exists("_kmz_work"):
            shutil.rmtree("_kmz_work")

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
        ttk.Spinbox(self, from_=1, to=15, textvariable=self.sp, width=5).grid(row=1, column=1, columnspan=2, sticky="w")

        self.ph = tk.BooleanVar(value=False)
        ttk.Checkbutton(self, text="写真撮影", variable=self.ph, command=self.update_ctrl).grid(row=2, column=0, sticky="w")
        self.vd = tk.BooleanVar(value=False)
        ttk.Checkbutton(self, text="動画撮影", variable=self.vd, command=self.update_ctrl).grid(row=2, column=1, sticky="w")
        self.vd_suffix_label = ttk.Label(self, text="動画ファイル名:")
        self.vd_suffix_var   = tk.StringVar(value="video_01")
        self.vd_suffix_entry = ttk.Entry(self, textvariable=self.vd_suffix_var, width=20)

        ttk.Label(self, text="センサー選択:").grid(row=3, column=0, sticky="w")
        self.sm_vars = {m: tk.BooleanVar(value=False) for m in SENSOR_MODES}
        for i, m in enumerate(SENSOR_MODES):
            ttk.Checkbutton(self, text=m, variable=self.sm_vars[m]).grid(row=3, column=1+i, sticky="w")

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
            self.he.config(state="normal")
            self.he.delete(0, tk.END)
            self.he.focus()
        else:
            self.he.config(state="disabled")
            self.he.delete(0, tk.END)

    def update_ctrl(self):
        if self.ph.get():
            self.vd.set(False)
        if self.vd.get():
            self.ph.set(False)
        if self.ph.get() or self.vd.get():
            self.gc.state(["disabled"])
            self.gm.set(True)
        else:
            self.gc.state(["!disabled"])
        if self.vd.get():
            self.vd_suffix_label.grid(row=2, column=2, sticky="e", padx=(10,2))
            self.vd_suffix_entry.grid(row=2, column=3, sticky="w")
        else:
            self.vd_suffix_label.grid_forget()
            self.vd_suffix_entry.grid_forget()

    def update_yaw(self, event=None):
        if self.yf.get():
            self.yc.grid(row=5, column=1, padx=5, columnspan=2, sticky="w")
            if not self.yc.get():
                self.yc.set(next(iter(YAW_OPTIONS)))
            if self.yc.get() == "手動入力":
                self.ye.config(state="normal")
                self.ye.grid(row=5, column=3)
            else:
                self.ye.config(state="disabled")
                self.ye.grid_forget()
        else:
            self.yc.grid_forget()
            self.ye.grid_forget()

    def get_params(self):
        offset = 0.0
        v = HEIGHT_OPTIONS.get(self.hc.get())
        if v == "custom":
            try:
                offset = float(self.he.get())
            except:
                pass
        else:
            offset = float(v)

        yaw_angle = None
        if self.yf.get():
            yval = YAW_OPTIONS.get(self.yc.get())
            if yval == "custom":
                try:
                    yaw_angle = float(self.ye.get())
                except:
                    pass
            else:
                yaw_angle = float(yval)

        return {
            "offset":       offset,
            "do_photo":     self.ph.get(),
            "do_video":     self.vd.get(),
            "video_suffix": self.vd_suffix_var.get(),
            "do_gimbal":    self.gm.get(),
            "yaw_fix":      self.yf.get(),
            "yaw_angle":    yaw_angle,
            "speed":        max(1, min(15, self.sp.get())),
            "sensor_modes": [m for m, var in self.sm_vars.items() if var.get()]
        }

def main():
    root = TkinterDnD.Tk()
    root.title("ATL→ASL 変換＋撮影制御ツール (ver. GUI24)")
    root.geometry("750x650")
    frm = ttk.Frame(root, padding=10)
    frm.pack(fill="both", expand=True)

    app = AppGUI(frm)
    app.pack(fill="x", pady=(0,10))

    drop = tk.Label(frm, text=".kmzをここにドロップ", bg="lightgray", width=70, height=5, relief=tk.RIDGE)
    drop.pack(pady=12, fill="x")
    drop.drop_target_register(DND_FILES)

    log_frame = ttk.LabelFrame(frm, text="ログ")
    log_frame.pack(fill="both", expand=True)
    log = scrolledtext.ScrolledText(log_frame, height=16)
    log.pack(fill="both", expand=True)

    def on_drop(event):
        path = event.data.strip("{}")
        if not path.lower().endswith(".kmz"):
            messagebox.showwarning("警告", ".kmzファイルのみ対応しています。")
            return
        params = app.get_params()
        params.update({"path": path, "log": log})
        threading.Thread(target=process_kmz, kwargs=params, daemon=True).start()

    drop.dnd_bind("<<Drop>>", on_drop)
    root.mainloop()

if __name__ == "__main__":
    main()
