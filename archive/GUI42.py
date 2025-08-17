#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
GUI42_FULL.py – ATL→ASL 変換＋撮影制御ツール（imageFormat 方式）
2025-07-31
────────────────────────────────────────────────────────
■ 主要機能（GUI41 完全継承＋新仕様）
 1. ATL(対地高度) ⇔ ASL(海抜高度) 変換
 2. Waypoint 編集・検証
 3. カメラ指定を <wpml:imageFormat>zoom,wide 方式へ変更
    └ Zoom / Wide / IR の複数チェック
 4. 撮影モード排他選択（写真＝orientedShoot / 動画＝startRecord＆stopRecord）
 5. ジンバル・ヨー固定制御（チェック時のみ追加）
 6. 既存 GUI41 機能は削除せず完全維持
────────────────────────────────────────────────────────
"""

import os, shutil, zipfile, math, threading, tkinter as tk
from tkinter import ttk, messagebox, scrolledtext
from tkinterdnd2 import TkinterDnD, DND_FILES
from lxml import etree
from datetime import datetime

# ───────────────────────────────────────────────────────
# 定数 / 名前空間
# ───────────────────────────────────────────────────────
NS = {"kml": "http://www.opengis.net/kml/2.2",
      "wpml": "http://www.dji.com/wpmz/1.0.6"}

# 参照点 (海抜標高[m]) – GUI41 既存
REFERENCE_POINTS = {
    "本部":   (136.55595225, 36.07295176, 612.2),
    "烏帽子": (136.55999999, 36.07499999, 962.02)
}

# imageFormat 選択肢
IMAGE_FORMATS = ["zoom", "wide", "ir"]

# 高度モード選択肢 – GUI41 既存
ALT_MODE = ["ASL (EGM96)", "ATL (離陸点基準)"]

# ───────────────────────────────────────────────────────
# 低水準ユーティリティ
# ───────────────────────────────────────────────────────
def extract_kmz(path, work_dir="_kmz_work"):
    if os.path.exists(work_dir):
        shutil.rmtree(work_dir)
    with zipfile.ZipFile(path, "r") as zf:
        zf.extractall(work_dir)
    return work_dir

def repackage_kmz(src_dir, out_path):
    tmp_zip = out_path + ".zip"
    with zipfile.ZipFile(tmp_zip, "w", zipfile.ZIP_DEFLATED) as zf:
        for root, _, files in os.walk(src_dir):
            for f in files:
                full = os.path.join(root, f)
                rel  = os.path.relpath(full, src_dir)
                zf.write(full, rel)
    if os.path.exists(out_path):
        os.remove(out_path)
    os.rename(tmp_zip, out_path)

def haversine(lon1, lat1, lon2, lat2):
    """緯経度(m) 距離計算 – GUI41 既存"""
    R = 6378137.0
    dlat, dlon = math.radians(lat2-lat1), math.radians(lon2-lon1)
    a = (math.sin(dlat/2)**2 +
         math.cos(math.radians(lat1))*math.cos(math.radians(lat2))*math.sin(dlon/2)**2)
    return 2*R*math.asin(math.sqrt(a))

# ───────────────────────────────────────────────────────
# ATL ⇔ ASL 変換ロジック（GUI41 完全維持）
# ───────────────────────────────────────────────────────
def atl_to_asl(ellipsoid_h, rel_alt, ref_alt):
    """
    DJI WPML:
      <wpml:ellipsoidHeight> = 楕円体高 (WGS84)
      <wpml:height>          = ATL (relativeToStartPoint)
    1. ATL を ASL(EGM96) に変換:
       ASL = ref_alt + rel_alt
    2. EGM96 → WGS84 楕円体補正:
       DJI Pilot2 は内部で (ellipsoidH - geoid_undulation) ≒ ASL を保持
       本ツールでは geoid_undulation ≒ 35m (日本平均) と仮定し近似
    """
    GEOID_UNDULATION = 35.0  # 地域ごとに差異あり – 必要に応じ UI に出す
    asl = ref_alt + rel_alt                      # 海抜
    new_ellipsoid = asl + GEOID_UNDULATION       # 楕円体高へ再格納
    return new_ellipsoid, asl                   # (楕円体高, ASL)

def asl_to_atl(ellipsoid_h, asl, ref_alt):
    GEOID_UNDULATION = 35.0
    atl = asl - ref_alt
    new_ellipsoid = asl + GEOID_UNDULATION
    return new_ellipsoid, atl

# ───────────────────────────────────────────────────────
# WPML アクション生成
# ───────────────────────────────────────────────────────
def new_action(func):
    act = etree.Element("{wpml}action".format(**NS))
    etree.SubElement(act, "{wpml}actionId".format(**NS)).text = "0"
    etree.SubElement(act, "{wpml}actionActuatorFunc".format(**NS)).text = func
    param = etree.SubElement(act, "{wpml}actionActuatorFuncParam".format(**NS))
    return act, param

def act_gimbal(angle):
    act, p = new_action("gimbalRotate")
    p_dict = {"gimbalRotateMode":"absoluteAngle",
              "gimbalPitchRotateEnable":"1",
              "gimbalPitchRotateAngle":str(angle),
              "gimbalRollRotateEnable":"0",
              "gimbalYawRotateEnable":"0"}
    for k,v in p_dict.items():
        etree.SubElement(p, "{wpml}{}".format(k)).text = v
    return act

def act_yaw(angle):
    act, p = new_action("rotateYaw")
    etree.SubElement(p, "{wpml}aircraftHeading".format(**NS)).text = str(angle)
    etree.SubElement(p, "{wpml}aircraftPathMode".format(**NS)).text = "counterClockwise"
    return act

def act_start_record():   return new_action("startRecord")[0]
def act_stop_record():    return new_action("stopRecord")[0]

# ───────────────────────────────────────────────────────
# KML 編集
# ───────────────────────────────────────────────────────
def set_image_format(tree, fmt_text):
    node = tree.find(".//wpml:payloadParam/wpml:imageFormat", NS)
    if node is None:
        pp = tree.find(".//wpml:payloadParam", NS)
        node = etree.SubElement(pp, "{wpml}imageFormat".format(**NS))
    node.text = fmt_text

def refresh_action_group(group):
    for a in list(group):
        if a.tag.endswith("action"):
            group.remove(a)

def reorganize_actions(group, idx, first_i, last_i, prm):
    refresh_action_group(group)
    seq = []
    # Yaw
    if prm["yaw_enable"] and prm["yaw_angle"] is not None:
        seq.append(act_yaw(prm["yaw_angle"]))
    # Gimbal
    if prm["gimbal_enable"] and prm["gimbal_ang"] is not None:
        seq.append(act_gimbal(prm["gimbal_ang"]))

    if prm["mode"] == "photo":
        # 既存 orientedShoot を復元 (GUI41 のまま)
        orig_os = prm["os_proto"]
        if prm["gimbal_enable"]:
            gp = orig_os.find(".//wpml:gimbalPitchRotateAngle", NS)
            if gp is not None:
                gp.text = str(prm["gimbal_ang"])
        seq.append(orig_os)
    else:  # video
        if idx == first_i:
            seq.append(act_start_record())
        if idx == last_i:
            seq.append(act_stop_record())

    # append & reindex
    for n,a in enumerate(seq):
        a.find("wpml:actionId", NS).text = str(n)
        group.append(a)

def convert_kml(tree, prm, logger):
    placemarks = tree.findall(".//kml:Placemark", NS)
    first_i = int(placemarks[0].find("wpml:index", NS).text)
    last_i  = int(placemarks[-1].find("wpml:index", NS).text)

    # imageFormat
    set_image_format(tree, prm["img_fmt"])

    # orientedShoot プロトタイプ取得 (最初の WP から)
    os_proto = None
    for act in placemarks[0].find(".//wpml:actionGroup", NS):
        fn = act.find("wpml:actionActuatorFunc", NS)
        if fn is not None and fn.text == "orientedShoot":
            os_proto = act
            break
    prm["os_proto"] = os_proto

    # 各 WP
    for pm in placemarks:
        idx = int(pm.find("wpml:index", NS).text)
        grp = pm.find(".//wpml:actionGroup", NS)
        if grp is None:
            grp = etree.SubElement(pm, "{wpml}actionGroup".format(**NS))
        reorganize_actions(grp, idx, first_i, last_i, prm)
        logger.append(f"WP{idx}: actions rebuilt")

    # ATL⇔ASL 変換
    if prm["alt_mode"] == "ASL":
        ref_alt = REFERENCE_POINTS[prm["ref"]][2]
        for pm in placemarks:
            elli = pm.find("wpml:ellipsoidHeight", NS)
            relh = pm.find("wpml:height", NS)
            new_elli, _ = atl_to_asl(float(elli.text), float(relh.text), ref_alt)
            elli.text = f"{new_elli:.6f}"
            # height(ATL) は変えない = Pilot が自動再計算
    else:  # ATL
        pass  # 変換不要

# ───────────────────────────────────────────────────────
# GUI
# ───────────────────────────────────────────────────────
class MainUI(ttk.Frame):
    def __init__(self, master, ctl):
        super().__init__(master, padding=8)
        self.ctl = ctl
        self.vars()
        self.widgets()
        self.grid(sticky="nsew")

    def vars(self):
        self.photo = tk.BooleanVar()
        self.video = tk.BooleanVar()

        self.img_v = {k: tk.BooleanVar() for k in IMAGE_FORMATS}

        self.alt_mode = tk.StringVar(value="ASL (EGM96)")
        self.alt_ref  = tk.StringVar(value="本部")

        self.gim_en  = tk.BooleanVar()
        self.gim_ang = tk.StringVar(value="-90")  # default

        self.yaw_en  = tk.BooleanVar()
        self.yaw_ang = tk.StringVar(value="88.00")

    def widgets(self):
        # モード選択
        fr_mode = ttk.LabelFrame(self, text="撮影モード")
        ttk.Checkbutton(fr_mode, text="写真", variable=self.photo,
                        command=self.ctl.exclusive).grid(row=0, column=0, sticky="w")
        ttk.Checkbutton(fr_mode, text="動画", variable=self.video,
                        command=self.ctl.exclusive).grid(row=0, column=1, sticky="w")
        fr_mode.grid(row=0, column=0, sticky="w")

        # imageFormat
        fr_img = ttk.LabelFrame(self, text="imageFormat (複数可)")
        for c,(k,v) in enumerate(self.img_v.items()):
            ttk.Checkbutton(fr_img, text=k.upper(), variable=v).grid(row=0, column=c, sticky="w")
        fr_img.grid(row=1, column=0, sticky="w", pady=4)

        # ALT/ASL
        fr_alt = ttk.LabelFrame(self, text="高度モード")
        ttk.Radiobutton(fr_alt, text="ASL (EGM96)", variable=self.alt_mode,
                        value="ASL (EGM96)").grid(row=0, column=0, sticky="w")
        ttk.Radiobutton(fr_alt, text="ATL (離陸点)", variable=self.alt_mode,
                        value="ATL (離陸点)").grid(row=0, column=1, sticky="w")
        ttk.Label(fr_alt, text="参照点:").grid(row=0, column=2, padx=(10,2))
        ttk.Combobox(fr_alt, textvariable=self.alt_ref,
                     values=list(REFERENCE_POINTS.keys()), width=8,
                     state="readonly").grid(row=0, column=3)
        fr_alt.grid(row=2, column=0, sticky="w", pady=4)

        # Gimbal / Yaw
        fr_att = ttk.LabelFrame(self, text="ジンバル / ヨー制御")
        ttk.Checkbutton(fr_att, text="ジンバルピッチ", variable=self.gim_en).grid(row=0, column=0, sticky="w")
        ttk.Entry(fr_att, textvariable=self.gim_ang, width=6).grid(row=0, column=1)
        ttk.Label(fr_att, text="deg").grid(row=0, column=2, padx=(0,6))

        ttk.Checkbutton(fr_att, text="ヨー固定", variable=self.yaw_en).grid(row=0, column=3, sticky="w")
        ttk.Entry(fr_att, textvariable=self.yaw_ang, width=6).grid(row=0, column=4)
        ttk.Label(fr_att, text="deg").grid(row=0, column=5)
        fr_att.grid(row=3, column=0, sticky="w", pady=4)

        # ドロップエリア
        self.drop = tk.Label(self, text=".kmz をドロップ", relief=tk.RIDGE,
                             bg="#f0f0f0", width=60, height=3)
        self.drop.grid(row=4, column=0, pady=6, sticky="we")

        # ログ
        self.log = scrolledtext.ScrolledText(self, height=18)
        self.log.grid(row=5, column=0, sticky="nsew", pady=4)

        # DnD
        self.drop.drop_target_register(DND_FILES)
        self.drop.dnd_bind("<<Drop>>", self.ctl.on_drop)

        # resize weights
        self.rowconfigure(5, weight=1)
        self.columnconfigure(0, weight=1)

# ───────────────────────────────────────────────────────
# Controller
# ───────────────────────────────────────────────────────
class Controller:
    def __init__(self, root):
        self.ui = MainUI(root, self)

    # 排他チェック
    def exclusive(self):
        if self.ui.photo.get():
            self.ui.video.set(False)
        elif self.ui.video.get():
            self.ui.photo.set(False)

    # ドロップイベント
    def on_drop(self, evt):
        kmz_path = evt.data.strip("{}")
        if not kmz_path.lower().endswith(".kmz"):
            messagebox.showwarning("注意", ".kmz ファイルのみ対応")
            return

        # パラメータ収集
        try:
            p = self.collect_params()
        except ValueError as e:
            messagebox.showerror("入力エラー", str(e))
            return
        threading.Thread(target=self.process_kmz, args=(kmz_path, p), daemon=True).start()

    def collect_params(self):
        prm = {}
        # mode
        if self.ui.photo.get():
            prm["mode"] = "photo"
        elif self.ui.video.get():
            prm["mode"] = "video"
        else:
            raise ValueError("写真か動画のいずれかを選択してください。")

        # imageFormat
        fmt = [k for k,v in self.ui.img_v.items() if v.get()]
        if not fmt:
            raise ValueError("少なくとも1つ imageFormat を選択")
        prm["img_fmt"] = ",".join(fmt)

        # Alt mode
        prm["alt_mode"] = "ASL" if self.ui.alt_mode.get().startswith("ASL") else "ATL"
        prm["ref"] = self.ui.alt_ref.get()

        # Gimbal / Yaw
        prm["gimbal_enable"] = self.ui.gim_en.get()
        prm["gimbal_ang"] = float(self.ui.gim_ang.get()) if prm["gimbal_enable"] else None
        prm["yaw_enable"] = self.ui.yaw_en.get()
        prm["yaw_angle"] = float(self.ui.yaw_ang.get()) if prm["yaw_enable"] else None
        return prm

    # KMZ 処理
    def process_kmz(self, kmz_path, prm):
        log = self.ui.log
        log.insert(tk.END, f"[{datetime.now():%H:%M:%S}] 開始: {os.path.basename(kmz_path)}\n")
        work = extract_kmz(kmz_path)

        # template.kml 探索
        kml_path = None
        for root,_,files in os.walk(work):
            for f in files:
                if f.lower().endswith(".kml"):
                    kml_path = os.path.join(root, f)
                    break
            if kml_path: break
        if not kml_path:
            messagebox.showerror("エラー", "KML が見つかりません")
            return

        tree = etree.parse(kml_path)
        buf = []
        convert_kml(tree, prm, buf)
        for ln in buf: log.insert(tk.END, ln+"\n")

        # 上書き保存
        tree.write(kml_path, encoding="utf-8",
                   pretty_print=True, xml_declaration=True)

        # 出力 KMZ
        out_kmz = os.path.splitext(kmz_path)[0] + "_converted.kmz"
        repackage_kmz(work, out_kmz)
        shutil.rmtree(work)

        log.insert(tk.END, f"--> 出力: {out_kmz}\n\n")
        messagebox.showinfo("完了", f"変換完了:\n{out_kmz}")

# ───────────────────────────────────────────────────────
def main():
    root = TkinterDnD.Tk()
    root.title("ATL→ASL 変換＋撮影制御ツール v4 (Full)")
    root.geometry("820x720")
    Controller(root)
    root.mainloop()

if __name__ == "__main__":
    main()
