import os
import shutil
import zipfile
import glob
import tkinter as tk
from tkinter import messagebox, ttk
from tkinterdnd2 import TkinterDnD, DND_FILES
from lxml import etree

# ------------------------- 1. 初期デフォルト値 ------------------------- #
DEFAULTS = {
    "BASE_HEIGHT": 613.5,                        # 高度補正[m]
    "NEW_HEIGHT_MODE": "als",                    # "als" or "relativeToStartPoint"
    "AUTO_FLIGHT_SPEED": 15,                     # 自動飛行速度[m/s]
    "TAKEOFF_SECURITY_HEIGHT": 20,               # 安全離陸高度[m]
    "TURN_MODE": "toPointAndStopWithDiscontinuityCurvature",  # 旋回モード
    "HOVER_SECONDS": 2,                          # ホバリング時間[秒]
    "GIMBAL_PITCH": -90,                         # ジンバルピッチ角度[°]
    "YAW_ANGLE": 87.37,                          # デフォルトヨー角度[°]
    "ZOOM_FACTOR": 5,                            # ズーム倍率
    "ENABLE_ZOOM": True,                         # ズームカメラ有効化
    "ENABLE_WIDE": False,                        # ワイドカメラ有効化
    "ENABLE_IR": False,                          # IR カメラ有効化
}

# 各カメラに対応するペイロードインデックス
CAMERA_INDEX = {"zoom": 0, "wide": 1, "ir": 2}

LABEL_FONT = ("Meiryo", 11)


# ------------------------- 2. KMZ 展開・編集・再パック関連関数 ------------------------- #
def convert_and_extract_kmz(kmz_path):
    if not kmz_path.lower().endswith(".kmz"):
        return False, ".kmz ファイルではありません"
    base_dir = os.path.dirname(kmz_path)
    output_dir = os.path.join(base_dir, "Converted")
    zip_path = os.path.splitext(kmz_path)[0] + ".zip"
    try:
        shutil.copy(kmz_path, zip_path)
        if os.path.isdir(output_dir):
            shutil.rmtree(output_dir)
        os.makedirs(output_dir)
        with zipfile.ZipFile(zip_path, "r") as zf:
            zf.extractall(output_dir)
        os.remove(zip_path)
        return True, output_dir
    except Exception as e:
        return False, f"展開失敗: {e}"

def repack_to_kmz_only(folder_path, original_kmz_path):
    new_kmz_name = os.path.splitext(os.path.basename(original_kmz_path))[0] + "_movie.kmz"
    new_kmz_path = os.path.join(folder_path, new_kmz_name)
    temp_zip = new_kmz_path.replace(".kmz", ".zip")
    with zipfile.ZipFile(temp_zip, "w", zipfile.ZIP_DEFLATED) as zf:
        for root_dir, _, files in os.walk(folder_path):
            for f in files:
                full = os.path.join(root_dir, f)
                rel = os.path.relpath(full, folder_path)
                zf.write(full, rel)
    if os.path.exists(new_kmz_path):
        os.remove(new_kmz_path)
    os.rename(temp_zip, new_kmz_path)
    for item in os.listdir(folder_path):
        p = os.path.join(folder_path, item)
        if not item.endswith(".kmz"):
            shutil.rmtree(p) if os.path.isdir(p) else os.remove(p)
    return new_kmz_path

def convert_kml_logic(kml_in, kml_out, cfg):
    ns = {"kml": "http://www.opengis.net/kml/2.2", "wpml": "http://www.dji.com/wpmz/1.0.6"}
    tree = etree.parse(kml_in)
    root = tree.getroot()
    folder = root.find(".//kml:Folder", ns)

    # 高度モード変更
    hm = root.find(".//wpml:waylineCoordinateSysParam/wpml:heightMode", ns)
    if hm is not None:
        hm.text = cfg["NEW_HEIGHT_MODE"]

    for pm in root.findall(".//kml:Placemark", ns):
        # 高度補正
        for tag in ("wpml:height", "wpml:ellipsoidHeight"):
            he = pm.find(tag, ns)
            if he is not None:
                try:
                    he.text = str(float(he.text) + cfg["BASE_HEIGHT"])
                except ValueError:
                    pass
        # 飛行速度設定
        sp = pm.find("wpml:waypointSpeed", ns)
        if sp is not None:
            sp.text = str(cfg["AUTO_FLIGHT_SPEED"])
        # ヘディング設定
        hp = pm.find("wpml:waypointHeadingParam", ns)
        if hp is None:
            hp = etree.SubElement(pm, "{http://www.dji.com/wpmz/1.0.6}waypointHeadingParam")
        for tag, val in [
            ("waypointHeadingMode", "followWayline"),
            ("waypointHeadingAngle", str(cfg["YAW_ANGLE"])),
            ("waypointPoiPoint", "0,0,0"),
            ("waypointHeadingPoiIndex", "0")
        ]:
            e = hp.find(f"wpml:{tag}", ns)
            if e is None:
                etree.SubElement(hp, f"{{http://www.dji.com/wpmz/1.0.6}}{tag}").text = val
            else:
                e.text = val

        # アクショングループクリア
        ag = pm.find(".//wpml:actionGroup", ns)
        if ag is None:
            ag = etree.SubElement(pm, "{http://www.dji.com/wpmz/1.0.6}actionGroup")
        for old in ag.findall("wpml:action", ns):
            ag.remove(old)

        # アクション追加
        def new_action(act_id, func, params):
            act = etree.SubElement(ag, "{http://www.dji.com/wpmz/1.0.6}action")
            etree.SubElement(act, "{http://www.dji.com/wpmz/1.0.6}actionId").text = str(act_id)
            etree.SubElement(act, "{http://www.dji.com/wpmz/1.0.6}actionActuatorFunc").text = func
            p = etree.SubElement(act, "{http://www.dji.com/wpmz/1.0.6}actionActuatorFuncParam")
            for k, v in params.items():
                etree.SubElement(p, f"{{http://www.dji.com/wpmz/1.0.6}}{k}").text = str(v)

        wp_idx = int(pm.findtext("wpml:actionGroup/actionGroupId", default="0", namespaces=ns))
        if wp_idx == 0:
            if cfg["ENABLE_ZOOM"]:
                new_action(998, "startRecordVideo", {
                    "payloadPositionIndex": CAMERA_INDEX["zoom"],
                    "videoType": "zoom", "zoomFactor": cfg["ZOOM_FACTOR"]
                })
            if cfg["ENABLE_WIDE"]:
                new_action(999, "startRecordVideo", {
                    "payloadPositionIndex": CAMERA_INDEX["wide"],
                    "videoType": "wide"
                })
            if cfg["ENABLE_IR"]:
                new_action(1000, "startRecordVideo", {
                    "payloadPositionIndex": CAMERA_INDEX["ir"],
                    "videoType": "ir"
                })

        # ジンバル＆ホバリング
        preferred_idx = (
            CAMERA_INDEX["zoom"] if cfg["ENABLE_ZOOM"]
            else CAMERA_INDEX["wide"] if cfg["ENABLE_WIDE"]
            else CAMERA_INDEX["ir"]
        )
        new_action(1001, "gimbalRotate", {
            "gimbalRotateMode": "absoluteAngle",
            "gimbalPitchRotateEnable": 1,
            "gimbalPitchRotateAngle": cfg["GIMBAL_PITCH"],
            "gimbalRollRotateEnable": 0,
            "gimbalYawRotateEnable": 0,
            "gimbalRotateTimeEnable": 0,
            "payloadPositionIndex": preferred_idx
        })
        new_action(1002, "stayForSeconds", {
            "stayTime": cfg["HOVER_SECONDS"],
            "gimbalPitchRotateAngle": cfg["GIMBAL_PITCH"],
            "aircraftHeading": cfg["YAW_ANGLE"]
        })

    # グローバル設定編集
    if folder is not None:
        fs = folder.find("wpml:autoFlightSpeed", ns)
        if fs is not None:
            fs.text = str(cfg["AUTO_FLIGHT_SPEED"])
        tkh = root.find(".//wpml:takeOffSecurityHeight", ns)
        if tkh is not None:
            tkh.text = str(cfg["TAKEOFF_SECURITY_HEIGHT"])
        tm = folder.find("wpml:globalWaypointTurnMode", ns)
        if tm is not None:
            tm.text = cfg["TURN_MODE"]
        gh = folder.find("wpml:globalWaypointHeadingParam", ns)
        if gh is not None:
            for tag, val in [
                ("waypointHeadingMode", "followWayline"),
                ("waypointHeadingAngle", str(cfg["YAW_ANGLE"]))
            ]:
                e = gh.find(f"wpml:{tag}", ns)
                if e is not None:
                    e.text = val

    tree.write(kml_out, encoding="utf-8", pretty_print=True, xml_declaration=True)
    os.remove(kml_in)

def process_kmz_and_edit_kml(kmz_path, cfg):
    ok, res = convert_and_extract_kmz(kmz_path)
    if not ok:
        return False, res
    work_dir = res
    kmls = glob.glob(os.path.join(work_dir, "**", "wpmz", "template.kml"), recursive=True)
    if not kmls:
        shutil.rmtree(work_dir)
        return False, "template.kml が見つかりません。"
    for src in kmls:
        dst = src.replace("template.kml", "template_movie.kml")
        convert_kml_logic(src, dst, cfg)
    new_kmz = repack_to_kmz_only(work_dir, kmz_path)
    return True, f"変換完了: {new_kmz}"


# ------------------------- 3. Tkinter GUI ------------------------- #
class SettingsPanel(tk.LabelFrame):
    def __init__(self, master):
        super().__init__(master, text="パラメータ設定", font=LABEL_FONT, padx=8, pady=4)
        self.entries = {}
        self.booleans = {}
        self.height_values = {
            "613.5 - 事務所前": 613.5,
            "962.02 - 烏帽子": 962.02,
            "その他 - 手動入力": "custom"
        }
        # ヨー角「1Q」「2Q」「4Q」
        self.yaw_values = {
            "1Q: 87.37°": 87.37,
            "2Q: 96.92°": 96.92,
            "4Q: 87.31°": 87.31,
            "その他 - 手動入力": "custom"
        }
        self._build()

    def _build(self):
        r = 0
        # 基準高度オフセット
        tk.Label(self, text="基準高度オフセット (m)", anchor="w").grid(row=r, column=0, sticky="w")
        self.height_combo = ttk.Combobox(self, width=22, values=list(self.height_values.keys()), state="readonly")
        self.height_combo.set("613.5 - 事務所前")
        self.height_combo.bind("<<ComboboxSelected>>", self._on_height_selected)
        self.height_combo.grid(row=r, column=1, sticky="w", padx=2, pady=1)
        r += 1

        tk.Label(self, text="手動入力値 (m)", anchor="w").grid(row=r, column=0, sticky="w")
        self.height_entry = ttk.Entry(self, width=10, state="disabled")
        self.height_entry.grid(row=r, column=1, sticky="w", padx=2, pady=1)
        r += 1

        # 飛行パラメータ
        for key, lbl in [
            ("AUTO_FLIGHT_SPEED", "飛行速度 (m/s)"),
            ("TAKEOFF_SECURITY_HEIGHT", "安全離陸高度 (m)"),
            ("HOVER_SECONDS", "ホバリング時間 (s)"),
            ("GIMBAL_PITCH", "ジンバルピッチ (°)"),
            ("ZOOM_FACTOR", "ズーム倍率"),
        ]:
            tk.Label(self, text=lbl, anchor="w").grid(row=r, column=0, sticky="w")
            e = ttk.Entry(self, width=10)
            e.insert(0, str(DEFAULTS[key]))
            e.grid(row=r, column=1, sticky="w", padx=2, pady=1)
            self.entries[key] = e
            r += 1

        # ヨー角度選択（1Q/2Q/4Q/その他）
        tk.Label(self, text="ヨー角度選択", anchor="w").grid(row=r, column=0, sticky="w")
        self.yaw_combo = ttk.Combobox(self, width=22, values=list(self.yaw_values.keys()), state="readonly")
        self.yaw_combo.set("1Q: 87.37°")
        self.yaw_combo.bind("<<ComboboxSelected>>", self._on_yaw_selected)
        self.yaw_combo.grid(row=r, column=1, sticky="w", padx=2, pady=1)
        r += 1

        tk.Label(self, text="手動入力角度 (°)", anchor="w").grid(row=r, column=0, sticky="w")
        self.yaw_entry = ttk.Entry(self, width=10, state="disabled")
        self.yaw_entry.grid(row=r, column=1, sticky="w", padx=2, pady=1)
        r += 1

        # 高度モード
        tk.Label(self, text="高度モード", anchor="w").grid(row=r, column=0, sticky="w")
        self.height_mode = tk.StringVar(value=DEFAULTS["NEW_HEIGHT_MODE"])
        om = ttk.OptionMenu(self, self.height_mode, DEFAULTS["NEW_HEIGHT_MODE"], "als", "relativeToStartPoint")
        om.grid(row=r, column=1, sticky="w", padx=2, pady=1)
        r += 1

        # カメラチェックボックス
        for key, lbl in [
            ("ENABLE_ZOOM", "ズームカメラ有効"),
            ("ENABLE_WIDE", "ワイドカメラ有効"),
            ("ENABLE_IR",   "IR カメラ有効")
        ]:
            var = tk.BooleanVar(value=DEFAULTS[key])
            cb = ttk.Checkbutton(self, text=lbl, variable=var)
            cb.grid(row=r, column=0, columnspan=2, sticky="w", padx=2, pady=1)
            self.booleans[key] = var
            r += 1

    def _on_height_selected(self, _):
        if self.height_combo.get() == "その他 - 手動入力":
            self.height_entry.config(state="normal")
            self.height_entry.delete(0, tk.END)
            self.height_entry.insert(0, str(DEFAULTS["BASE_HEIGHT"]))
            self.height_entry.focus()
        else:
            self.height_entry.config(state="disabled")
            self.height_entry.delete(0, tk.END)

    def _on_yaw_selected(self, _):
        if self.yaw_combo.get() == "その他 - 手動入力":
            self.yaw_entry.config(state="normal")
            self.yaw_entry.delete(0, tk.END)
            self.yaw_entry.insert(0, str(DEFAULTS["YAW_ANGLE"]))
            self.yaw_entry.focus()
        else:
            self.yaw_entry.config(state="disabled")
            self.yaw_entry.delete(0, tk.END)

    def read_config(self):
        cfg = {}
        # 基準高度
        sel_h = self.height_combo.get()
        if sel_h == "その他 - 手動入力":
            try:
                cfg["BASE_HEIGHT"] = float(self.height_entry.get())
            except ValueError:
                cfg["BASE_HEIGHT"] = DEFAULTS["BASE_HEIGHT"]
        else:
            cfg["BASE_HEIGHT"] = self.height_values[sel_h]
        # 数値パラメータ
        for k, ent in self.entries.items():
            v = ent.get()
            try:
                cfg[k] = float(v) if "." in v else int(v)
            except ValueError:
                cfg[k] = v
        cfg["NEW_HEIGHT_MODE"] = self.height_mode.get()
        # ヨー角度
        sel_y = self.yaw_combo.get()
        if sel_y == "その他 - 手動入力":
            try:
                cfg["YAW_ANGLE"] = float(self.yaw_entry.get())
            except ValueError:
                cfg["YAW_ANGLE"] = DEFAULTS["YAW_ANGLE"]
        else:
            cfg["YAW_ANGLE"] = self.yaw_values[sel_y]
        # カメラ有効化
        for k, var in self.booleans.items():
            cfg[k] = var.get()
        cfg["TURN_MODE"] = DEFAULTS["TURN_MODE"]
        return cfg

# ------------------------- 4. ドラッグ＆ドロップ処理 & GUI 初期化 ------------------------- #
def drop_file(event):
    paths = event.data.strip().split()
    if len(paths) != 1:
        messagebox.showwarning("エラー", "ファイルは1つだけドロップしてください。", parent=root)
        return
    path = paths[0].strip("{}")
    label_drop.config(text=f"ファイル：{os.path.basename(path)}")
    if not path.lower().endswith(".kmz"):
        messagebox.showinfo("情報", ".kmz ファイルのみ対応", parent=root)
        return
    cfg = panel.read_config()
    ok, msg = process_kmz_and_edit_kml(path, cfg)
    if ok:
        messagebox.showinfo("結果", msg, parent=root)
    else:
        messagebox.showerror("エラー", msg, parent=root)

root = TkinterDnD.Tk()
root.title("KMZ 動画ミッション変換ツール")
root.geometry("550x620")

panel = SettingsPanel(root)
panel.pack(fill="x", padx=14, pady=6)

label_drop = tk.Label(
    root,
    text="ここに .kmz ファイルをドラッグ＆ドロップ",
    bg="lightgray",
    width=64,
    height=10,
    anchor="nw",
    justify="left",
    font=LABEL_FONT
)
label_drop.pack(padx=14, pady=12, fill="both", expand=True)
label_drop.drop_target_register(DND_FILES)
label_drop.dnd_bind("<<Drop>>", drop_file)

root.mainloop()
