#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
convert_height_gui_asl_no_offset.py (ver. GUI61)

修正点:
• ウェイポイント間の移動時の機体ヘディング制御機能を追加
• 「撮影方向に合わせる」「飛行経路に従う」「元の設定を維持」の選択肢
• Global設定とLocal設定の両方で適切に制御
• オフセット補正機能を削除し、それ以外の機能は維持

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
from datetime import datetime
import pyperclip
import math

# --- 定数 ------------------------------------------------------------------

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
    "元の角度維持": "original",
    "フリー": "free",
    "1Q: 88.00°": 88.00,
    "2Q: 96.92°": 96.92,
    "4Q: 87.31°": 87.31,
    "北": 0.00,
    "手動入力": "custom"
}

GIMBAL_PITCH_OPTIONS = {
    "元の角度維持": "original",
    "真下: -90°": -90.0,
    "前: 0°": 0.0,
    "フリー": "free",  # 追加
    "手動入力": "custom"
}

ZOOM_RATIO_OPTIONS = {
    "元の倍率を維持": "original",
    "5倍": 5.0,
    "10倍": 10.0,
    "フリー": "free",  # 追加
    "手動入力": "custom"
}

# 機体ヘディング制御の選択肢
HEADING_MODE_OPTIONS = {
    "元の設定を維持": "original",
    "フリー": "manually",
    "次の撮影方向を向く": "follow_gimbal",
    "飛行方向を向く": "follow_wayline"
}

SENSOR_MODES = ["Wide", "Zoom", "IR"]

# --- ジンバル・ズーム情報取得 ---------------------------------------------

def extract_original_gimbal_angles(tree):
    """
    KMLツリーから、各ウェイポイントのジンバルピッチ／ヨー／機体ヘディング／焦点距離を取得。
    - orientedShoot → gimbalRotate → rotateYaw → zoom アクションの順で情報を収集し、
    - 最後に payloadParam の zoom 設定をフォールバックで補う。
    """
    original = {}

    for pm in tree.findall(".//kml:Placemark", NS):
        idx_elem = pm.find("wpml:index", NS)
        if idx_elem is None:
            continue
        idx = int(idx_elem.text)
        info = {}

        # 1) orientedShoot から pitch/yaw/heading/focalLength
        for action in pm.findall(".//wpml:action", NS):
            if action.findtext("wpml:actionActuatorFunc", namespaces=NS) == "orientedShoot":
                p = action.findtext(".//wpml:gimbalPitchRotateAngle", namespaces=NS)
                y = action.findtext(".//wpml:gimbalYawRotateAngle",   namespaces=NS)
                h = action.findtext(".//wpml:aircraftHeading",         namespaces=NS)
                f = action.findtext(".//wpml:focalLength",             namespaces=NS)
                if p: info["pitch"]   = float(p)
                if y: info["yaw"]     = float(y)
                if h: info["heading"] = float(h)
                if f:
                    fl = float(f)
                    info["focal_length"] = fl
                    info["zoom_ratio"]   = fl / 24.0
                break

        # 2) gimbalRotate から pitch/yaw （フォールバック）
        if "pitch" not in info or "yaw" not in info:
            for action in pm.findall(".//wpml:action", NS):
                if action.findtext("wpml:actionActuatorFunc", namespaces=NS) == "gimbalRotate":
                    param = action.find("wpml:actionActuatorFuncParam", NS)
                    if param is None:
                        continue
                    if param.findtext("wpml:gimbalPitchRotateEnable", namespaces=NS) == "1" and "pitch" not in info:
                        ap = param.findtext("wpml:gimbalPitchRotateAngle", namespaces=NS)
                        if ap: info["pitch"] = float(ap)
                    if param.findtext("wpml:gimbalYawRotateEnable", namespaces=NS) == "1" and "yaw" not in info:
                        ay = param.findtext("wpml:gimbalYawRotateAngle", namespaces=NS)
                        if ay: info["yaw"] = float(ay)
                    if "pitch" in info and "yaw" in info:
                        break

        # 3) rotateYaw から heading （フォールバック）
        if "heading" not in info:
            for action in pm.findall(".//wpml:action", NS):
                if action.findtext("wpml:actionActuatorFunc", namespaces=NS) == "rotateYaw":
                    h = action.findtext(".//wpml:aircraftHeading", namespaces=NS)
                    if h:
                        info["heading"] = float(h)
                        break

        # 4) zoom アクションから focalLength （フォールバック）
        if "zoom_ratio" not in info:
            for action in pm.findall(".//wpml:action", NS):
                if action.findtext("wpml:actionActuatorFunc", namespaces=NS) == "zoom":
                    f = action.findtext(".//wpml:focalLength", namespaces=NS)
                    if f:
                        fl = float(f)
                        info["focal_length"] = fl
                        info["zoom_ratio"]   = fl / 24.0
                    break

        # 5) payloadParam に zoom 指定がある場合の最終フォールバック
        fmt = tree.findtext(".//wpml:payloadParam/wpml:imageFormat", namespaces=NS) or ""
        if "zoom" in fmt.lower() and "zoom_ratio" not in info:
            info["zoom_ratio"] = None

        if info:
            original[idx] = info

    return original

def extract_original_heading_settings(tree):
    """元のヘディング設定を取得"""
    heading_settings = {}
    
    # グローバル設定取得
    global_heading = tree.find(".//wpml:globalWaypointHeadingParam", NS)
    if global_heading is not None:
        mode_elem = global_heading.find("wpml:waypointHeadingMode", NS)
        angle_elem = global_heading.find("wpml:waypointHeadingAngle", NS)
        poi_elem = global_heading.find("wpml:waypointPoiPoint", NS)
        poi_idx_elem = global_heading.find("wpml:waypointHeadingPoiIndex", NS)
        
        heading_settings["global"] = {
            "mode": mode_elem.text if mode_elem is not None else "followWayline",
            "angle": float(angle_elem.text) if angle_elem is not None else 0,
            "poi_point": poi_elem.text if poi_elem is not None else "0.000000,0.000000,0.000000",
            "poi_index": int(poi_idx_elem.text) if poi_idx_elem is not None else 0
        }
    else:
        heading_settings["global"] = {
            "mode": "followWayline",
            "angle": 0,
            "poi_point": "0.000000,0.000000,0.000000",
            "poi_index": 0
        }
    
    # 各ウェイポイント設定取得
    for pm in tree.findall(".//kml:Placemark", NS):
        idx_elem = pm.find("wpml:index", NS)
        if idx_elem is None:
            continue
        idx = int(idx_elem.text)
        
        heading_param = pm.find("wpml:waypointHeadingParam", NS)
        if heading_param is not None:
            mode_elem = heading_param.find("wpml:waypointHeadingMode", NS)
            angle_elem = heading_param.find("wpml:waypointHeadingAngle", NS)
            poi_elem = heading_param.find("wpml:waypointPoiPoint", NS)
            poi_idx_elem = heading_param.find("wpml:waypointHeadingPoiIndex", NS)
            path_mode_elem = heading_param.find("wpml:waypointHeadingPathMode", NS)
            
            heading_settings[idx] = {
                "mode": mode_elem.text if mode_elem is not None else None,
                "angle": float(angle_elem.text) if angle_elem is not None else None,
                "poi_point": poi_elem.text if poi_elem is not None else None,
                "poi_index": int(poi_idx_elem.text) if poi_idx_elem is not None else None,
                "path_mode": path_mode_elem.text if path_mode_elem is not None else "followBadArc"
            }
    
    return heading_settings

def calculate_gimbal_heading_direction(gimbal_yaw, aircraft_heading):
    """ジンバルヨー角と機体ヘディングから撮影方向を計算"""
    # ジンバルヨーは機体基準での角度、機体ヘディングは北基準
    # 撮影方向 = 機体ヘディング + ジンバルヨー
    direction = aircraft_heading + gimbal_yaw
    
    # -180 ~ 180 の範囲に正規化
    while direction > 180:
        direction -= 360
    while direction <= -180:
        direction += 360
    
    return direction

def get_next_waypoint_shooting_direction(waypoints_info, current_idx, original_angles):
    """次のウェイポイントの撮影方向を取得"""
    next_idx = current_idx + 1
    if next_idx in original_angles:
        next_gimbal = original_angles[next_idx]
        if "yaw" in next_gimbal and "heading" in next_gimbal:
            return calculate_gimbal_heading_direction(
                next_gimbal["yaw"], 
                next_gimbal["heading"]
            )
    return None

def zoom_ratio_to_focal_length(ratio):
    return ratio * 24.0

# --- KMZ ユーティリティ -----------------------------------------------------
def extract_kmz(path, work_dir="_kmz_work"):
    if os.path.exists(work_dir):
        shutil.rmtree(work_dir)
    os.makedirs(work_dir)
    
    with zipfile.ZipFile(path, "r") as zf:
        zf.extractall(work_dir)
    
    return work_dir

def prepare_output_dirs(input_kmz, do_photo, do_video, sensor_modes):
    base = os.path.splitext(os.path.basename(input_kmz))[0]
    
    if do_photo:
        mode_suffix = "Photo"
    elif do_video:
        mode_suffix = "Video"
    else:
        mode_suffix = "None"
    
    sensor_suffix = ""
    if sensor_modes:
        chars = []
        for s in sensor_modes:
            if s == "Wide": chars.append("W")
            elif s == "Zoom": chars.append("Z")
            elif s == "IR": chars.append("IR")
        sensor_suffix = "".join(chars)
    
    if sensor_suffix:
        out_root = os.path.join(os.path.dirname(input_kmz),
                               f"{base}_{mode_suffix}_{sensor_suffix}")
    else:
        out_root = os.path.join(os.path.dirname(input_kmz),
                               f"{base}_{mode_suffix}")
    
    if os.path.exists(out_root):
        shutil.rmtree(out_root)
    os.makedirs(out_root)
    
    wpmz_dir = os.path.join(out_root, "wpmz")
    os.makedirs(wpmz_dir)
    
    return out_root, wpmz_dir

def repackage_to_kmz(out_root, input_kmz, do_photo, do_video, sensor_modes):
    base = os.path.splitext(os.path.basename(input_kmz))[0]
    
    if do_photo:
        mode_suffix = "Photo"
    elif do_video:
        mode_suffix = "Video"
    else:
        mode_suffix = "None"
    
    sensor_suffix = ""
    if sensor_modes:
        chars = []
        for s in sensor_modes:
            if s == "Wide": chars.append("W")
            elif s == "Zoom": chars.append("Z")
            elif s == "IR": chars.append("IR")
        sensor_suffix = "".join(chars)
    
    if sensor_suffix:
        out_kmz = os.path.join(os.path.dirname(out_root),
                              f"{base}_{mode_suffix}_{sensor_suffix}.kmz")
    else:
        out_kmz = os.path.join(os.path.dirname(out_root),
                              f"{base}_{mode_suffix}.kmz")
    
    tmp = out_kmz + ".zip"
    
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zf:
        for root_dir, _, files in os.walk(out_root):
            for f in files:
                if f.lower().endswith(".wpml"):
                    continue
                full = os.path.join(root_dir, f)
                rel = os.path.relpath(full, out_root)
                zf.write(full, rel)
    
    if os.path.exists(out_kmz):
        os.remove(out_kmz)
    os.rename(tmp, out_kmz)
    
    return out_kmz

# --- KML 変換 ---------------------------------------------------------------

def create_gimbal_yaw_action(group, yaw_angle):
    act = etree.SubElement(group, f"{{{NS['wpml']}}}action")
    etree.SubElement(act, f"{{{NS['wpml']}}}actionId").text = "0"
    etree.SubElement(act, f"{{{NS['wpml']}}}actionActuatorFunc").text = "gimbalRotate"
    
    p = etree.SubElement(act, f"{{{NS['wpml']}}}actionActuatorFuncParam")
    etree.SubElement(p, f"{{{NS['wpml']}}}gimbalRotateMode").text = "absoluteAngle"
    etree.SubElement(p, f"{{{NS['wpml']}}}gimbalPitchRotateEnable").text = "0"
    etree.SubElement(p, f"{{{NS['wpml']}}}gimbalPitchRotateAngle").text = "0"
    etree.SubElement(p, f"{{{NS['wpml']}}}gimbalRollRotateEnable").text = "0"
    etree.SubElement(p, f"{{{NS['wpml']}}}gimbalRollRotateAngle").text = "0"
    etree.SubElement(p, f"{{{NS['wpml']}}}gimbalYawRotateEnable").text = "1"
    etree.SubElement(p, f"{{{NS['wpml']}}}gimbalYawRotateAngle").text = str(yaw_angle)
    etree.SubElement(p, f"{{{NS['wpml']}}}gimbalRotateTimeEnable").text = "0"
    etree.SubElement(p, f"{{{NS['wpml']}}}gimbalRotateTime").text = "0"
    etree.SubElement(p, f"{{{NS['wpml']}}}payloadPositionIndex").text = "0"

def apply_heading_settings(tree, heading_mode, original_heading_settings, original_angles, log):
    """ヘディング設定を適用"""
    log.insert(tk.END, f"\n=== 機体ヘディング制御設定 ===\n")
    log.insert(tk.END, f"制御モード: {heading_mode}\n")
    
    # ウェイポイントリスト取得
    waypoints = sorted(
        [p for p in tree.findall(".//kml:Placemark", NS) if p.find("wpml:index", NS) is not None],
        key=lambda x: int(x.find("wpml:index", NS).text)
    )
    
    if heading_mode == "follow_wayline":
        # 飛行経路に従う
        log.insert(tk.END, "飛行経路に従う設定を適用\n")
        
        # グローバル設定
        for hp in tree.findall(".//wpml:globalWaypointHeadingParam", NS):
            m = hp.find("wpml:waypointHeadingMode", NS)
            if m is not None:
                m.text = "followWayline"
        
        # 各ウェイポイントのローカル設定をクリア（グローバル設定を使用）
        for pm in waypoints:
            hp = pm.find("wpml:waypointHeadingParam", NS)
            if hp is not None:
                pm.remove(hp)
    
    elif heading_mode == "original":
        # 元の設定を維持
        log.insert(tk.END, "元のヘディング設定を維持\n")
        
        # グローバル設定復元
        global_settings = original_heading_settings.get("global", {})
        for hp in tree.findall(".//wpml:globalWaypointHeadingParam", NS):
            m = hp.find("wpml:waypointHeadingMode", NS)
            a = hp.find("wpml:waypointHeadingAngle", NS)
            p = hp.find("wpml:waypointPoiPoint", NS)
            pi = hp.find("wpml:waypointHeadingPoiIndex", NS)
            
            if m is not None:
                m.text = global_settings.get("mode", "followWayline")
            if a is not None:
                a.text = str(int(global_settings.get("angle", 0)))
            if p is not None:
                p.text = global_settings.get("poi_point", "0.000000,0.000000,0.000000")
            if pi is not None:
                pi.text = str(global_settings.get("poi_index", 0))
        
        # ローカル設定復元
        for pm in waypoints:
            idx = int(pm.find("wpml:index", NS).text)
            if idx in original_heading_settings:
                local_settings = original_heading_settings[idx]
                
                hp = pm.find("wpml:waypointHeadingParam", NS)
                if hp is None:
                    hp = etree.SubElement(pm, f"{{{NS['wpml']}}}waypointHeadingParam")
                else:
                    # 既存の設定をクリア
                    for child in list(hp):
                        hp.remove(child)
                
                # ローカル設定を適用（Noneでない値のみ）
                if local_settings.get("mode") is not None:
                    etree.SubElement(hp, f"{{{NS['wpml']}}}waypointHeadingMode").text = local_settings["mode"]
                if local_settings.get("angle") is not None:
                    etree.SubElement(hp, f"{{{NS['wpml']}}}waypointHeadingAngle").text = str(int(local_settings["angle"]))
                if local_settings.get("poi_point") is not None:
                    etree.SubElement(hp, f"{{{NS['wpml']}}}waypointPoiPoint").text = local_settings["poi_point"]
                if local_settings.get("path_mode") is not None:
                    etree.SubElement(hp, f"{{{NS['wpml']}}}waypointHeadingPathMode").text = local_settings["path_mode"]
                if local_settings.get("poi_index") is not None:
                    etree.SubElement(hp, f"{{{NS['wpml']}}}waypointHeadingPoiIndex").text = str(local_settings["poi_index"])
    
    elif heading_mode == "follow_gimbal":
        # 撮影方向に合わせる
        log.insert(tk.END, "撮影方向に合わせる設定を適用\n")
        
        # グローバル設定を固定モードに設定
        for hp in tree.findall(".//wpml:globalWaypointHeadingParam", NS):
            m = hp.find("wpml:waypointHeadingMode", NS)
            if m is not None:
                m.text = "fixed"
        
        # 各ウェイポイントで次のウェイポイントの撮影方向を設定
        for i, pm in enumerate(waypoints):
            idx = int(pm.find("wpml:index", NS).text)
            
            # 次のウェイポイントの撮影方向を取得
            shooting_direction = get_next_waypoint_shooting_direction(waypoints, i, original_angles)
            
            if shooting_direction is not None:
                # ローカル設定で撮影方向を設定
                hp = pm.find("wpml:waypointHeadingParam", NS)
                if hp is None:
                    hp = etree.SubElement(pm, f"{{{NS['wpml']}}}waypointHeadingParam")
                else:
                    # 既存設定をクリア
                    for child in list(hp):
                        hp.remove(child)
                
                etree.SubElement(hp, f"{{{NS['wpml']}}}waypointHeadingMode").text = "fixed"
                etree.SubElement(hp, f"{{{NS['wpml']}}}waypointHeadingAngle").text = str(int(shooting_direction))
                etree.SubElement(hp, f"{{{NS['wpml']}}}waypointPoiPoint").text = "0.000000,0.000000,0.000000"
                etree.SubElement(hp, f"{{{NS['wpml']}}}waypointHeadingPathMode").text = "followBadArc"
                etree.SubElement(hp, f"{{{NS['wpml']}}}waypointHeadingPoiIndex").text = "0"
                
                log.insert(tk.END, f"[WP {idx}] 次の撮影方向: {shooting_direction:.1f}°\n")
            else:
                # 撮影方向が取得できない場合はfollowWaylineを使用
                hp = pm.find("wpml:waypointHeadingParam", NS)
                if hp is None:
                    hp = etree.SubElement(pm, f"{{{NS['wpml']}}}waypointHeadingParam")
                else:
                    for child in list(hp):
                        hp.remove(child)
                
                etree.SubElement(hp, f"{{{NS['wpml']}}}waypointHeadingMode").text = "followWayline"
                etree.SubElement(hp, f"{{{NS['wpml']}}}waypointHeadingAngle").text = "0"
                etree.SubElement(hp, f"{{{NS['wpml']}}}waypointPoiPoint").text = "0.000000,0.000000,0.000000"
                etree.SubElement(hp, f"{{{NS['wpml']}}}waypointHeadingPathMode").text = "followBadArc"
                etree.SubElement(hp, f"{{{NS['wpml']}}}waypointHeadingPoiIndex").text = "0"
                
                log.insert(tk.END, f"[WP {idx}] 撮影方向取得不可→経路に従う\n")

    elif heading_mode == "manually":
        # 手動モード
        log.insert(tk.END, "手動モード設定を適用\n")
        
        # グローバル設定をmanuallyに設定
        for hp in tree.findall(".//wpml:globalWaypointHeadingParam", NS):
            m = hp.find("wpml:waypointHeadingMode", NS)
            if m is not None:
                m.text = "manually"
            # 必要に応じて angle や POI もデフォルト0／空に
            a = hp.find("wpml:waypointHeadingAngle", NS)
            if a is not None:
                a.text = "0"
        
        # 各ウェイポイントはグローバル設定を使用（local 設定は削除）
        for pm in waypoints:
            local_hp = pm.find("wpml:waypointHeadingParam", NS)
            if local_hp is not None:
                pm.remove(local_hp)
        
        log.insert(tk.END, "手動モード：グローバルmanuallyを適用し、各WPではlocal設定をクリア\n")

    log.insert(tk.END, "=== 機体ヘディング制御設定完了 ===\n\n")
    log.see(tk.END)

def convert_kml(tree,
                do_photo, do_video,
                do_gimbal, gimbal_pitch_angle, gimbal_pitch_mode,
                yaw_angle, yaw_mode, speed,
                sensor_modes, hover_time,
                zoom_ratio, zoom_mode,
                original_angles, heading_mode, original_heading_settings, log, wp_stop_mode):

    # 1) グローバル WP 停止モード設定
    gw_turn = tree.find(".//wpml:globalWaypointTurnMode", NS)
    if gw_turn is not None:
        if wp_stop_mode == "stop":
            gw_turn.text = "toPointAndStopWithDiscontinuityCurvature"
        else:
            gw_turn.text = "coordinateTurn"
        log.insert(tk.END, f"globalWaypointTurnMode → {gw_turn.text}\n")
    
    # グローバル高度モード確認
    global_height_mode_elem = tree.find(".//wpml:waylineCoordinateSysParam/wpml:heightMode", NS)
    global_height_mode = global_height_mode_elem.text if global_height_mode_elem is not None else "ATL"
    
    log.insert(tk.END, "\n=== 高度補正なし処理 ===\n")
    log.insert(tk.END, f"グローバル高度モード: {global_height_mode}\n")
    log.see(tk.END)
    
    processed_count = 0
    skipped_count = 0
    
    for pm in tree.findall(".//kml:Placemark", NS):
        idx_elem = pm.find("wpml:index", NS)
        idx = int(idx_elem.text) if idx_elem is not None else "不明"
        
        height_mode_elem = pm.find("wpml:heightMode", NS)
        height_mode = height_mode_elem.text if height_mode_elem is not None else global_height_mode
        
        # ASL判定
        is_asl = height_mode in ["ASL", "EGM96", "absoluteHeight", "WGS84"]
        
        for tag in ("height", "ellipsoidHeight"):
            el = pm.find(f"wpml:{tag}", NS)
            if el is not None and el.text:
                try:
                    original_height = float(el.text)
                    log.insert(tk.END, f"[WP {idx}] 標高モード: {height_mode}, 高度維持: {original_height}m\n")
                    if is_asl:
                        skipped_count += 1
                    else:
                        processed_count += 1
                except ValueError:
                    log.insert(tk.END, f"[WP {idx}] 高度値解析エラー: {el.text}\n")
    
    log.insert(tk.END, f"処理サマリー: 処理点 {processed_count}点, ASLスキップ {skipped_count}点\n")
    log.see(tk.END)
    
    # 高度モードをEGM96に統一
    for hm in tree.findall(".//wpml:heightMode", NS):
        if hm.text not in ["ASL", "EGM96", "absoluteHeight", "WGS84"]:
            old = hm.text
            hm.text = "EGM96"
            log.insert(tk.END, f"高度モード変更: {old} → EGM96\n")
    
    if global_height_mode_elem is not None and global_height_mode_elem.text not in ["ASL", "EGM96", "absoluteHeight", "WGS84"]:
        old = global_height_mode_elem.text
        global_height_mode_elem.text = "EGM96"
        log.insert(tk.END, f"グローバル高度モード変更: {old} → EGM96\n")
    
    # 機体ヘディング制御設定を適用
    apply_heading_settings(tree, heading_mode, original_heading_settings, original_angles, log)
    
    # 速度設定
    for tag, path in [
        ("globalTransitionalSpeed", ".//wpml:missionConfig"),
        ("autoFlightSpeed", ".//kml:Folder")
    ]:
        el = tree.find(f"{path}/wpml:{tag}", NS)
        if el is not None:
            el.text = str(speed)
        else:
            parent = tree.find(path, NS)
            if parent is not None:
                etree.SubElement(parent, f"{{{NS['wpml']}}}{tag}").text = str(speed)
    
    for pm in tree.findall(".//kml:Placemark", NS):
        for t, v in [("waypointSpeed", speed), ("useGlobalSpeed", 1)]:
            el = pm.find(f"wpml:{t}", NS)
            if el is not None:
                el.text = str(v)
            else:
                etree.SubElement(pm, f"{{{NS['wpml']}}}{t}").text = str(v)
    
    # アクション削除
    for pm in tree.findall(".//kml:Placemark", NS):
        for ag in list(pm.findall("wpml:actionGroup", NS)):
            pm.remove(ag)
    # ジンバルピッチ「フリー」時は既存のジンバルピッチ関連アクションのみ消去
    if gimbal_pitch_mode == "free":
        for pm in tree.findall(".//kml:Placemark", NS):
            for ag in pm.findall("wpml:actionGroup", NS):
                for act in list(ag.findall("wpml:action", NS)):
                    func = act.find("wpml:actionActuatorFunc", NS)
                    if func is not None and func.text == "gimbalRotate":
                        ag.remove(act)
    # ズーム倍率「フリー」時は既存のズーム関連アクションのみ消去
    if zoom_mode == "free":
        for pm in tree.findall(".//kml:Placemark", NS):
            for ag in pm.findall("wpml:actionGroup", NS):
                for act in list(ag.findall("wpml:action", NS)):
                    func = act.find("wpml:actionActuatorFunc", NS)
                    if func is not None and func.text == "zoom":
                        ag.remove(act)
    
    pp = tree.find(".//wpml:payloadParam", NS)
    if pp is not None:
        img = pp.find("wpml:imageFormat", NS)
        if img is not None:
            pp.remove(img)
    
    # センサー選択
    if sensor_modes:
        if pp is None:
            fld = tree.find(".//kml:Folder", NS)
            pp = etree.SubElement(fld, f"{{{NS['wpml']}}}payloadParam")
            etree.SubElement(pp, f"{{{NS['wpml']}}}payloadPositionIndex").text = "0"
        
        fmt = ",".join(m.lower() for m in sensor_modes)
        etree.SubElement(pp, f"{{{NS['wpml']}}}imageFormat").text = fmt
    
    # ヨー固定
    if yaw_mode == "fixed" and yaw_angle is not None:
        for hp in tree.findall(".//wpml:globalWaypointHeadingParam", NS):
            m = hp.find("wpml:waypointHeadingMode", NS)
            a = hp.find("wpml:waypointHeadingAngle", NS)
            if m is not None: m.text = "fixed"
            if a is not None: a.text = str(int(yaw_angle))
        
        for pm in tree.findall(".//kml:Placemark", NS):
            hp = pm.find("wpml:waypointHeadingParam", NS)
            if hp is None:
                hp = etree.SubElement(pm, f"{{{NS['wpml']}}}waypointHeadingParam")
            
            for child in list(hp):
                hp.remove(child)
            
            etree.SubElement(hp, f"{{{NS['wpml']}}}waypointHeadingMode").text = "fixed"
            etree.SubElement(hp, f"{{{NS['wpml']}}}waypointHeadingAngle").text = str(int(yaw_angle))
            etree.SubElement(hp, f"{{{NS['wpml']}}}waypointPoiPoint").text = "0.000000,0.000000,0.000000"
            etree.SubElement(hp, f"{{{NS['wpml']}}}waypointHeadingPathMode").text = "followBadArc"
            etree.SubElement(hp, f"{{{NS['wpml']}}}waypointHeadingPoiIndex").text = "0"
    
    # ウェイポイント取得・アクション追加
    pms = sorted(
        [p for p in tree.findall(".//kml:Placemark", NS) if p.find("wpml:index", NS) is not None],
        key=lambda x: int(x.find("wpml:index", NS).text)
    )

    for i, pm in enumerate(pms):
        idx = int(pm.find("wpml:index", NS).text)
        ag = etree.SubElement(pm, f"{{{NS['wpml']}}}actionGroup")
        etree.SubElement(ag, f"{{{NS['wpml']}}}actionGroupId").text = str(idx)
        etree.SubElement(ag, f"{{{NS['wpml']}}}actionGroupStartIndex").text = str(idx)
        etree.SubElement(ag, f"{{{NS['wpml']}}}actionGroupEndIndex").text = str(idx)
        etree.SubElement(ag, f"{{{NS['wpml']}}}actionGroupMode").text = "sequence"

        trg = etree.SubElement(ag, f"{{{NS['wpml']}}}actionTrigger")
        etree.SubElement(trg, f"{{{NS['wpml']}}}actionTriggerType").text = "reachPoint"

        # 動画開始 (最初のみ)
        if do_video and i == 0:
            sr = etree.SubElement(ag, f"{{{NS['wpml']}}}action")
            etree.SubElement(sr, f"{{{NS['wpml']}}}actionId").text = "0"
            etree.SubElement(sr, f"{{{NS['wpml']}}}actionActuatorFunc").text = "startRecord"
            sp = etree.SubElement(sr, f"{{{NS['wpml']}}}actionActuatorFuncParam")
            etree.SubElement(sp, f"{{{NS['wpml']}}}payloadPositionIndex").text = "0"
        
        # ヨー固定アクション追加（「フリー」以外のみ）
        if yaw_mode != "free":
            yt = None
            if yaw_mode == "original" and idx in original_angles:
                yt = original_angles[idx].get("heading")
            elif yaw_mode == "fixed" and yaw_angle is not None:
                yt = yaw_angle
            
            if yt is not None:
                ya = etree.SubElement(ag, f"{{{NS['wpml']}}}action")
                etree.SubElement(ya, f"{{{NS['wpml']}}}actionId").text = "0"
                etree.SubElement(ya, f"{{{NS['wpml']}}}actionActuatorFunc").text = "rotateYaw"
                yp = etree.SubElement(ya, f"{{{NS['wpml']}}}actionActuatorFuncParam")
                etree.SubElement(yp, f"{{{NS['wpml']}}}aircraftHeading").text = str(int(yt))
                etree.SubElement(yp, f"{{{NS['wpml']}}}aircraftPathMode").text = "counterClockwise"
            
            # ジンバルヨー自動維持
            if yaw_mode == "original" and idx in original_angles:
                gy = original_angles[idx].get("yaw")
                if gy is not None:
                    create_gimbal_yaw_action(ag, gy)
        
        # ジンバルピッチ
        if do_gimbal:
            pt = None
            if gimbal_pitch_mode == "original" and idx in original_angles:
                pt = original_angles[idx].get("pitch")
            elif gimbal_pitch_mode == "fixed" and gimbal_pitch_angle is not None:
                pt = gimbal_pitch_angle
            # gimbal_pitch_mode == "free" の場合は何も追加しない
            if pt is not None:
                gp = etree.SubElement(ag, f"{{{NS['wpml']}}}action")
                etree.SubElement(gp, f"{{{NS['wpml']}}}actionId").text = "0"
                etree.SubElement(gp, f"{{{NS['wpml']}}}actionActuatorFunc").text = "gimbalRotate"
                param = etree.SubElement(gp, f"{{{NS['wpml']}}}actionActuatorFuncParam")
                etree.SubElement(param, f"{{{NS['wpml']}}}gimbalRotateMode").text = "absoluteAngle"
                etree.SubElement(param, f"{{{NS['wpml']}}}gimbalPitchRotateEnable").text = "1"
                etree.SubElement(param, f"{{{NS['wpml']}}}gimbalPitchRotateAngle").text = str(int(pt))
                etree.SubElement(param, f"{{{NS['wpml']}}}gimbalRollRotateEnable").text = "0"
                etree.SubElement(param, f"{{{NS['wpml']}}}gimbalRollRotateAngle").text = "0"
                etree.SubElement(param, f"{{{NS['wpml']}}}gimbalYawRotateEnable").text = "0"
                etree.SubElement(param, f"{{{NS['wpml']}}}gimbalYawRotateAngle").text = "0"
                etree.SubElement(param, f"{{{NS['wpml']}}}gimbalRotateTimeEnable").text = "0"
                etree.SubElement(param, f"{{{NS['wpml']}}}gimbalRotateTime").text = "0"
                etree.SubElement(param, f"{{{NS['wpml']}}}payloadPositionIndex").text = "0"
        
        # Zoom センサー選択時
        if "Zoom" in sensor_modes and zoom_mode != "free":
            if zoom_mode == "fixed" and zoom_ratio is not None:
                # 指定倍率設定
                ft = zoom_ratio_to_focal_length(zoom_ratio)
                za = etree.SubElement(ag, f"{{{NS['wpml']}}}action")
                etree.SubElement(za, f"{{{NS['wpml']}}}actionId").text = "0"
                etree.SubElement(za, f"{{{NS['wpml']}}}actionActuatorFunc").text = "zoom"
                zp = etree.SubElement(za, f"{{{NS['wpml']}}}actionActuatorFuncParam")
                etree.SubElement(zp, f"{{{NS['wpml']}}}focalLength").text = str(ft)
                etree.SubElement(zp, f"{{{NS['wpml']}}}payloadPositionIndex").text = "0"
            elif zoom_mode == "original" and idx in original_angles:
                # 元データの focal_length を維持
                ft = original_angles[idx].get("focal_length")
                if ft is not None:
                    za = etree.SubElement(ag, f"{{{NS['wpml']}}}action")
                    etree.SubElement(za, f"{{{NS['wpml']}}}actionId").text = "0"
                    etree.SubElement(za, f"{{{NS['wpml']}}}actionActuatorFunc").text = "zoom"
                    zp = etree.SubElement(za, f"{{{NS['wpml']}}}actionActuatorFuncParam")
                    etree.SubElement(zp, f"{{{NS['wpml']}}}focalLength").text = str(ft)
                    etree.SubElement(zp, f"{{{NS['wpml']}}}payloadPositionIndex").text = "0"


        # 写真撮影
        if do_photo and not do_video:
            # 先にホバリング
            if hover_time > 0:
                hv = etree.SubElement(ag, f"{{{NS['wpml']}}}action")
                etree.SubElement(hv, f"{{{NS['wpml']}}}actionId").text = "0"
                etree.SubElement(hv, f"{{{NS['wpml']}}}actionActuatorFunc").text = "hover"
                hp = etree.SubElement(hv, f"{{{NS['wpml']}}}actionActuatorFuncParam")
                etree.SubElement(hp, f"{{{NS['wpml']}}}hoverTime").text = str(int(hover_time))
            
            # その後に写真撮影
            ph = etree.SubElement(ag, f"{{{NS['wpml']}}}action")
            etree.SubElement(ph, f"{{{NS['wpml']}}}actionId").text = "0"
            etree.SubElement(ph, f"{{{NS['wpml']}}}actionActuatorFunc").text = "takePhoto"
            pp_act = etree.SubElement(ph, f"{{{NS['wpml']}}}actionActuatorFuncParam")
            etree.SubElement(pp_act, f"{{{NS['wpml']}}}fileSuffix").text = f"ウェイポイント{idx}"
            etree.SubElement(pp_act, f"{{{NS['wpml']}}}payloadPositionIndex").text = "0"
            etree.SubElement(pp_act, f"{{{NS['wpml']}}}useGlobalPayloadLensIndex").text = "1"
        
        # 動画モードホバリング（制御後）
        if do_video and hover_time > 0:
            hv2 = etree.SubElement(ag, f"{{{NS['wpml']}}}action")
            etree.SubElement(hv2, f"{{{NS['wpml']}}}actionId").text = "0"
            etree.SubElement(hv2, f"{{{NS['wpml']}}}actionActuatorFunc").text = "hover"
            hp2 = etree.SubElement(hv2, f"{{{NS['wpml']}}}actionActuatorFuncParam")
            etree.SubElement(hp2, f"{{{NS['wpml']}}}hoverTime").text = str(int(hover_time))
        
        # 動画停止 (最後のみ)
        if do_video and i == len(pms) - 1:
            st = etree.SubElement(ag, f"{{{NS['wpml']}}}action")
            etree.SubElement(st, f"{{{NS['wpml']}}}actionId").text = "0"
            etree.SubElement(st, f"{{{NS['wpml']}}}actionActuatorFunc").text = "stopRecord"
            sparam = etree.SubElement(st, f"{{{NS['wpml']}}}actionActuatorFuncParam")
            etree.SubElement(sparam, f"{{{NS['wpml']}}}payloadPositionIndex").text = "0"

# --- KMZ 一括処理 -----------------------------------------------------------
def process_kmz(path,
                do_photo, do_video,
                do_gimbal, gimbal_pitch_angle, gimbal_pitch_mode,
                yaw_angle, yaw_mode,
                speed, sensor_modes, hover_time,
                zoom_ratio, zoom_mode,
                heading_mode, wp_stop_mode, log):

    try:
        log.insert(tk.END, f"=== 処理開始: {os.path.basename(path)} ===\n")

        wd = extract_kmz(path)
        
        # 解凍後すぐに res フォルダを削除
        res_paths = glob.glob(os.path.join(wd, "**", "res"), recursive=True)
        for res_path in res_paths:
            if os.path.isdir(res_path):
                shutil.rmtree(res_path)
                log.insert(tk.END, f"削除: {res_path}\n")
        
        kmls = glob.glob(os.path.join(wd, "**", "template.kml"), recursive=True)
        if not kmls:
            raise FileNotFoundError("template.kml が見つかりませんでした。")

        sensor_list = ', '.join(sensor_modes) if sensor_modes else 'デフォルト'
        log.insert(tk.END, f"使用センサー: {sensor_list}\n")

        heading_mode_name = next((k for k, v in HEADING_MODE_OPTIONS.items() if v == heading_mode), "不明")
        log.insert(tk.END, f"機体ヘディング制御: {heading_mode_name}\n")

        if "Zoom" in sensor_modes:
            if zoom_mode == "original":
                log.insert(tk.END, "ズーム設定: 元を維持\n")
            elif zoom_mode == "fixed" and zoom_ratio is not None:
                log.insert(tk.END, f"ズーム設定: {zoom_ratio:.1f}倍 ({zoom_ratio*24.0:.1f}mm)\n")


        out_root, outdir = prepare_output_dirs(path, do_photo, do_video, sensor_modes)

        for kml in kmls:
            log.insert(tk.END, f"-- テンプレート読み込み: {os.path.basename(kml)}\n")

            parser = etree.XMLParser(remove_blank_text=True)
            tree = etree.parse(kml, parser)

            global_height_elem = tree.find(".//wpml:globalHeight", NS)
            global_height = global_height_elem.text if global_height_elem is not None else "未設定"

            global_height_mode_elem = tree.find(
                ".//wpml:waylineCoordinateSysParam/wpml:heightMode", NS)
            global_height_mode = global_height_mode_elem.text if global_height_mode_elem is not None else "未設定"

            log.insert(tk.END,
                       f"グローバル高度: {global_height}, 標高モード: {global_height_mode}\n")

            placemarks = tree.findall(".//kml:Placemark", NS)
            wpms = [pm for pm in placemarks if pm.find("wpml:index", NS) is not None]
            log.insert(tk.END, f"総ウェイポイント数: {len(wpms)}\n")

            original_angles = extract_original_gimbal_angles(tree)
            if original_angles:
                log.insert(tk.END,
                           f"ジンバル・ズーム情報を持つウェイポイント: {len(original_angles)}個\n")
            else:
                log.insert(tk.END, "元データにジンバル・ズーム情報なし\n")

            original_heading_settings = extract_original_heading_settings(tree)
            log.insert(tk.END, "元のヘディング設定取得完了\n")

            # 各ウェイポイント詳細表示
            for pm in wpms:
                idx = int(pm.find("wpml:index", NS).text)
                mode = pm.find("wpml:heightMode", NS)
                height_mode = mode.text if mode is not None else global_height_mode
                h_elem = pm.find("wpml:ellipsoidHeight", NS)
                if h_elem is None:
                    h_elem = pm.find("wpml:height", NS)
                height_val = h_elem.text if h_elem is not None else global_height

                g = original_angles.get(idx, {})
                pitch = g.get("pitch", "N/A")
                yaw   = g.get("yaw",   "N/A")
                head  = g.get("heading","N/A")
                fl    = g.get("focal_length", "N/A")
                zr    = g.get("zoom_ratio",    None)

                # zoom_ratio が None の場合は "元設定維持"、数値ならフォーマット
                if zr is None:
                    zoom_info = "元設定維持"
                else:
                    zoom_info = f"{zr:.1f}倍({fl:.1f}mm)"

                log.insert(tk.END,
                           f"[WP {idx}] 標高モード={height_mode}, 高度={height_val}, "
                           f"ジンバルピッチ={pitch}°, ジンバルヨー={yaw}°, "
                           f"機体ヘディング={head}°, ズーム={zoom_info}\n")

            log.insert(tk.END, "\n高度補正なし処理開始\n")

            convert_kml(tree,
                        do_photo, do_video,
                        do_gimbal, gimbal_pitch_angle, gimbal_pitch_mode,
                        yaw_angle, yaw_mode, speed,
                        sensor_modes, hover_time,
                        zoom_ratio, zoom_mode,
                        original_angles, heading_mode, original_heading_settings, log, wp_stop_mode)

            log.insert(tk.END, "高度補正なし処理完了\n\n")

            out_path = os.path.join(outdir, os.path.basename(kml))
            tree.write(out_path, encoding="utf-8", pretty_print=True, xml_declaration=True)
            log.insert(tk.END, f"書き出し完了: {out_path}\n")

        # リソースコピー部分を削除（resフォルダをコピーしない）
        # 元のコードの以下の部分をコメントアウト
        # for name in ["res"]:
        #     srcs = glob.glob(os.path.join(wd, "**", name), recursive=True)
        #     if srcs:
        #         src = srcs[0]; dst = os.path.join(outdir, os.path.basename(src))
        #         if os.path.isdir(src):
        #             shutil.copytree(src, dst, dirs_exist_ok=True)
        #         else:
        #             shutil.copy2(src, dst)

        out_kmz = repackage_to_kmz(out_root, path, do_photo, do_video, sensor_modes)
        log.insert(tk.END, f"最終KMZ: {out_kmz}\n")
        log.insert(tk.END, "=== 処理完了 ===\n\n")

        messagebox.showinfo("完了", f"変換完了:\n{out_kmz}")

    except Exception as e:
        messagebox.showerror("エラー", str(e))
        log.insert(tk.END, f"エラー: {e}\n\n")
    finally:
        if os.path.exists("_kmz_work"):
            shutil.rmtree("_kmz_work")



# --- GUI --------------------------------------------------------------------
class AppGUI(ttk.Frame):
    def __init__(self, master):
        super().__init__(master)
        # --- 高度選択 ---
        ttk.Label(self, text="基準高度(ATLの時のみ動作):").grid(row=0, column=0, sticky="w")
        self.hc = ttk.Combobox(self, values=list(HEIGHT_OPTIONS), state="readonly", width=20)
        self.hc.set(next(iter(HEIGHT_OPTIONS)))
        self.hc.grid(row=0, column=1, padx=5, columnspan=2, sticky="w")
        self.hc.bind("<<ComboboxSelected>>", self.on_height_change)
        self.he = ttk.Entry(self, width=10, state="disabled")
        self.he.grid(row=0, column=3, padx=5)

        # --- 速度 ---
        ttk.Label(self, text="速度 (1–15 m/s):").grid(row=1, column=0, sticky="w", pady=5)
        self.sp = tk.IntVar(value=15)
        ttk.Spinbox(self, from_=1, to=15, textvariable=self.sp, width=5).grid(row=1, column=1, columnspan=2, sticky="w")

        # --- 撮影モード選択（排他） ---
        self.capture_mode_var = tk.StringVar(value="none")  # デフォルトは「撮影なし」
        modes = [("撮影なし", "none"), ("写真撮影", "photo"), ("動画撮影", "video")]
        for i, (text, val) in enumerate(modes):
            ttk.Radiobutton(
                self,
                text=text,
                variable=self.capture_mode_var,
                value=val,
                command=self.update_capture_mode
            ).grid(row=2, column=i, sticky="w", padx=5)

        # --- センサー選択 ---
        self.camera_frame = ttk.LabelFrame(self, text="カメラ選択（必須）")
        # チェックボタン配置
        self.sm_vars = {m: tk.BooleanVar(value=False) for m in SENSOR_MODES}
        for i, m in enumerate(SENSOR_MODES):
            ttk.Checkbutton(
                self.camera_frame,
                text=m,
                variable=self.sm_vars[m],
                command=self.update_zoom
            ).grid(row=0, column=1 + i, sticky="w")
        # フレームを撮影モード選択の下に固定配置
        self.camera_frame.grid(row=3, column=0, columnspan=4, sticky="we", pady=5)

        # ズーム倍率選択
        ttk.Label(self, text="ズーム倍率:").grid(row=4, column=0, sticky="w", pady=5)
        zoom_values = ["元の設定を維持"] + [k for k in ZOOM_RATIO_OPTIONS if ZOOM_RATIO_OPTIONS[k] != "original"]
        self.zm_mode = ttk.Combobox(self, values=zoom_values, state="readonly", width=15)
        self.zm_mode.set("元の設定を維持")
        self.zm_mode.grid(row=4, column=1, padx=5, columnspan=2, sticky="w")
        self.zm_mode.bind("<<ComboboxSelected>>", self.update_zoom)
        # 手動入力用エントリ
        self.zm_entry = ttk.Entry(self, width=8, state="disabled")

        # --- ジンバルピッチ角 (撮影モード時のみ表示) ---
        ttk.Label(self, text="ジンバルピッチ角:").grid(row=5, column=0, sticky="w", pady=5)
        self.gp_mode = ttk.Combobox(
            self, values=list(GIMBAL_PITCH_OPTIONS),
            state="disabled", width=15
        )
        self.gp_mode.set("元の角度維持")
        self.gp_mode.bind("<<ComboboxSelected>>", self.update_gimbal_pitch)
        self.gp_mode.grid(row=5, column=1, padx=5, columnspan=2, sticky="w", pady=5)
        # 入力欄の行をズーム倍率・ヨーと揃える（column=3, sticky="w"）
        self.gp_entry = ttk.Entry(self, width=8, state="disabled")
        # --- ヨー固定（チェックボックス削除、プルダウンのみ） ---
        ttk.Label(self, text="撮影時ヨー角:").grid(row=6, column=0, sticky="w")
        self.yc = ttk.Combobox(self, values=list(YAW_OPTIONS), state="readonly", width=15)
        self.yc.bind("<<ComboboxSelected>>", self.update_yaw)
        self.yc.set("元の角度維持")
        self.yc.grid(row=6, column=1, padx=5, columnspan=2, sticky="w")
        self.ye = ttk.Entry(self, width=8, state="disabled")

        # --- 飛行時ヨー角制御 ---
        ttk.Label(self, text="飛行時ヨー角:").grid(row=7, column=0, sticky="w", pady=5)
        # デフォルトを「元の設定を維持」に設定
        self.heading_mode_var = tk.StringVar(value="元の設定を維持")
        heading_frame = ttk.Frame(self)
        heading_frame.grid(row=7, column=1, columnspan=3, sticky="w", padx=5)
        self.heading_mode_radios = []
        for i, (text, value) in enumerate(HEADING_MODE_OPTIONS.items()):
            rb = ttk.Radiobutton(
                heading_frame,
                text=text,
                variable=self.heading_mode_var,
                value=text
            )
            rb.grid(row=0, column=i, sticky="w", padx=5)
            self.heading_mode_radios.append((text, rb))


        # --- ウェイポイント停止モード選択 ---
        ttk.Label(self, text="WP到達時動作:").grid(row=8, column=0, sticky="w", pady=5)
        self.stop_mode_var = tk.StringVar(value="stop")  # 初期値：停止
        rb_stop = ttk.Radiobutton(self, text="停止する", variable=self.stop_mode_var, value="stop")
        rb_cont = ttk.Radiobutton(self, text="停止しない", variable=self.stop_mode_var, value="continuous")
        rb_stop.grid(row=8, column=1, sticky="w")
        rb_cont.grid(row=8, column=2, sticky="w")

        # --- ホバリング ---
        self.hv = tk.BooleanVar(value=False)
        self.hover_check = ttk.Checkbutton(self, text="ホバリング", variable=self.hv, command=self.update_hover)
        self.hover_check.grid(row=9, column=0, sticky="w", pady=5)
        self.hover_time_label = ttk.Label(self, text="ホバリング時間 (秒):")
        self.hover_time_var = tk.StringVar(value="2")
        self.hover_time_entry = ttk.Entry(self, textvariable=self.hover_time_var, width=8)

        # --- UI 初期化 ---
        self.update_capture_mode()
        self.update_zoom()
        self.update_gimbal_pitch()
        self.update_yaw()
        self.update_hover()
        self.update_hover_mask()  # ← 追加

        # WP停止モードの変更時にホバリング有効/無効を切り替え
        self.stop_mode_var.trace_add("write", lambda *args: self.update_hover_mask())
    
    def on_height_change(self, event=None):
        if HEIGHT_OPTIONS[self.hc.get()] == "custom":
            self.he.config(state="normal")
            self.he.delete(0, tk.END)
            self.he.focus()
        else:
            self.he.config(state="disabled")
            self.he.delete(0, tk.END)
    
    def update_capture_mode(self):
        mode = self.capture_mode_var.get()

        # カメラ選択チェックボタンの有効／無効を切り替え
        for child in self.camera_frame.winfo_children():
            if isinstance(child, ttk.Checkbutton):
                if mode in ("photo", "video"):
                    child.state(["!disabled"])
                else:
                    child.state(["disabled"])

        # ラベル付きフレームのタイトルを変更
        title = "カメラ選択"
        if mode in ("photo", "video"):
            title += "（必須）"
        else:
            title += "（使用不可）"
        self.camera_frame.configure(text=title)

        # ジンバルピッチプルダウンの有効／無効切り替え
        if mode in ("photo", "video"):
            self.gp_mode.config(state="readonly")
        else:
            self.gp_mode.config(state="disabled")
            # StringVar への設定を削除し、Combobox に直接設定
            self.gp_mode.set("元の角度維持")
            self.gp_entry.grid_forget()
            self.gp_entry.config(state="disabled")

        # ヨー角プルダウンの有効/無効切り替え
        # 写真または動画のときは撮影時ヨー角を選択可能にする
        if mode in ("photo", "video"):
             self.yc.config(state="readonly")
             self.update_yaw()
        else:
             self.yc.config(state="disabled")
             self.ye.config(state="disabled")
             self.ye.grid_forget()

        # 飛行時ヨー角「次の撮影方向を向く」ラジオボタンの有効/無効切り替え
        for text, rb in self.heading_mode_radios:
            if text == "次の撮影方向を向く":
                if mode == "none":
                    rb.state(["disabled"])
                    # 「撮影なし」時に選択されていたら他に切り替え
                    if self.heading_mode_var.get() == "次の撮影方向を向く":
                        self.heading_mode_var.set("飛行方向を向く")
                else:
                    rb.state(["!disabled"])
            else:
                rb.state(["!disabled"])

        # 撮影なし選択時にはZoomチェックと設定をリセット
        if mode == "none":
            # Zoomセンサー選択をOFFに
            self.sm_vars["Zoom"].set(False)
            # Comboboxをデフォルトに戻し、グレーアウト
            self.zm_mode.set("元の設定を維持")
        # 撮影モード変更後にもズームUI表示を再評価
        self.update_zoom()

    
    def update_zoom(self, event=None):
        """
        ズーム倍率UIの表示制御:
        - 撮影モードが photo/video かつ Zoom センサー選択時のみ有効(グレーアウト解除)
        - それ以外はグレーアウト(非活性)にする
        """
        mode = self.capture_mode_var.get()
        zoom_selected = self.sm_vars["Zoom"].get()

        # Comboboxは常に表示しておき、stateだけ切り替える
        self.zm_mode.grid(row=4, column=1, padx=5, columnspan=2, sticky="w")

        if mode in ("photo", "video") and zoom_selected:
            # 有効化
            self.zm_mode.config(state="readonly")
            if self.zm_mode.get() == "手動入力":
                self.zm_entry.config(state="normal")
                self.zm_entry.grid(row=4, column=3, sticky="w")
            else:
                self.zm_entry.config(state="disabled")
                self.zm_entry.grid_forget()
        else:
            # グレーアウト
            self.zm_mode.config(state="disabled")
            self.zm_entry.config(state="disabled")
            self.zm_entry.grid_forget()

    def update_gimbal_pitch(self, event=None):
        # Combobox から直接値を取得
        choice = self.gp_mode.get()
        state = str(self.gp_mode.cget("state"))
        if choice == "手動入力" and state in ("readonly", "normal"):
            self.gp_entry.grid(row=5, column=3, sticky="w")
            self.gp_entry.config(state="normal")
        else:
            self.gp_entry.grid_forget()
            self.gp_entry.config(state="disabled")
        
    def update_yaw(self, event=None):
        if self.capture_mode_var.get() not in ("photo", "video"):
            self.ye.config(state="disabled")
            self.ye.grid_forget()
            return
        choice = self.yc.get()
        if choice == "手動入力":
            self.ye.config(state="normal")
            self.ye.grid(row=6, column=3, sticky="w")
        else:
            self.ye.config(state="disabled")
            self.ye.grid_forget()
    
    def update_hover(self):
        if self.hv.get():
            self.hover_time_label.grid(row=9, column=1, sticky="e", padx=(10, 2))
            self.hover_time_entry.grid(row=9, column=2, sticky="w")
        else:
            self.hover_time_label.grid_forget()
            self.hover_time_entry.grid_forget()
    
    def update_hover_mask(self):
        """WP到達時動作が「停止する」以外のときはホバリング選択をグレーアウト＆解除"""
        stop_mode = self.stop_mode_var.get()
        if stop_mode == "stop":
            self.hover_check.state(["!disabled"])
        else:
            self.hover_check.state(["disabled"])
            self.hv.set(False)
            self.hover_time_label.grid_forget()
            self.hover_time_entry.grid_forget()
    
    def get_params(self):
        # offset削除 -> 高度オフセットは常に0
        offset = 0.0
        v = HEIGHT_OPTIONS[self.hc.get()]
        if v == "custom":
            try:
                offset = float(self.he.get())
            except:
                offset = 0.0
        else:
            offset = float(v)

        # ヨー設定
        yaw_angle = None
        yaw_mode = "none"
        mode = self.capture_mode_var.get()
        # 写真または動画選択時はプルダウンの値を反映（動画時も選択可能にしたため）
        if mode in ("photo", "video"):
             yval = YAW_OPTIONS[self.yc.get()]
             if yval == "original":
                 yaw_mode = "original"
             elif yval == "custom":
                 yaw_mode = "fixed"
                 try:
                     yaw_angle = float(self.ye.get())
                 except:
                     yaw_angle = None
             elif yval == "free":
                 yaw_mode = "free"
             else:
                 yaw_mode = "fixed"
                 yaw_angle = float(yval)
        else:
            # 写真撮影なし時は強制的に"free"
            yaw_mode = "free"

        # --- ジンバルピッチ設定 ---
        capture_mode = self.capture_mode_var.get()
        if capture_mode in ("photo", "video"):
            val = GIMBAL_PITCH_OPTIONS[self.gp_mode.get()]  # 直接取得に変更
            if val == "original":
                gimbal_pitch_mode = "original"
                gimbal_pitch_angle = None
            elif val == "custom":
                gimbal_pitch_mode = "fixed"
                try:
                    gimbal_pitch_angle = float(self.gp_entry.get())
                except:
                    gimbal_pitch_angle = None
            elif val == "free":
                gimbal_pitch_mode = "free"
                gimbal_pitch_angle = None
            else:
                gimbal_pitch_mode = "fixed"
                gimbal_pitch_angle = float(val)
        else:
            # 撮影なしではジンバル操作なし
            gimbal_pitch_mode = "none"
            gimbal_pitch_angle = None

        # do_gimbal フラグを gimbal_pitch_mode に合わせて定義
        do_gimbal = (gimbal_pitch_mode != "none")

        # ズーム設定（プルダウンのみ）
        zoom_choice = self.zm_mode.get()
        if zoom_choice == "元の設定を維持":
            zoom_mode = "original"
            zoom_ratio = None
        elif zoom_choice == "手動入力":
            zoom_mode = "fixed"
            try:
                zoom_ratio = float(self.zm_entry.get())
            except:
                zoom_ratio = None
        elif zoom_choice == "フリー":
            zoom_mode = "free"
            zoom_ratio = None
        else:
            zoom_mode = "fixed"
            zoom_ratio = ZOOM_RATIO_OPTIONS[zoom_choice]

        # ホバリング時間
        hover_time = 0
        # WP到達時動作が「停止する」以外のときは強制的に0
        if self.stop_mode_var.get() == "stop" and self.hv.get():
            try:
                hover_time = max(0, float(self.hover_time_var.get()))
            except:
                hover_time = 2.0
        else:
            hover_time = 0

        # 機体ヘディング制御モード取得
        heading_mode = HEADING_MODE_OPTIONS[self.heading_mode_var.get()]

        mode = capture_mode

        return {
            "do_photo": (mode == "photo"),
            "do_video": (mode == "video"),
            "do_gimbal": do_gimbal,
            "gimbal_pitch_angle": gimbal_pitch_angle,
            "gimbal_pitch_mode": gimbal_pitch_mode,
            "yaw_angle": yaw_angle,
            "yaw_mode": yaw_mode,
            "speed": max(1, min(15, self.sp.get())),
            "sensor_modes": [m for m, var in self.sm_vars.items() if var.get()],
            "hover_time": hover_time,
            "zoom_ratio": zoom_ratio,
            "zoom_mode": zoom_mode,
            "heading_mode": heading_mode,
            "wp_stop_mode": self.stop_mode_var.get()
        }



# --- エントリポイント -------------------------------------------------------

def main():
    root = TkinterDnD.Tk()
    root.title("ASL 変換＋撮影制御ツール")
    root.geometry("800x900")
    
    frm = ttk.Frame(root, padding=10)
    frm.pack(fill="both", expand=True)
    
    app = AppGUI(frm)
    app.pack(fill="x", pady=(0, 10))
    
    drop = tk.Label(frm, text=".kmz をここにドロップ", bg="lightgray",
                   width=70, height=5, relief=tk.RIDGE)
    drop.pack(pady=12, fill="x")
    drop.drop_target_register(DND_FILES)
    
    log_frame = ttk.LabelFrame(frm, text="ログ")
    log_frame.pack(fill="both", expand=True)
    
    log = scrolledtext.ScrolledText(log_frame, height=16)
    log.pack(fill="both", expand=True)
    
    def on_drop(event):
        path = event.data.strip("{}")
        if not path.lower().endswith(".kmz"):
            messagebox.showwarning("警告", ".kmz ファイルのみ対応しています。")
            return
        
        params = app.get_params()
        params.update({"path": path, "log": log})
        
        threading.Thread(target=process_kmz, kwargs=params, daemon=True).start()
    
    drop.dnd_bind("<<Drop>>", on_drop)
    root.mainloop()

if __name__ == "__main__":
    main()