#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
Script Name: Text+ Style Copier (Force Restore v10)
Description: 
    Copies style from Source to Target clips.
    Explicitly re-applies the original text AFTER style transfer to prevent blank text.

[使い方]
1. コピー元（Source）のText+クリップの上に再生ヘッド（赤いバー）を置く。
2. スタイルを適用したい先（Target）のText+クリップを「ピンク (Pink)」に色付けする。
3. スクリプトを実行する。
"""

import sys
import time

# -----------------------------------------------------------------------------
# Logger
# -----------------------------------------------------------------------------
class Logger:
    def __init__(self):
        self.lines = []
    def log(self, message):
        print(message)
        self.lines.append(message)
    def get_full_report(self):
        return "\n".join(self.lines)

# -----------------------------------------------------------------------------
# API Helpers
# -----------------------------------------------------------------------------
def get_resolve():
    try:
        import DaVinciResolveScript as bmd
        return bmd.scriptapp("Resolve")
    except ImportError:
        return sys.modules.get("fusionscript").scriptapp("Resolve")

def find_textplus_tool(fusion_comp, logger):
    if not fusion_comp: return None
    # ID search
    tools_dict = fusion_comp.GetToolList(False, "TextPlus")
    if tools_dict:
        for tool in tools_dict.values():
            return tool
    # Name search fallback
    all_tools = fusion_comp.GetToolList(False)
    if all_tools:
        for tool in all_tools.values():
            name = tool.GetAttrs("TOOLS_Name")
            if "Template" in name or "Text" in name:
                return tool
    return None

# -----------------------------------------------------------------------------
# Adaptive Settings Getter
# -----------------------------------------------------------------------------
def get_tool_settings(tool, logger):
    """
    辞書型(dict)を返すメソッドが見つかるまで試行する。
    """
    candidates = [
        ("GetSettings", getattr(tool, "GetSettings", None)),
        ("SaveSettings", getattr(tool, "SaveSettings", None)),
        ("GetCurrentSettings", getattr(tool, "GetCurrentSettings", None)),
    ]

    for name, method in candidates:
        if method:
            try:
                result = method()
                # 戻り値が有効な辞書かチェック
                if isinstance(result, dict) and len(result) > 0:
                    # 中身が空でないかチェック
                    if 'Inputs' in result or 'Tools' in result or len(result.keys()) > 5:
                        return result, name
            except:
                pass     
    return None, None

# -----------------------------------------------------------------------------
# Force Apply Logic (The Fix)
# -----------------------------------------------------------------------------
def apply_style_and_restore_text(target_tool, source_settings, source_method, logger):
    """
    スタイルを適用した後、強制的にテキストを元の値に戻す。
    """
    import copy
    
    tgt_name = target_tool.GetAttrs("TOOLS_Name")
    
    # 1. テキストの退避 (Backup)
    # StyledTextの内容を確実に保持
    original_text = target_tool.GetInput("StyledText")
    
    # 万が一取得できなかった場合の空文字ガード
    if original_text is None:
        original_text = ""

    # 2. 設定データの準備 (Pre-patch)
    # ※念のため辞書内も書き換えるが、メインは後述のStep 4
    new_settings = copy.deepcopy(source_settings)
    
    # GetSettings型 (Nested)
    if 'Tools' in new_settings:
        for key, val in new_settings['Tools'].items():
            if isinstance(val, dict) and 'Inputs' in val:
                # ソースのキーを削除し、ターゲット名で登録し直す
                if key != tgt_name:
                    new_settings['Tools'][tgt_name] = val
                    del new_settings['Tools'][key]
                break
    
    # 3. スタイルの適用 (Apply)
    apply_success = False
    try:
        if source_method == "SaveSettings" and hasattr(target_tool, "LoadSettings"):
            target_tool.LoadSettings(new_settings)
            apply_success = True
        elif hasattr(target_tool, "SetSettings"):
            target_tool.SetSettings(new_settings)
            apply_success = True
        elif hasattr(target_tool, "LoadSettings"):
            target_tool.LoadSettings(new_settings)
            apply_success = True
            
    except Exception as e:
        logger.log(f"  [Error] Apply failed: {e}")
        return False
    
    if not apply_success:
        return False

    # 4. テキストの強制復元 (Force Restore)
    # スタイル適用によってテキストが書き換わったり消えたりしたものを、
    # 直後に上書きして元に戻す。これが最も確実。
    try:
        target_tool.SetInput("StyledText", original_text)
        # logger.log(f"  [Info] Restored text: '{str(original_text)[:10]}...'")
        return True
    except Exception as e:
        logger.log(f"  [Error] Failed to restore text: {e}")
        return False

# -----------------------------------------------------------------------------
# UI Functions
# -----------------------------------------------------------------------------
def show_report_window(fusion, report_text):
    if not fusion: return
    try:
        if not hasattr(fusion, "UIDispatcher") or fusion.UIDispatcher is None: return
        ui = fusion.UIManager
        dispatcher = fusion.UIDispatcher(ui)
        if not dispatcher: return

        window_layout = [
            ui.VGroup([
                ui.Label({ "Text": "Style Transfer Report", "Weight": 0, "Font": ui.Font({ "PixelSize": 16, "Bold": True }), "Alignment": { "AlignHCenter": True } }),
                ui.VGap(10),
                ui.TextEdit({ "ID": "ReportField", "Text": report_text, "ReadOnly": True, "Font": ui.Font({ "Family": "Monospace" }), "Weight": 1.0, "MinimumSize": [600, 300] }),
                ui.VGap(10),
                ui.HGroup({ "Weight": 0 }, [ ui.HGap(), ui.Button({ "ID": "CloseBtn", "Text": "Close", "MinimumSize": [100, 30] }) ])
            ])
        ]
        dlg = dispatcher.AddWindow({ "ID": "ReportWin", "WindowTitle": "Result v10", "Geometry": [400, 200, 640, 400] }, window_layout)
        def on_close(ev): dispatcher.ExitLoop()
        dlg.On.ReportWin.Close = on_close
        dlg.On.CloseBtn.Clicked = on_close
        dlg.Show()
        dispatcher.RunLoop()
        dlg.Hide()
    except: pass

# -----------------------------------------------------------------------------
# Main Logic
# -----------------------------------------------------------------------------
def main():
    logger = Logger()
    resolve = get_resolve()
    if not resolve: return

    fusion = resolve.Fusion()
    project = resolve.GetProjectManager().GetCurrentProject()
    timeline = project.GetCurrentTimeline()
    
    if not timeline:
        logger.log("[Error] No timeline open.")
        return

    # 1. Source
    source_item = timeline.GetCurrentVideoItem()
    if not source_item:
        logger.log("[Error] No clip found at playhead.")
        return

    source_name = source_item.GetName()
    logger.log(f"Source: {source_name}")

    # Auto-load Comp
    source_comp = source_item.GetFusionCompByIndex(1)
    if not source_comp:
        resolve.OpenPage("Fusion")
        time.sleep(1.0)
        source_comp = source_item.GetFusionCompByIndex(1)
        resolve.OpenPage("Edit")
        if not source_comp:
            logger.log("[Error] Could not load Fusion Comp.")
            return

    source_tool = find_textplus_tool(source_comp, logger)
    if not source_tool:
        logger.log("[Error] Text+ tool not found.")
        return

    # Get Settings
    source_settings, method_used = get_tool_settings(source_tool, logger)
    
    if not source_settings:
        logger.log("[Fatal Error] Could not retrieve settings.")
        show_report_window(fusion, logger.get_full_report())
        return

    logger.log(f"[Success] Settings captured via '{method_used}'")

    # 2. Targets
    target_color = "Pink"
    target_items = []
    
    track_count = timeline.GetTrackCount("video")
    for i in range(1, track_count + 1):
        items = timeline.GetItemListInTrack("video", i)
        if items:
            for item in items:
                if item == source_item: continue
                if item.GetClipColor() == target_color:
                    target_items.append(item)

    if not target_items:
        logger.log(f"[Warning] No '{target_color}' clips found.")
        show_report_window(fusion, logger.get_full_report())
        return

    logger.log(f"Targets: {len(target_items)} clips.")

    # 3. Apply
    success_count = 0
    resolve.OpenPage("Edit")

    for item in target_items:
        tgt_name = item.GetName()
        tgt_comp = item.GetFusionCompByIndex(1)
        
        if not tgt_comp:
             logger.log(f"[Skip] {tgt_name}: No Comp.")
             continue

        tgt_tool = find_textplus_tool(tgt_comp, logger)
        if not tgt_tool:
             logger.log(f"[Skip] {tgt_name}: No Text+ tool.")
             continue

        try:
            # Apply Style AND Restore Text
            if apply_style_and_restore_text(tgt_tool, source_settings, method_used, logger):
                success_count += 1
                logger.log(f"[OK] {tgt_name}")
            else:
                logger.log(f"[Fail] {tgt_name}: Apply failed.")

        except Exception as e:
            logger.log(f"[Fail] {tgt_name}: {e}")

    logger.log("-" * 40)
    logger.log(f"Done. Updated {success_count} clips.")
    
    show_report_window(fusion, logger.get_full_report())

if __name__ == "__main__":
    main()