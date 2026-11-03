import os
import sys
import subprocess
import uuid
import time
import shutil
import io
import stat
from datetime import datetime
from PIL import Image as PILImage
from openpyxl import Workbook
from openpyxl.drawing.image import Image as OpenpyxlImage

# ================= 魔法代码：一键运行 =================
if __name__ == "__main__" and os.environ.get("STREAMLIT_RUNNING") != "1":
    os.environ["STREAMLIT_RUNNING"] = "1"
    script_path = os.path.abspath(__file__)
    subprocess.run(["streamlit", "run", script_path])
    sys.exit()
# ==========================================================

import streamlit as st
import pandas as pd

st.set_page_config(page_title="物料管理 / Material Tracker", layout="wide")
st.title("📦 物料档案(上海) / Material Management(Only Shanghai office)")

# 自动获取当前电脑的用户主目录
user_home = os.path.expanduser("~")
SHAREPOINT_BASE = os.path.join(user_home, r"OneDrive - spectricity.com\Spectricity Home - MaterialTracker")

CSV_FILE = os.path.join(SHAREPOINT_BASE, "materials.csv")
PHOTO_DIR = os.path.join(SHAREPOINT_BASE, "photos")

if not os.path.exists(PHOTO_DIR):
    os.makedirs(PHOTO_DIR)

cols = ["时间", "SN码", "物料类型", "供应商", "物料名称", "动作", "去向/来源", "经办人", "备注", "照片路径列表", "档案备注"]

# ================= 安全的读写函数 =================
def safe_to_csv(df, filepath, mode='w'):
    for i in range(3):
        try:
            if mode == 'a':
                df.to_csv(filepath, mode='a', header=False, index=False, encoding='utf-8-sig')
            else:
                df.to_csv(filepath, index=False, encoding='utf-8-sig')
            return True
        except PermissionError:
            if i == 2:
                st.error("⚠️ 写入失败！请检查是否正在用 Excel 打开 materials.csv，请关闭它后重试。 / Write failed! Please close the open CSV file.")
                return False
            time.sleep(1)
        except Exception as e:
            st.error(f"保存失败 / Save failed：{e}")
            return False
    return False

# ================= 数据读取与防错 =================
if os.path.exists(CSV_FILE):
    try:
        df_temp = pd.read_csv(CSV_FILE, dtype=str).fillna("")
        need_rewrite = False
        for c in cols:
            if c not in df_temp.columns: 
                df_temp[c] = ""
                need_rewrite = True
        if need_rewrite:
            safe_to_csv(df_temp, CSV_FILE, mode='w')
    except Exception:
        pd.DataFrame(columns=cols).to_csv(CSV_FILE, index=False, encoding='utf-8-sig')
else:
    pd.DataFrame(columns=cols).to_csv(CSV_FILE, index=False, encoding='utf-8-sig')

df = pd.read_csv(CSV_FILE, dtype=str).fillna("")
for c in cols:
    if c not in df.columns: df[c] = ""
    df[c] = df[c].astype(str).str.strip()
df['动作'] = df['动作'].apply(lambda x: '入库' if '入' in x else ('出库' if '出' in x else ''))


# ================= 核心：生成带图 Excel =================
def export_excel_with_images(df):
    wb = Workbook()
    ws = wb.active
    ws.title = "物料清单-Material List"
    
    unique_sns = df['SN码'].drop_duplicates()
    max_photos_count = 0
    export_data = []
    
    for sn in unique_sns:
        mat_history = df[df['SN码'] == sn].copy()
        mat_history['时间_dt'] = pd.to_datetime(mat_history['时间'], errors='coerce')
        mat_history = mat_history.sort_values('时间_dt', ascending=False)
        
        latest = mat_history.iloc[0]
        last_in = mat_history[mat_history['动作'] == '入库'].head(1)
        last_out = mat_history[mat_history['动作'] == '出库'].head(1)
        
        in_str = f"{last_in.iloc[0]['时间']} | {last_in.iloc[0]['去向/来源']}" if not last_in.empty else "无 / None"
        out_str = f"{last_out.iloc[0]['时间']} | {last_out.iloc[0]['去向/来源']}" if not last_out.empty else "无 / None"
        
        all_imgs = []
        for p_list in mat_history['照片路径列表']:
            if p_list and p_list != "nan":
                for p in p_list.split(','):
                    p = p.strip()
                    if os.path.exists(p):
                        all_imgs.append(p)
        all_imgs = list(dict.fromkeys(all_imgs))
        
        if len(all_imgs) > max_photos_count:
            max_photos_count = len(all_imgs)
            
        export_data.append({
            "base_info": [
                str(sn), str(latest['物料名称']), str(latest['物料类型']), 
                str(latest['供应商']), "已出库 / Out" if latest['动作'] == '出库' else "在库 / In",
                str(in_str), str(out_str), str(latest['档案备注'])
            ],
            "imgs": all_imgs
        })
        
    headers = ["SN码 / SN", "物料名称 / Name", "类型 / Type", "供应商 / Vendor", "状态 / Status", "最近入库 / Last In", "最近出库 / Last Out", "备注 / Remarks"]
    for i in range(max_photos_count):
        headers.append(f"照片 {i+1} / Photo {i+1}")
    ws.append(headers)
    
    for cell in ws[1]: cell.font = cell.font.copy(bold=True)
    
    widths = [15, 20, 10, 12, 10, 25, 25, 30]
    for i, w in enumerate(widths):
        ws.column_dimensions[chr(65+i)].width = w
    for i in range(max_photos_count):
        ws.column_dimensions[chr(73+i)].width = 35

    for row_idx, item in enumerate(export_data):
        row_data = item["base_info"] + [""] * max_photos_count
        ws.append(row_data)
        current_row = ws.max_row
        
        for col_idx in range(1, len(headers) + 1):
            ws.cell(row=current_row, column=col_idx).number_format = '@'
            
        ws.row_dimensions[current_row].height = 150 

        for i, img_path in enumerate(item["imgs"]):
            try:
                pil_img = PILImage.open(img_path)
                pil_img.thumbnail((600, 600), PILImage.Resampling.LANCZOS) 
                img_byte_arr = io.BytesIO()
                pil_img.save(img_byte_arr, format='PNG')
                img_byte_arr.seek(0)
                
                xl_img = OpenpyxlImage(img_byte_arr)
                target_height = 140
                target_width = int(target_height * (pil_img.width / pil_img.height))
                xl_img.width = target_width
                xl_img.height = target_height
                
                col_letter = chr(73 + i)
                ws.add_image(xl_img, f"{col_letter}{current_row}")
            except Exception:
                pass
                
    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer

# ================= 侧边栏 UI =================
with st.sidebar:
    st.header("📝 登记新物料 / Register New Material")
    
    # ================= 自动检测 Git 更新模块 (全双语) =================
    if "has_update" not in st.session_state:
        st.session_state.has_update = False

    def check_git_update():
        try:
            # 静默拉取，超时3秒防卡死
            subprocess.run(["git", "fetch", "origin", "master"], capture_output=True, timeout=3, check=True)
            local_hash = subprocess.check_output(["git", "rev-parse", "HEAD"]).decode("utf-8").strip()
            remote_hash = subprocess.check_output(["git", "rev-parse", "origin/master"]).decode("utf-8").strip()
            
            if local_hash != remote_hash:
                st.session_state.has_update = True
            else:
                st.session_state.has_update = False
        except Exception:
            # 如果未配置 Git 环境（如直接下载的压缩包），则静默忽略
            st.session_state.has_update = False

    check_git_update()

    if st.session_state.has_update:
        if st.button("🚀 发现新版本，点击更新 / New Updates Available!", type="primary", use_container_width=True):
            with st.spinner("正在拉取最新代码... / Pulling latest code..."):
                try:
                    subprocess.run(["git", "pull", "origin", "master"], check=True)
                    st.success("✅ 更新成功！请按 F5 刷新网页 / Update successful! Please press F5 to refresh.")
                    st.session_state.has_update = False
                    time.sleep(2)
                    st.rerun()
                except Exception as e:
                    st.error(f"更新失败 / Update failed: {e}")
    else:
        st.button("✅ 已是最新版本 / Up to Date", disabled=True, use_container_width=True)
        
    st.divider()
    # ================= 更新模块结束 =================

    material = st.text_input("物料名称 / Material Name *")
    
    # 1. 动态合并物料类型
    default_types = ["Sensor", "Lens", "IRC", "PCB", "机构件 / Mechanical"]
    if not df.empty:
        existing_types = [t for t in df["物料类型"].unique() if t and t != "nan"]
        for t in existing_types:
            if t not in default_types:
                default_types.append(t)
    
    type_sel = st.selectbox("物料类型 / Material Type", default_types + ["其他 / Other"])
    if "其他" in type_sel:
        final_type = st.text_input("请输入自定义类型 / Enter Custom Type")
    else:
        final_type = type_sel

    # 2. 动态合并供应商
    default_vendors = ["GC", "Qtech", "Newmax", "Omnivision"]
    if not df.empty:
        existing_vendors = [v for v in df["供应商"].unique() if v and v != "nan"]
        for v in existing_vendors:
            if v not in default_vendors:
                default_vendors.append(v)
                
    vendor_sel = st.selectbox("供应商 / Vendor", default_vendors + ["其他 / Other"])
    if "其他" in vendor_sel:
        final_vendor = st.text_input("请输入自定义供应商 / Enter Custom Vendor")
    else:
        final_vendor = vendor_sel

    # 3. 默认值设置
    handler = st.text_input("入库经办人 / Inbound Operator *", value="Jerry Wang")
    destination = st.text_input("入库地点/来源 / Inbound Location/Source *", value="SH office")
    photos = st.file_uploader("上传照片 (支持多选) / Upload Photos (Multiple)", type=["jpg", "png", "jpeg"], accept_multiple_files=True)
    note = st.text_area("备注/目的 / Remarks/Purpose")
    
    # 4. 独立按钮提交 (防回车误触)
    if st.button("提交登记 / Submit Registration", type="primary", use_container_width=True):
        if not material.strip() or not handler.strip():
            st.error("物料名称和经办人必填！ / Material Name and Operator are required!")
        elif "其他" in type_sel and not final_type.strip():
            st.error("请填写自定义类型 / Please enter custom type")
        elif "其他" in vendor_sel and not final_vendor.strip():
            st.error("请填写自定义供应商 / Please enter custom vendor")
        else:
            sn = f"MAT-{uuid.uuid4().hex[:6].upper()}"
            sn_folder = os.path.join(PHOTO_DIR, sn)
            os.makedirs(sn_folder, exist_ok=True)
            
            saved_photos = []
            if photos:
                for p in photos:
                    p_name = f"{datetime.now().strftime('%H%M%S')}_{p.name}"
                    p_path = os.path.join(sn_folder, p_name)
                    with open(p_path, "wb") as f:
                        f.write(p.getbuffer())
                    saved_photos.append(p_path)
            
            new_row = pd.DataFrame([[ 
                datetime.now().strftime("%Y-%m-%d %H:%M"), sn, final_type, final_vendor, 
                material.strip(), "入库", destination.strip(), handler.strip(), note.strip(), 
                ",".join(saved_photos), ""
            ]], columns=cols)
            
            if safe_to_csv(new_row, CSV_FILE, mode='a'):
                st.success(f"登记成功 / Registration Successful! SN: {sn}")
                time.sleep(0.5)
                st.rerun()
                
    st.divider()
    st.header("📤 导出给同事 / Export for Colleagues")
    if st.button("生成带图 Excel / Generate Excel with Images", use_container_width=True):
        with st.spinner("正在打包全部原图，请稍等... / Packing all original images, please wait..."):
            excel_data = export_excel_with_images(df)
            st.download_button(
                label="👇 点击这里下载 .xlsx 文件 / Click here to download .xlsx",
                data=excel_data,
                file_name=f"物料总览_MaterialList_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )

# ================= 主界面展示 =================
st.markdown("### 🔍 筛选与查找 / Filter & Search")
f_col1, f_col2, f_col3 = st.columns([2, 1, 1])
with f_col1: search = st.text_input("关键词 / Keyword (名称/SN/供应商/经办人/去向/目的/备注)")
with f_col2:
    all_types = ["全部 / All"] + sorted(list(df["物料类型"].unique()))
    type_filter = st.selectbox("按物料类型 / By Type", all_types)
with f_col3:
    all_vendors = ["全部 / All"] + sorted(list(df["供应商"].unique()))
    vendor_filter = st.selectbox("按供应商 / By Vendor", all_vendors)

unique_sns = df[['SN码']].drop_duplicates().sort_values('SN码')
if search:
    unique_sns = unique_sns[unique_sns['SN码'].apply(lambda sn: search.lower() in df[df['SN码']==sn].astype(str).to_string().lower())]
if type_filter != "全部 / All":
    unique_sns = unique_sns[unique_sns['SN码'].isin(df[df["物料类型"] == type_filter]["SN码"].unique())]
if vendor_filter != "全部 / All":
    unique_sns = unique_sns[unique_sns['SN码'].isin(df[df["供应商"] == vendor_filter]["SN码"].unique())]

st.divider()

if unique_sns.empty:
    st.info("没有找到符合条件的物料。 / No matching materials found.")
else:
    for idx, (_, mat_row) in enumerate(unique_sns.iterrows()):
        sn = mat_row['SN码']
        mat_history = df[df['SN码'] == sn].copy().sort_values('时间', ascending=False)
        latest_record = mat_history.iloc[0]
        name, m_type, vendor, archive_note = latest_record['物料名称'], latest_record['物料类型'], latest_record['供应商'], latest_record['档案备注']
        
        last_in = mat_history[mat_history['动作'] == '入库'].head(1)
        last_out = mat_history[mat_history['动作'] == '出库'].head(1)
        
        status = "已出库 / Out 🔴" if latest_record['动作'] == '出库' else "在库 / In 🟢"
        status_color = "red" if latest_record['动作'] == '出库' else "green"

        all_photos = []
        for p_list in mat_history['照片路径列表']:
            if p_list and p_list != "nan":
                for p in p_list.split(','):
                    if os.path.exists(p.strip()): all_photos.append(p.strip())
        all_photos = list(dict.fromkeys(all_photos))

        with st.container(border=True):
            c1, c2 = st.columns([1.5, 3])
            with c1:
                if all_photos:
                    img_cols = st.columns(2)
                    for i, p in enumerate(all_photos[:2]):
                        img_cols[i].image(p, use_container_width=True)
                else:
                    st.write("📷 无照片 / No Photos")
            
            with c2:
                st.markdown(f"#### [{m_type}] {name}")
                st.markdown(f"**供应商 / Vendor:** {vendor} | **SN码 / SN:** `{sn}` | **状态 / Status:** :{status_color}[{status}]")
                st.divider()
                
                detail_col1, detail_col2 = st.columns(2)
                with detail_col1:
                    st.markdown("**📥 最近入库 / Last Inbound**")
                    if not last_in.empty:
                        r = last_in.iloc[0]
                        st.markdown(f"时间 / Time: {r['时间']}  \n地点 / Loc: {r['去向/来源']}  \n经办人 / Op: {r['经办人']}  \n目的 / Purpose: {r['备注']}")
                    else: 
                        st.write("无记录 / No record")
                        
                with detail_col2:
                    st.markdown("**📤 最近出库 / Last Outbound**")
                    if not last_out.empty:
                        r = last_out.iloc[0]
                        st.markdown(f"时间 / Time: {r['时间']}  \n去向 / Dest: {r['去向/来源']}  \n经办人 / Op: {r['经办人']}  \n目的 / Purpose: {r['备注']}")
                    else: 
                        st.write("无出库记录 / No outbound record")
                
                if archive_note:
                    st.write("") 
                    st.markdown(f"**📝 备注 / Remarks：** {archive_note}")

                st.divider()
                
                btn_col1, btn_col2, btn_col3 = st.columns([1, 1, 2])
                with btn_col1:
                    if st.button("📤 借出/出库 / Borrow", key=f"btn_out_{sn}_{idx}", disabled=(status=="已出库 / Out 🔴")):
                        st.session_state[f"action_{sn}_{idx}"] = "出库"
                with btn_col2:
                    if st.button("📥 归还/入库 / Return", key=f"btn_in_{sn}_{idx}", disabled=(status=="在库 / In 🟢")):
                        st.session_state[f"action_{sn}_{idx}"] = "入库"
                
                current_action = st.session_state.get(f"action_{sn}_{idx}")
                if current_action:
                    action_zh = "出库 / Outbound" if current_action == "出库" else "入库 / Inbound"
                    with st.form(key=f"form_action_{sn}_{idx}"):
                        st.info(f"正在执行 / Executing：{action_zh}")
                        act_handler = st.text_input("经办人 / Operator *", key=f"h_{sn}_{idx}", value="Jerry Wang")
                        act_dest = st.text_input("去向/地点 / Destination *", key=f"d_{sn}_{idx}", value="借用人/部门 / Borrower" if current_action=="出库" else "SH office")
                        act_note = st.text_input("本次动作目的/备注 / Purpose/Remarks", key=f"n_{sn}_{idx}")
                        act_photos = st.file_uploader("补充照片 / Supplementary Photos", type=["jpg", "png", "jpeg"], accept_multiple_files=True, key=f"p_{sn}_{idx}")
                        
                        col_sub, col_can = st.columns([1, 1])
                        if col_sub.form_submit_button("确认提交 / Confirm", type="primary"):
                            if not act_handler.strip() or not act_dest.strip():
                                st.error("经办人和去向必填 / Operator and Destination are required")
                            else:
                                new_photos = []
                                if act_photos:
                                    sn_folder = os.path.join(PHOTO_DIR, sn)
                                    for p in act_photos:
                                        p_name = f"{datetime.now().strftime('%H%M%S')}_{p.name}"
                                        p_path = os.path.join(sn_folder, p_name)
                                        with open(p_path, "wb") as f: f.write(p.getbuffer())
                                        new_photos.append(p_path)
                                
                                new_row = pd.DataFrame([[
                                    datetime.now().strftime("%Y-%m-%d %H:%M"), sn, m_type, vendor, name, current_action,
                                    act_dest.strip(), act_handler.strip(), act_note.strip(), ",".join(new_photos), archive_note
                                ]], columns=cols)
                                if safe_to_csv(new_row, CSV_FILE, mode='a'):
                                    del st.session_state[f"action_{sn}_{idx}"]
                                    st.rerun()
                        if col_can.form_submit_button("取消 / Cancel"):
                            del st.session_state[f"action_{sn}_{idx}"]
                            st.rerun()

                with st.expander("✏️ 编辑物料档案 / Edit Material Profile (改名称/供应商/追加照片)"):
                    with st.form(key=f"edit_mat_{sn}"):
                        new_name = st.text_input("物料名称 / Material Name", value=name)
                        new_type = st.text_input("物料类型 / Material Type", value=m_type)
                        new_vendor = st.text_input("供应商 / Vendor", value=vendor)
                        new_archive_note = st.text_area("📝 备注 / Remarks", value=archive_note)
                        extra_photos = st.file_uploader("追加照片 / Add Photos", type=["jpg", "png", "jpeg"], accept_multiple_files=True)
                        
                        if st.form_submit_button("保存修改 / Save Changes"):
                            df.loc[df['SN码'] == sn, '物料名称'] = new_name.strip()
                            df.loc[df['SN码'] == sn, '物料类型'] = new_type.strip()
                            df.loc[df['SN码'] == sn, '供应商'] = new_vendor.strip()
                            df.loc[df['SN码'] == sn, '档案备注'] = new_archive_note.strip()
                            
                            if extra_photos:
                                sn_folder = os.path.join(PHOTO_DIR, sn)
                                existing_paths = str(latest_record['照片路径列表']).split(',') if latest_record['照片路径列表'] else []
                                if existing_paths == [""]: existing_paths = []
                                for p in extra_photos:
                                    p_name = f"{datetime.now().strftime('%H%M%S')}_{p.name}"
                                    p_path = os.path.join(sn_folder, p_name)
                                    with open(p_path, "wb") as f: f.write(p.getbuffer())
                                    existing_paths.append(p_path)
                                latest_idx = mat_history.index[0]
                                df.loc[latest_idx, '照片路径列表'] = ",".join(existing_paths)
                                
                            if safe_to_csv(df, CSV_FILE, mode='w'):
                                st.success("档案已更新 / Profile updated!")
                                st.rerun()
                    
                    st.divider()
                    
                    if st.button("🗑️ 彻底删除此物料 / Permanently Delete Material", key=f"del_all_{sn}", type="primary"):
                        df = df[df['SN码'] != sn]
                        if safe_to_csv(df, CSV_FILE, mode='w'):
                            folder_path = os.path.join(PHOTO_DIR, sn)
                            if os.path.exists(folder_path):
                                try:
                                    def on_rm_error(func, path, exc_info):
                                        os.chmod(path, stat.S_IWRITE)
                                        func(path)
                                    shutil.rmtree(folder_path, onerror=on_rm_error)
                                except Exception as e:
                                    st.warning(f"⚠️ 数据已从系统删除，但本地照片文件夹被占用未能彻底清除（可能是 OneDrive 正在同步）。请稍后手动删除：{folder_path}")
                            
                            st.success("物料记录已彻底删除 / Material record permanently deleted!")
                            st.rerun()

                with st.expander("📜 查看完整历史记录 / View Full History"):
                    hist_c1, hist_c2 = st.columns([2, 1])
                    with hist_c1:
                        for h_idx, r in mat_history.iterrows():
                            color = "green" if r['动作'] == "入库" else "red"
                            act_zh = "入库 / In" if r['动作'] == "入库" else "出库 / Out"
                            st.markdown(f"**:{color}[{act_zh}]** | {r['时间']} | {r['去向/来源']} | {r['经办人']} | 目的 / Purpose: {r['备注']}")
                            
                            with st.expander("修改或删除这条记录 / Edit or Delete"):
                                with st.form(key=f"edit_hist_{sn}_{h_idx}"):
                                    new_h_dest = st.text_input("去向/地点 / Location", value=r['去向/来源'])
                                    new_h_handler = st.text_input("经办人 / Operator", value=r['经办人'])
                                    new_h_note = st.text_input("目的/备注 / Remarks", value=r['备注'])
                                    if st.form_submit_button("保存单条修改 / Save Record"):
                                        df.loc[h_idx, '去向/来源'] = new_h_dest.strip()
                                        df.loc[h_idx, '经办人'] = new_h_handler.strip()
                                        df.loc[h_idx, '备注'] = new_h_note.strip()
                                        if safe_to_csv(df, CSV_FILE, mode='w'):
                                            st.success("单条明细已更新 / Record updated!")
                                            st.rerun()
                                
                                if st.button("🗑️ 删除这条流水 / Delete this record", key=f"del_hist_{sn}_{h_idx}"):
                                    df = df.drop(h_idx)
                                    if safe_to_csv(df, CSV_FILE, mode='w'):
                                        st.success("流水已删除 / Record deleted!")
                                        st.rerun()
                    with hist_c2:
                        if all_photos:
                            st.write("**关联照片 / Associated Photos：**")
                            for p in all_photos:
                                st.image(p, use_container_width=True)