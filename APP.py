import os
import sys
import subprocess
import uuid
import time
import shutil  # 用于删除照片文件夹

# ================= 魔法代码：一键运行 =================
if __name__ == "__main__" and os.environ.get("STREAMLIT_RUNNING") != "1":
    os.environ["STREAMLIT_RUNNING"] = "1"
    script_path = os.path.abspath(__file__)
    subprocess.run(["streamlit", "run", script_path])
    sys.exit()
# ==========================================================

import streamlit as st
import pandas as pd
from datetime import datetime

st.set_page_config(page_title="物料档案与借还管理", layout="wide")
st.title("📦 物料档案与借还管理")

# 把下面引号里的路径，换成你刚才复制的真实路径！
SHAREPOINT_BASE = r"C:\Users\Administrator\OneDrive - spectricity.com\Spectricity Home - MaterialTracker"

# 系统会自动在这个共享文件夹里创建一个 Excel 和 photos 文件夹
CSV_FILE = os.path.join(SHAREPOINT_BASE, "materials.csv")
PHOTO_DIR = os.path.join(SHAREPOINT_BASE, "photos")

VENDOR_OPTIONS = ["GC", "Qtech", "Newmax", "Omnivision", "其他"]

if not os.path.exists(PHOTO_DIR):
    os.makedirs(PHOTO_DIR)

# 核心数据结构：新增了 "档案备注" 独立列
cols = ["时间", "SN码", "物料类型", "供应商", "物料名称", "动作", "去向/来源", "经办人", "备注", "照片路径列表", "档案备注"]

# ================= 终极防错位逻辑 =================
if os.path.exists(CSV_FILE):
    try:
        df_temp = pd.read_csv(CSV_FILE, dtype=str).fillna("")
        # 自动补充缺失的新列，确保旧数据不丢
        for c in cols:
            if c not in df_temp.columns:
                df_temp[c] = ""
        df_temp.to_csv(CSV_FILE, index=False, encoding='utf-8-sig')
    except Exception:
        pd.DataFrame(columns=cols).to_csv(CSV_FILE, index=False, encoding='utf-8-sig')
else:
    pd.DataFrame(columns=cols).to_csv(CSV_FILE, index=False, encoding='utf-8-sig')

df = pd.read_csv(CSV_FILE, dtype=str).fillna("")
for c in cols:
    if c not in df.columns: df[c] = ""
    df[c] = df[c].astype(str).str.strip()

df['动作'] = df['动作'].apply(lambda x: '入库' if '入' in x else ('出库' if '出' in x else ''))

# ================= 侧边栏：登记新物料 =================
with st.sidebar:
    st.header("📝 登记新物料 (初始入库)")
    with st.form("add_form", clear_on_submit=True):
        material = st.text_input("物料名称 *")
        mat_type = st.selectbox("物料类型", ["Sensor", "Lens", "IRC", "PCB", "机构件", "其他"])
        final_type = mat_type if mat_type != "其他" else st.text_input("自定义类型")
        
        vendor = st.selectbox("供应商", VENDOR_OPTIONS)
        final_vendor = vendor if vendor != "其他" else st.text_input("自定义供应商")
        
        handler = st.text_input("入库经办人 *")
        destination = st.text_input("入库地点/来源 *", value="公司仓库")
        photos = st.file_uploader("上传照片 (支持多选)", type=["jpg", "png", "jpeg"], accept_multiple_files=True)
        note = st.text_area("备注/目的 (例如: 采购新料)")
        
        submitted = st.form_submit_button("提交登记")
        
        if submitted:
            if not material.strip() or not handler.strip():
                st.error("物料名称和经办人必填！")
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
                    datetime.now().strftime("%Y-%m-%d %H:%M"),
                    sn, final_type, final_vendor, material.strip(), "入库",
                    destination.strip(), handler.strip(), note.strip(), ",".join(saved_photos),
                    "" # 初始没有档案备注
                ]], columns=cols)
                
                new_row.to_csv(CSV_FILE, mode='a', header=False, index=False, encoding='utf-8-sig')
                st.success(f"登记成功！专属 SN 码: {sn}")
                st.rerun()

# ================= 主界面 =================
st.markdown("### 🔍 筛选与查找")
f_col1, f_col2, f_col3 = st.columns([2, 1, 1])
with f_col1:
    search = st.text_input("关键词 (名称/SN/供应商/经办人/去向/目的/档案备注)")
with f_col2:
    all_types = ["全部"] + sorted(list(df["物料类型"].unique()))
    type_filter = st.selectbox("按物料类型", all_types)
with f_col3:
    all_vendors = ["全部"] + sorted(list(df["供应商"].unique()))
    vendor_filter = st.selectbox("按供应商", all_vendors)

unique_sns = df[['SN码']].drop_duplicates().sort_values('SN码')

if search:
    unique_sns = unique_sns[unique_sns['SN码'].apply(lambda sn: search.lower() in df[df['SN码']==sn].astype(str).to_string().lower())]
if type_filter != "全部":
    valid_sns = df[df["物料类型"] == type_filter]["SN码"].unique()
    unique_sns = unique_sns[unique_sns['SN码'].isin(valid_sns)]
if vendor_filter != "全部":
    valid_sns = df[df["供应商"] == vendor_filter]["SN码"].unique()
    unique_sns = unique_sns[unique_sns['SN码'].isin(valid_sns)]

st.divider()

if unique_sns.empty:
    st.info("没有找到符合条件的物料。")
else:
    for idx, (_, mat_row) in enumerate(unique_sns.iterrows()):
        sn = mat_row['SN码']
        
        mat_history = df[df['SN码'] == sn].copy()
        mat_history = mat_history.sort_values('时间', ascending=False)
        
        latest_record = mat_history.iloc[0]
        name = latest_record['物料名称']
        m_type = latest_record['物料类型']
        vendor = latest_record['供应商']
        archive_note = latest_record['档案备注'] # 提取整体备注
        
        last_in = mat_history[mat_history['动作'] == '入库'].head(1)
        last_out = mat_history[mat_history['动作'] == '出库'].head(1)
        
        if latest_record['动作'] == '出库':
            status = "已出库 🔴"
            status_color = "red"
        else:
            status = "在库 🟢"
            status_color = "green"

        all_photos = []
        for p_list in mat_history['照片路径列表']:
            if p_list and p_list != "nan":
                for p in p_list.split(','):
                    p = p.strip()
                    if p and os.path.exists(p):
                        all_photos.append(p)
        all_photos = list(dict.fromkeys(all_photos))

        with st.container(border=True):
            c1, c2 = st.columns([1.5, 3])
            
            with c1:
                if all_photos:
                    img_cols = st.columns(2)
                    for i, p in enumerate(all_photos[:2]):
                        img_cols[i].image(p, use_container_width=True)
                    if len(all_photos) > 2:
                        st.caption(f"共 {len(all_photos)} 张照片，展开详情查看全部")
                else:
                    st.write("📷 无照片")
            
            with c2:
                st.markdown(f"#### [{m_type}] {name}")
                st.markdown(f"**供应商:** {vendor} | **SN码:** `{sn}` | **当前状态:** :{status_color}[{status}]")
                st.divider()
                
                detail_col1, detail_col2 = st.columns(2)
                with detail_col1:
                    st.markdown("**📥 最近入库**")
                    if not last_in.empty:
                        r = last_in.iloc[0]
                        st.write(f"时间: {r['时间']}")
                        st.write(f"地点: {r['去向/来源']}")
                        st.write(f"经办人: {r['经办人']}")
                        st.write(f"目的: {r['备注'] if r['备注'] else '无'}")
                    else:
                        st.write("无记录")
                        
                with detail_col2:
                    st.markdown("**📤 最近出库**")
                    if not last_out.empty:
                        r = last_out.iloc[0]
                        st.write(f"时间: {r['时间']}")
                        st.write(f"去向: {r['去向/来源']}")
                        st.write(f"经办人: {r['经办人']}")
                        st.write(f"目的: {r['备注'] if r['备注'] else '无'}")
                    else:
                        st.write("无出库记录")
                
                # ====== 这里就是你要的红圈区域 ======
                if archive_note:
                    st.write("") # 加点空隙
                    st.markdown(f"**📝 备注：** {archive_note}")
                # ====================================

                st.divider()
                
                # 快捷借还按钮
                btn_col1, btn_col2, btn_col3 = st.columns([1, 1, 2])
                with btn_col1:
                    if st.button("📤 借出/出库", key=f"btn_out_{sn}_{idx}", disabled=(status=="已出库 🔴")):
                        st.session_state[f"action_{sn}_{idx}"] = "出库"
                with btn_col2:
                    if st.button("📥 归还/入库", key=f"btn_in_{sn}_{idx}", disabled=(status=="在库 🟢")):
                        st.session_state[f"action_{sn}_{idx}"] = "入库"
                
                current_action = st.session_state.get(f"action_{sn}_{idx}")
                if current_action:
                    with st.form(key=f"form_action_{sn}_{idx}"):
                        st.info(f"正在执行：{current_action}")
                        act_handler = st.text_input("经办人 *", key=f"h_{sn}_{idx}")
                        act_dest = st.text_input("去向/地点 *", key=f"d_{sn}_{idx}", value="借用人/部门" if current_action=="出库" else "公司仓库")
                        act_note = st.text_input("本次动作目的/备注", key=f"n_{sn}_{idx}")
                        act_photos = st.file_uploader("补充照片", type=["jpg", "png", "jpeg"], accept_multiple_files=True, key=f"p_{sn}_{idx}")
                        
                        col_sub, col_can = st.columns([1, 1])
                        if col_sub.form_submit_button("确认提交", type="primary"):
                            if not act_handler.strip() or not act_dest.strip():
                                st.error("经办人和去向必填")
                            else:
                                new_photos = []
                                if act_photos:
                                    sn_folder = os.path.join(PHOTO_DIR, sn)
                                    for p in act_photos:
                                        p_name = f"{datetime.now().strftime('%H%M%S')}_{p.name}"
                                        p_path = os.path.join(sn_folder, p_name)
                                        with open(p_path, "wb") as f:
                                            f.write(p.getbuffer())
                                        new_photos.append(p_path)
                                
                                # 继承原来的档案备注，不被覆盖
                                new_row = pd.DataFrame([[
                                    datetime.now().strftime("%Y-%m-%d %H:%M"),
                                    sn, m_type, vendor, name, current_action,
                                    act_dest.strip(), act_handler.strip(), act_note.strip(), ",".join(new_photos),
                                    archive_note
                                ]], columns=cols)
                                new_row.to_csv(CSV_FILE, mode='a', header=False, index=False, encoding='utf-8-sig')
                                del st.session_state[f"action_{sn}_{idx}"]
                                st.rerun()
                        if col_can.form_submit_button("取消"):
                            del st.session_state[f"action_{sn}_{idx}"]
                            st.rerun()

                # 1. 编辑物料档案
                with st.expander("✏️ 编辑物料档案 (改名称/供应商/写备注/追加照片)"):
                    with st.form(key=f"edit_mat_{sn}"):
                        new_name = st.text_input("物料名称", value=name)
                        new_type = st.text_input("物料类型", value=m_type)
                        current_vendor_idx = VENDOR_OPTIONS.index(vendor) if vendor in VENDOR_OPTIONS else VENDOR_OPTIONS.index("其他")
                        new_vendor_sel = st.selectbox("供应商", VENDOR_OPTIONS, index=current_vendor_idx)
                        new_vendor = new_vendor_sel if new_vendor_sel != "其他" else st.text_input("自定义供应商", value=vendor)
                        
                        # ===== 新增的备注区域 =====
                        new_archive_note = st.text_area("📝 备注 / 补充说明 (将显示在卡片外层)", value=archive_note)
                        
                        extra_photos = st.file_uploader("为该物料追加照片", type=["jpg", "png", "jpeg"], accept_multiple_files=True)
                        
                        if st.form_submit_button("保存档案修改"):
                            df.loc[df['SN码'] == sn, '物料名称'] = new_name.strip()
                            df.loc[df['SN码'] == sn, '物料类型'] = new_type.strip()
                            df.loc[df['SN码'] == sn, '供应商'] = new_vendor.strip()
                            
                            # 同步更新所有历史记录的档案备注字段
                            df.loc[df['SN码'] == sn, '档案备注'] = new_archive_note.strip()
                            
                            if extra_photos:
                                sn_folder = os.path.join(PHOTO_DIR, sn)
                                existing_paths = str(latest_record['照片路径列表']).split(',') if latest_record['照片路径列表'] else []
                                if existing_paths == [""]: existing_paths = []
                                for p in extra_photos:
                                    p_name = f"{datetime.now().strftime('%H%M%S')}_{p.name}"
                                    p_path = os.path.join(sn_folder, p_name)
                                    with open(p_path, "wb") as f:
                                        f.write(p.getbuffer())
                                    existing_paths.append(p_path)
                                latest_idx = mat_history.index[0]
                                df.loc[latest_idx, '照片路径列表'] = ",".join(existing_paths)
                                
                            df.to_csv(CSV_FILE, index=False, encoding='utf-8-sig')
                            st.success("物料档案已更新！")
                            st.rerun()
                    
                    # ===== 删除整个物料 =====
                    st.divider()
                    if st.button("🗑️ 彻底删除此物料及其所有记录", key=f"del_all_{sn}", type="primary"):
                        df = df[df['SN码'] != sn]
                        df.to_csv(CSV_FILE, index=False, encoding='utf-8-sig')
                        folder_path = os.path.join(PHOTO_DIR, sn)
                        if os.path.exists(folder_path):
                            shutil.rmtree(folder_path)
                        st.success("物料及照片已彻底删除！")
                        st.rerun()

                # 2. 历史记录
                with st.expander("📜 查看完整历史记录 (可单独修改/删除明细)"):
                    hist_c1, hist_c2 = st.columns([2, 1])
                    with hist_c1:
                        for h_idx, r in mat_history.iterrows():
                            color = "green" if r['动作'] == "入库" else "red"
                            st.markdown(f"**:{color}[{r['动作']}]** | {r['时间']} | {r['去向/来源']} | {r['经办人']} | 目的: {r['备注']}")
                            
                            with st.expander("修改或删除这条记录"):
                                with st.form(key=f"edit_hist_{sn}_{h_idx}"):
                                    new_h_dest = st.text_input("去向/地点", value=r['去向/来源'])
                                    new_h_handler = st.text_input("经办人", value=r['经办人'])
                                    new_h_note = st.text_input("目的/备注", value=r['备注'])
                                    if st.form_submit_button("保存单条修改"):
                                        df.loc[h_idx, '去向/来源'] = new_h_dest.strip()
                                        df.loc[h_idx, '经办人'] = new_h_handler.strip()
                                        df.loc[h_idx, '备注'] = new_h_note.strip()
                                        df.to_csv(CSV_FILE, index=False, encoding='utf-8-sig')
                                        st.success("单条明细已更新！")
                                        st.rerun()
                                
                                if st.button("🗑️ 删除这条流水", key=f"del_hist_{sn}_{h_idx}"):
                                    df = df.drop(h_idx)
                                    df.to_csv(CSV_FILE, index=False, encoding='utf-8-sig')
                                    st.success("该条流水已删除！")
                                    st.rerun()
                    with hist_c2:
                        if all_photos:
                            st.write("**关联照片：**")
                            for p in all_photos:
                                st.image(p, use_container_width=True)