import uuid
import time
import io
from datetime import datetime
from PIL import Image as PILImage
from openpyxl import Workbook
from openpyxl.drawing.image import Image as OpenpyxlImage

import streamlit as st
import pandas as pd
import requests
from supabase import create_client, Client

st.set_page_config(page_title="物料管理 / Material Tracker", layout="wide")
st.title("📦 物料档案(上海) / Material Management(Only Shanghai office)")

# ================= 1. 连接 Supabase 云端 =================
@st.cache_resource
def init_supabase():
    try:
        url = st.secrets["SUPABASE_URL"]
        key = st.secrets["SUPABASE_KEY"]
        return create_client(url, key)
    except Exception as e:
        st.error("⚠️ 缺少 Supabase Secrets 配置，请在 Streamlit 后台设置 SUPABASE_URL 和 SUPABASE_KEY。")
        st.stop()

supabase = init_supabase()
TABLE_NAME = "Material"          # 你的 Supabase 表名 (注意大小写)
BUCKET_NAME = "material-photos"  # 你的图片桶名

cols = ["时间", "SN码", "物料类型", "供应商", "物料名称", "动作", "去向/来源", "经办人", "备注", "照片路径列表", "档案备注"]

# ================= 2. 从云端读取数据 =================
@st.cache_data(ttl=3) # 缓存 3 秒
def load_data():
    try:
        response = supabase.table(TABLE_NAME).select("*").order('id', desc=True).execute()
        df = pd.DataFrame(response.data)
        if df.empty:
            return pd.DataFrame(columns=cols)
        for c in cols:
            if c not in df.columns:
                df[c] = ""
        return df
    except Exception as e:
        st.error(f"读取云端数据失败 / Failed to load data: {e}")
        return pd.DataFrame(columns=cols)

df = load_data()

# ================= 3. 生成带图 Excel (从云端 URL 下载图片) =================
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
                    p_str = p.strip()
                    if p_str.startswith("http"):
                        all_imgs.append(p_str)
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
        
    headers = ["SN码 / SN ", "物料名称 / Material Name", "类型 / Type", "供应商 / Vendor", "状态 / Status", "地点 / Location", "记录时间 / Time", "备注 / Remarks"]
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

        for i, img_url in enumerate(item["imgs"]):
            try:
                # 从云端 URL 下载图片并写入 Excel
                resp = requests.get(img_url, timeout=5)
                resp.raise_for_status()
                pil_img = PILImage.open(io.BytesIO(resp.content))
                if pil_img.mode in ("RGBA", "P"): pil_img = pil_img.convert("RGB")
                pil_img.thumbnail((600, 600), PILImage.Resampling.LANCZOS) 
                img_byte_arr = io.BytesIO()
                pil_img.save(img_byte_arr, format='JPEG')
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

# ================= 4. 侧边栏 UI =================
with st.sidebar:
    st.header("📝 登记新物料 / Register New Material")
    
    material = st.text_input("物料名称 / Material Name *")
    
    default_types = ["Sensor", "Lens", "IRC", "PCB", "机构件 / Mechanical"]
    if not df.empty:
        existing_types = [t for t in df["物料类型"].unique() if t and t != "nan"]
        for t in existing_types:
            if t not in default_types: default_types.append(t)
    
    type_sel = st.selectbox("物料类型 / Material Type", default_types + ["其他 / Other"])
    if "其他" in type_sel:
        final_type = st.text_input("请输入自定义类型 / Enter Custom Type")
    else:
        final_type = type_sel

    default_vendors = ["GC", "Qtech", "Newmax", "Omnivision"]
    if not df.empty:
        existing_vendors = [v for v in df["供应商"].unique() if v and v != "nan"]
        for v in existing_vendors:
            if v not in default_vendors: default_vendors.append(v)
                
    vendor_sel = st.selectbox("供应商 / Vendor", default_vendors + ["其他 / Other"])
    if "其他" in vendor_sel:
        final_vendor = st.text_input("请输入自定义供应商 / Enter Custom Vendor")
    else:
        final_vendor = vendor_sel

    handler = st.text_input("入库经办人 / Inbound Operator *", value="Jerry Wang")
    destination = st.text_input("入库地点/来源 / Inbound Location/Source *", value="SH office")
    photos = st.file_uploader("上传照片 (支持多选) / Upload Photos (Multiple)", type=["jpg", "png", "jpeg"], accept_multiple_files=True)
    note = st.text_area("备注/目的 / Remarks/Purpose")
    
    if st.button("提交登记 / Submit Registration", type="primary", use_container_width=True):
        if not material.strip() or not handler.strip():
            st.error("物料名称和经办人必填！ / Material Name and Operator are required!")
        elif "其他" in type_sel and not final_type.strip():
            st.error("请填写自定义类型 / Please enter custom type")
        elif "其他" in vendor_sel and not final_vendor.strip():
            st.error("请填写自定义供应商 / Please enter custom vendor")
        else:
            sn = f"MAT-{uuid.uuid4().hex[:6].upper()}"
            saved_photos_urls = []
            
            if photos:
                for p in photos:
                    p_name = f"{datetime.now().strftime('%H%M%S')}_{p.name}"
                    file_path = f"{sn}/{p_name}"
                    try:
                        supabase.storage.from_(BUCKET_NAME).upload(
                            path=file_path,
                            file=p.getbuffer(),
                            file_options={"content-type": p.type}
                        )
                        img_url = supabase.storage.from_(BUCKET_NAME).get_public_url(file_path)
                        saved_photos_urls.append(img_url)
                    except Exception as e:
                        st.error(f"图片上传失败 / Upload failed: {e}")
            
            new_data = {
                "时间": datetime.now().strftime("%Y-%m-%d %H:%M"),
                "SN码": sn, "物料类型": final_type, "供应商": final_vendor, 
                "物料名称": material.strip(), "动作": "入库", "去向/来源": destination.strip(), 
                "经办人": handler.strip(), "备注": note.strip(), 
                "照片路径列表": ",".join(saved_photos_urls), "档案备注": ""
            }
            
            try:
                supabase.table(TABLE_NAME).insert(new_data).execute()
                st.success(f"登记成功 / Successful! SN: {sn}")
                st.cache_data.clear()
                time.sleep(0.5)
                st.rerun()
            except Exception as e:
                st.error(f"保存失败 / Save failed: {e}")
                
    st.divider()
    st.header("📤 导出给同事 / Export for Colleagues")
    if st.button("生成带图 Excel / Generate Excel", use_container_width=True):
        with st.spinner("正在下载云端原图并打包... / Packing..."):
            excel_data = export_excel_with_images(df)
            st.download_button(
                label="👇 点击这里下载 .xlsx / Download",
                data=excel_data,
                file_name=f"物料总览_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )

# ================= 5. 主界面展示 =================
st.markdown("### 🔍 筛选与查找 / Filter & Search")
f_col1, f_col2, f_col3 = st.columns([2, 1, 1])
with f_col1: search = st.text_input("关键词 / Keyword")
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
    st.info("没有找到符合条件的物料。 / No materials found.")
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
                    if p.strip().startswith("http"): all_photos.append(p.strip())
        all_photos = list(dict.fromkeys(all_photos))

        with st.container(border=True):
            c1, c2 = st.columns([1.5, 3])
            with c1:
                if all_photos:
                    img_cols = st.columns(2)
                    for i, p_url in enumerate(all_photos[:2]):
                        img_cols[i].image(p_url, use_container_width=True)
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
                
                if archive_note and archive_note != "nan":
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
                        act_note = st.text_input("本次动作目的/备注 / Purpose", key=f"n_{sn}_{idx}")
                        act_photos = st.file_uploader("补充照片 / Photos", type=["jpg", "png", "jpeg"], accept_multiple_files=True, key=f"p_{sn}_{idx}")
                        
                        col_sub, col_can = st.columns([1, 1])
                        if col_sub.form_submit_button("确认提交 / Confirm", type="primary"):
                            if not act_handler.strip() or not act_dest.strip():
                                st.error("经办人和去向必填 / Required")
                            else:
                                new_photos_urls = []
                                if act_photos:
                                    for p in act_photos:
                                        p_name = f"{datetime.now().strftime('%H%M%S')}_{p.name}"
                                        file_path = f"{sn}/{p_name}"
                                        try:
                                            supabase.storage.from_(BUCKET_NAME).upload(path=file_path, file=p.getbuffer(), file_options={"content-type": p.type})
                                            img_url = supabase.storage.from_(BUCKET_NAME).get_public_url(file_path)
                                            new_photos_urls.append(img_url)
                                        except: pass
                                
                                new_data = {
                                    "时间": datetime.now().strftime("%Y-%m-%d %H:%M"), "SN码": sn, 
                                    "物料类型": m_type, "供应商": vendor, "物料名称": name, 
                                    "动作": current_action, "去向/来源": act_dest.strip(), 
                                    "经办人": act_handler.strip(), "备注": act_note.strip(), 
                                    "照片路径列表": ",".join(new_photos_urls), "档案备注": archive_note
                                }
                                try:
                                    supabase.table(TABLE_NAME).insert(new_data).execute()
                                    del st.session_state[f"action_{sn}_{idx}"]
                                    st.cache_data.clear()
                                    st.rerun()
                                except Exception as e:
                                    st.error(f"保存失败: {e}")
                        if col_can.form_submit_button("取消 / Cancel"):
                            del st.session_state[f"action_{sn}_{idx}"]
                            st.rerun()

                # ===== 编辑物料档案 =====
                with st.expander("✏️ 编辑物料档案 / Edit Profile"):
                    with st.form(key=f"edit_mat_{sn}"):
                        new_name = st.text_input("物料名称 / Name", value=name)
                        new_type = st.text_input("物料类型 / Type", value=m_type)
                        new_vendor = st.text_input("供应商 / Vendor", value=vendor)
                        new_archive_note = st.text_area("📝 备注 / Remarks", value=archive_note)
                        extra_photos = st.file_uploader("追加照片 / Add Photos", type=["jpg", "png", "jpeg"], accept_multiple_files=True)
                        
                        if st.form_submit_button("保存修改 / Save Changes"):
                            update_data = {
                                "物料名称": new_name.strip(), "物料类型": new_type.strip(), 
                                "供应商": new_vendor.strip(), "档案备注": new_archive_note.strip()
                            }
                            if extra_photos:
                                new_urls = []
                                for p in extra_photos:
                                    p_name = f"{datetime.now().strftime('%H%M%S')}_{p.name}"
                                    file_path = f"{sn}/{p_name}"
                                    try:
                                        supabase.storage.from_(BUCKET_NAME).upload(path=file_path, file=p.getbuffer(), file_options={"content-type": p.type})
                                        img_url = supabase.storage.from_(BUCKET_NAME).get_public_url(file_path)
                                        new_urls.append(img_url)
                                    except: pass
                                existing_paths = str(latest_record['照片路径列表']).split(',') if latest_record['照片路径列表'] else []
                                if existing_paths == [""]: existing_paths = []
                                existing_paths.extend(new_urls)
                                update_data["照片路径列表"] = ",".join(existing_paths)
                            
                            try:
                                # 批量更新所有该 SN 码的记录
                                supabase.table(TABLE_NAME).update(update_data).eq("SN码", sn).execute()
                                st.success("档案已更新 / Updated!")
                                st.cache_data.clear()
                                st.rerun()
                            except Exception as e:
                                st.error(f"更新失败: {e}")
                    
                    st.divider()
                    if st.button("🗑️ 彻底删除此物料 / Permanently Delete", key=f"del_all_{sn}", type="primary"):
                        try:
                            # 删除数据库记录 (图片文件暂存在 Storage 中不会自动删，防止误删)
                            supabase.table(TABLE_NAME).delete().eq("SN码", sn).execute()
                            st.success("已删除 / Deleted!")
                            st.cache_data.clear()
                            st.rerun()
                        except Exception as e:
                            st.error(f"删除失败: {e}")

                # ===== 查看历史记录 =====
                with st.expander("📜 查看完整历史记录 / History"):
                    hist_c1, hist_c2 = st.columns([2, 1])
                    with hist_c1:
                        for h_idx, r in mat_history.iterrows():
                            color = "green" if r['动作'] == "入库" else "red"
                            act_zh = "入库 / In" if r['动作'] == "入库" else "出库 / Out"
                            st.markdown(f"**:{color}[{act_zh}]** | {r['时间']} | {r['去向/来源']} | {r['经办人']} | 备注: {r['备注']}")
                            
                            with st.expander("修改/删除 / Edit/Delete"):
                                with st.form(key=f"edit_hist_{sn}_{h_idx}"):
                                    new_h_dest = st.text_input("去向/地点 / Loc", value=r['去向/来源'])
                                    new_h_handler = st.text_input("经办人 / Op", value=r['经办人'])
                                    new_h_note = st.text_input("目的/备注 / Remark", value=r['备注'])
                                    if st.form_submit_button("保存 / Save"):
                                        try:
                                            # 通过 id 字段精确更新某一条
                                            supabase.table(TABLE_NAME).update({
                                                "去向/来源": new_h_dest.strip(), 
                                                "经办人": new_h_handler.strip(), 
                                                "备注": new_h_note.strip()
                                            }).eq("id", r['id']).execute()
                                            st.success("已更新 / Updated!")
                                            st.cache_data.clear()
                                            st.rerun()
                                        except Exception as e:
                                            st.error(f"失败: {e}")
                                
                                if st.button("🗑️ 删除这条 / Delete", key=f"del_hist_{sn}_{h_idx}"):
                                    try:
                                        # 通过 id 字段精确删除某一条
                                        supabase.table(TABLE_NAME).delete().eq("id", r['id']).execute()
                                        st.success("已删除 / Deleted!")
                                        st.cache_data.clear()
                                        st.rerun()
                                    except Exception as e:
                                        st.error(f"失败: {e}")
                    with hist_c2:
                        if all_photos:
                            st.write("**关联照片 / Photos：**")
                            for p_url in all_photos:
                                st.image(p_url, use_container_width=True)