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
        st.error("⚠️ 缺少 Supabase Secrets 配置，请检查后台。")
        st.stop()

supabase = init_supabase()
TABLE_NAME = "Material"          
BUCKET_NAME = "material-photos"  

# ================= 2. 严格对齐你的表头命名 =================
C_TIME = "Time"
C_SN = "SN (自定义)"
C_TYPE = "物料类型/Material type"
C_VENDOR = "供应商/Vendor"
C_NAME = "物料名称/Material name"
C_FROM = "来自/Shipped from"
C_CURRENT = "暂存于/Current location"
C_OP_IN = "操作员Operator_In"
C_TO = "寄往/Shipped to"
C_TARGET = "目的地/Target location"
C_OP_OUT = "操作员Operator_Out"
C_NOTE = "备注/Notes"
C_PHOTOS = "图像路径/Image path"

cols = [C_TIME, C_SN, C_TYPE, C_VENDOR, C_NAME, C_FROM, C_CURRENT, C_OP_IN, C_TO, C_TARGET, C_OP_OUT, C_NOTE, C_PHOTOS]

# ================= 3. 从云端读取数据 =================
@st.cache_data(ttl=3) 
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
        st.error(f"读取云端数据失败: {e}")
        return pd.DataFrame(columns=cols)

df = load_data()

# ================= 4. 生成带图 Excel =================
def export_excel_with_images(df):
    wb = Workbook()
    ws = wb.active
    ws.title = "物料清单-Material List"
    
    unique_sns = df[C_SN].drop_duplicates()
    max_photos_count = 0
    export_data = []
    
    for sn in unique_sns:
        mat_history = df[df[C_SN] == sn].copy()
        mat_history[C_TIME] = pd.to_datetime(mat_history[C_TIME], errors='coerce')
        mat_history = mat_history.sort_values(C_TIME, ascending=False)
        
        latest = mat_history.iloc[0]
        
        # 通过 Operator_In 和 Operator_Out 判断历史
        last_in = mat_history[mat_history[C_OP_IN] != ""].head(1)
        last_out = mat_history[mat_history[C_OP_OUT] != ""].head(1)
        
        in_str = f"{last_in.iloc[0][C_TIME].strftime('%Y-%m-%d %H:%M')} | {last_in.iloc[0][C_FROM]}" if not last_in.empty else "无 / None"
        out_str = f"{last_out.iloc[0][C_TIME].strftime('%Y-%m-%d %H:%M')} | {last_out.iloc[0][C_TO]}" if not last_out.empty else "无 / None"
        
        # 判断状态：最新一条记录如果是Out操作，就是已出库
        status = "已出库 / Out" if latest[C_OP_OUT] != "" else "在库 / In"
        
        all_imgs = []
        for p_list in mat_history[C_PHOTOS]:
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
                str(sn), str(latest[C_NAME]), str(latest[C_TYPE]), 
                str(latest[C_VENDOR]), status,
                str(in_str), str(out_str), str(latest[C_NOTE])
            ],
            "imgs": all_imgs
        })
        
    headers = ["SN码 / SN ", "物料名称 / Name", "类型 / Type", "供应商 / Vendor", "状态 / Status", "最近入库 / Last In", "最近出库 / Last Out", "备注 / Remarks"]
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
                resp = requests.get(img_url, timeout=5)
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
                
                ws.add_image(xl_img, f"{chr(73 + i)}{current_row}")
            except Exception:
                pass
                
    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer

# ================= 5. 侧边栏 UI =================
with st.sidebar:
    st.header("📝 登记新物料 / Register New Material")
    
    material = st.text_input("物料名称 / Material Name *")
    
    default_types = ["Sensor", "Lens", "IRC", "PCB", "机构件 / Mechanical"]
    if not df.empty:
        existing_types = [t for t in df[C_TYPE].unique() if t and t != "nan"]
        for t in existing_types:
            if t not in default_types: default_types.append(t)
    
    type_sel = st.selectbox("物料类型 / Material Type", default_types + ["其他 / Other"])
    if "其他" in type_sel:
        final_type = st.text_input("请输入自定义类型 / Enter Custom Type")
    else:
        final_type = type_sel

    default_vendors = ["GC", "Qtech", "Newmax", "Omnivision"]
    if not df.empty:
        existing_vendors = [v for v in df[C_VENDOR].unique() if v and v != "nan"]
        for v in existing_vendors:
            if v not in default_vendors: default_vendors.append(v)
                
    vendor_sel = st.selectbox("供应商 / Vendor", default_vendors + ["其他 / Other"])
    if "其他" in vendor_sel:
        final_vendor = st.text_input("请输入自定义供应商 / Enter Custom Vendor")
    else:
        final_vendor = vendor_sel

    handler = st.text_input("入库经办人 / Inbound Operator *", value="Jerry Wang")
    destination = st.text_input("入库地点/来源 / Inbound Location/Source *", value="SH office")
    photos = st.file_uploader("上传照片 / Upload Photos", type=["jpg", "png", "jpeg"], accept_multiple_files=True)
    note = st.text_area("备注/目的 / Remarks/Purpose")
    
    if st.button("提交登记 / Submit Registration", type="primary", use_container_width=True):
        if not material.strip() or not handler.strip():
            st.error("物料名称和经办人必填！ / Required!")
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
                        # 使用 p.getvalue() 防止 memoryview 报错
                        supabase.storage.from_(BUCKET_NAME).upload(path=file_path, file=p.getvalue(), file_options={"content-type": p.type})
                        img_url = supabase.storage.from_(BUCKET_NAME).get_public_url(file_path)
                        saved_photos_urls.append(img_url)
                    except Exception as e:
                        st.error(f"图片上传失败 / Upload failed: {e}")
            
            # 严格根据你的表头插入数据
            new_data = {
                C_TIME: datetime.now().strftime("%Y-%m-%d %H:%M"),
                C_SN: sn, C_TYPE: final_type, C_VENDOR: final_vendor, 
                C_NAME: material.strip(), 
                C_FROM: destination.strip(), 
                C_CURRENT: destination.strip(), 
                C_OP_IN: handler.strip(), 
                C_TO: "", C_TARGET: "", C_OP_OUT: "",  # 入库时不填出库字段
                C_NOTE: note.strip(), 
                C_PHOTOS: ",".join(saved_photos_urls)
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
    st.header("📤 导出给同事 / Export")
    if st.button("生成带图 Excel / Generate Excel", use_container_width=True):
        with st.spinner("正在下载原图并打包..."):
            excel_data = export_excel_with_images(df)
            st.download_button(
                label="👇 点击这里下载 .xlsx / Download",
                data=excel_data,
                file_name=f"物料总览_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )

# ================= 6. 主界面展示 =================
st.markdown("### 🔍 筛选与查找 / Filter & Search")
f_col1, f_col2, f_col3 = st.columns([2, 1, 1])
with f_col1: search = st.text_input("关键词 / Keyword")
with f_col2:
    all_types = ["全部 / All"] + sorted(list(df[C_TYPE].unique()))
    type_filter = st.selectbox("按物料类型 / By Type", all_types)
with f_col3:
    all_vendors = ["全部 / All"] + sorted(list(df[C_VENDOR].unique()))
    vendor_filter = st.selectbox("按供应商 / By Vendor", all_vendors)

unique_sns = df[[C_SN]].drop_duplicates().sort_values(C_SN)
if search:
    unique_sns = unique_sns[unique_sns[C_SN].apply(lambda sn: search.lower() in df[df[C_SN]==sn].astype(str).to_string().lower())]
if type_filter != "全部 / All":
    unique_sns = unique_sns[unique_sns[C_SN].isin(df[df[C_TYPE] == type_filter][C_SN].unique())]
if vendor_filter != "全部 / All":
    unique_sns = unique_sns[unique_sns[C_SN].isin(df[df[C_VENDOR] == vendor_filter][C_SN].unique())]

st.divider()

if unique_sns.empty:
    st.info("没有找到符合条件的物料。 / No materials found.")
else:
    for idx, (_, mat_row) in enumerate(unique_sns.iterrows()):
        sn = mat_row[C_SN]
        mat_history = df[df[C_SN] == sn].copy().sort_values(C_TIME, ascending=False)
        latest_record = mat_history.iloc[0]
        name, m_type, vendor = latest_record[C_NAME], latest_record[C_TYPE], latest_record[C_VENDOR]
        
        # 筛选出最近的入库和出库记录
        last_in = mat_history[mat_history[C_OP_IN] != ""].head(1)
        last_out = mat_history[mat_history[C_OP_OUT] != ""].head(1)
        
        # 如果最新一条记录的操作员是 Out，则为已出库
        is_out = latest_record[C_OP_OUT] != ""
        status = "已出库 / Out 🔴" if is_out else "在库 / In 🟢"
        status_color = "red" if is_out else "green"

        all_photos = []
        for p_list in mat_history[C_PHOTOS]:
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
                        st.markdown(f"时间 / Time: {r[C_TIME]}  \n地点 / Loc: {r[C_FROM]}  \n经办人 / Op: {r[C_OP_IN]}  \n目的 / Purpose: {r[C_NOTE]}")
                    else: 
                        st.write("无记录 / No record")
                        
                with detail_col2:
                    st.markdown("**📤 最近出库 / Last Outbound**")
                    if not last_out.empty:
                        r = last_out.iloc[0]
                        st.markdown(f"时间 / Time: {r[C_TIME]}  \n去向 / Dest: {r[C_TO]}  \n经办人 / Op: {r[C_OP_OUT]}  \n目的 / Purpose: {r[C_NOTE]}")
                    else: 
                        st.write("无出库记录 / No outbound record")
                
                st.divider()
                
                btn_col1, btn_col2, btn_col3 = st.columns([1, 1, 2])
                with btn_col1:
                    if st.button("📤 借出/出库 / Borrow", key=f"btn_out_{sn}_{idx}", disabled=is_out):
                        st.session_state[f"action_{sn}_{idx}"] = "出库"
                with btn_col2:
                    if st.button("📥 归还/入库 / Return", key=f"btn_in_{sn}_{idx}", disabled=not is_out):
                        st.session_state[f"action_{sn}_{idx}"] = "入库"
                
                current_action = st.session_state.get(f"action_{sn}_{idx}")
                if current_action:
                    action_zh = "出库 / Outbound" if current_action == "出库" else "入库 / Inbound"
                    with st.form(key=f"form_action_{sn}_{idx}"):
                        st.info(f"正在执行 / Executing：{action_zh}")
                        act_handler = st.text_input("经办人 / Operator *", key=f"h_{sn}_{idx}", value="Jerry Wang")
                        act_dest = st.text_input("去向/地点 / Location *", key=f"d_{sn}_{idx}", value="借用人/部门 / Borrower" if current_action=="出库" else "SH office")
                        act_note = st.text_input("目的/备注 / Purpose", key=f"n_{sn}_{idx}")
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
                                            supabase.storage.from_(BUCKET_NAME).upload(path=file_path, file=p.getvalue(), file_options={"content-type": p.type})
                                            img_url = supabase.storage.from_(BUCKET_NAME).get_public_url(file_path)
                                            new_photos_urls.append(img_url)
                                        except: pass
                                
                                # 根据动作填充对应的列
                                if current_action == "出库":
                                    new_data = {
                                        C_TIME: datetime.now().strftime("%Y-%m-%d %H:%M"), C_SN: sn, 
                                        C_TYPE: m_type, C_VENDOR: vendor, C_NAME: name, 
                                        C_FROM: "", C_CURRENT: "",
                                        C_OP_IN: "", 
                                        C_TO: act_dest.strip(), C_TARGET: act_dest.strip(), C_OP_OUT: act_handler.strip(),
                                        C_NOTE: act_note.strip(), C_PHOTOS: ",".join(new_photos_urls)
                                    }
                                else:
                                    new_data = {
                                        C_TIME: datetime.now().strftime("%Y-%m-%d %H:%M"), C_SN: sn, 
                                        C_TYPE: m_type, C_VENDOR: vendor, C_NAME: name, 
                                        C_FROM: act_dest.strip(), C_CURRENT: act_dest.strip(), C_OP_IN: act_handler.strip(), 
                                        C_TO: "", C_TARGET: "", C_OP_OUT: "", 
                                        C_NOTE: act_note.strip(), C_PHOTOS: ",".join(new_photos_urls)
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
                        
                        # 追加照片
                        extra_photos = st.file_uploader("追加照片 / Add Photos", type=["jpg", "png", "jpeg"], accept_multiple_files=True)
                        
                        if st.form_submit_button("保存修改 / Save Changes"):
                            update_data = {C_NAME: new_name.strip(), C_TYPE: new_type.strip(), C_VENDOR: new_vendor.strip()}
                            
                            if extra_photos:
                                new_urls = []
                                for p in extra_photos:
                                    p_name = f"{datetime.now().strftime('%H%M%S')}_{p.name}"
                                    file_path = f"{sn}/{p_name}"
                                    try:
                                        supabase.storage.from_(BUCKET_NAME).upload(path=file_path, file=p.getvalue(), file_options={"content-type": p.type})
                                        img_url = supabase.storage.from_(BUCKET_NAME).get_public_url(file_path)
                                        new_urls.append(img_url)
                                    except: pass
                                
                                # 把追加的照片写进最新一条记录
                                existing_paths = str(latest_record[C_PHOTOS]).split(',') if latest_record[C_PHOTOS] else []
                                if existing_paths == [""]: existing_paths = []
                                existing_paths.extend(new_urls)
                                
                                try:
                                    # 单独更新最新记录的图片
                                    supabase.table(TABLE_NAME).update({C_PHOTOS: ",".join(existing_paths)}).eq("id", latest_record['id']).execute()
                                except: pass
                            
                            try:
                                # 批量更新所有该 SN 码的名称、类型、供应商
                                supabase.table(TABLE_NAME).update(update_data).eq(C_SN, sn).execute()
                                st.success("档案已更新 / Updated!")
                                st.cache_data.clear()
                                st.rerun()
                            except Exception as e:
                                st.error(f"更新失败: {e}")
                    
                    st.divider()
                    if st.button("🗑️ 彻底删除此物料 / Permanently Delete", key=f"del_all_{sn}", type="primary"):
                        try:
                            supabase.table(TABLE_NAME).delete().eq(C_SN, sn).execute()
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
                            # 根据操作员字段判断动作
                            if r[C_OP_OUT] != "":
                                color, act_zh, dest, op = "red", "出库 / Out", r[C_TO], r[C_OP_OUT]
                            else:
                                color, act_zh, dest, op = "green", "入库 / In", r[C_FROM], r[C_OP_IN]
                                
                            st.markdown(f"**:{color}[{act_zh}]** | {r[C_TIME]} | {dest} | {op} | 备注: {r[C_NOTE]}")
                            
                            with st.expander("修改/删除 / Edit/Delete"):
                                with st.form(key=f"edit_hist_{sn}_{h_idx}"):
                                    new_h_dest = st.text_input("去向/地点 / Loc", value=dest)
                                    new_h_handler = st.text_input("经办人 / Op", value=op)
                                    new_h_note = st.text_input("目的/备注 / Remark", value=r[C_NOTE])
                                    if st.form_submit_button("保存 / Save"):
                                        try:
                                            # 判断是更新出库还是入库字段
                                            if r[C_OP_OUT] != "":
                                                update_dict = {C_TO: new_h_dest.strip(), C_TARGET: new_h_dest.strip(), C_OP_OUT: new_h_handler.strip(), C_NOTE: new_h_note.strip()}
                                            else:
                                                update_dict = {C_FROM: new_h_dest.strip(), C_CURRENT: new_h_dest.strip(), C_OP_IN: new_h_handler.strip(), C_NOTE: new_h_note.strip()}
                                                
                                            supabase.table(TABLE_NAME).update(update_dict).eq("id", r['id']).execute()
                                            st.success("已更新 / Updated!")
                                            st.cache_data.clear()
                                            st.rerun()
                                        except Exception as e:
                                            st.error(f"失败: {e}")
                                
                                if st.button("🗑️ 删除这条 / Delete", key=f"del_hist_{sn}_{h_idx}"):
                                    try:
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