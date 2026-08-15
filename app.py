import os
import io
import json
import base64
import requests
from io import BytesIO
from typing import Optional, List
from fastapi import FastAPI, Request, HTTPException, UploadFile, File, Form
from fastapi.responses import HTMLResponse, StreamingResponse, JSONResponse
from fastapi.templating import Jinja2Templates
from fastapi.staticfiles import StaticFiles
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Inches
from PIL import Image, ImageOps

# -------------------------------------------------------------------
# ⚙️ วาง Web App URL (Apps Script) ของคุณที่นี่
# -------------------------------------------------------------------
APPS_SCRIPT_URL = "https://script.google.com/macros/s/AKfycbyovncjGanRbNjUjLtQNZb_MKIwFrvovGkxeDUXr_tYYcnEVe6nXWxPt-J6RV7mHJ-8Qw/exec"

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
CACHE_DIR = os.path.join(BASE_DIR, "cache")
TEMPLATES_DIR = os.path.join(BASE_DIR, "templates")
LOCAL_IMAGES_DIR = os.path.join(BASE_DIR, "local_images")
os.makedirs(CACHE_DIR, exist_ok=True)
os.makedirs(LOCAL_IMAGES_DIR, exist_ok=True)

CACHE_DB_FILE = os.path.join(CACHE_DIR, "database_cache.json")

app = FastAPI(title="SAT Command Center Enterprise")
templates = Jinja2Templates(directory=TEMPLATES_DIR)

# -------------------------------------------------------------------
# Helper Functions: Image Resizing & Apps Script Bridge
# -------------------------------------------------------------------
def call_apps_script(payload: dict) -> dict:
    try:
        resp = requests.post(APPS_SCRIPT_URL, json=payload, timeout=20)
        if resp.status_code == 200:
            return resp.json()
        return {"success": False, "error": f"HTTP {resp.status_code}"}
    except Exception as e:
        return {"success": False, "error": str(e), "is_offline": True}

def fit_image_to_strict_bounds(img_bytes, target_width_in=1.9, target_height_in=2.0):
    try:
        img = Image.open(BytesIO(img_bytes))
        img = ImageOps.exif_transpose(img)
        if img.mode != 'RGB':
            img = img.convert('RGB')
        w_px, h_px = img.size
        if w_px == 0 or h_px == 0:
            return img_bytes, target_width_in, target_height_in

        aspect = h_px / w_px
        final_w_in = target_width_in
        final_h_in = final_w_in * aspect
        
        if final_h_in > target_height_in:
            final_h_in = target_height_in
            final_w_in = final_h_in / aspect if aspect != 0 else target_width_in

        if final_w_in > target_width_in:
            final_w_in = target_width_in
        
        target_w_px = max(1, int(final_w_in * 300))
        target_h_px = max(1, int(final_h_in * 300))
        
        img = img.resize((target_w_px, target_h_px), Image.Resampling.LANCZOS)
        buf = BytesIO()
        img.save(buf, format='JPEG', quality=95)
        buf.seek(0)
        return buf, final_w_in, final_h_in
    except Exception:
        return BytesIO(img_bytes), target_width_in, target_height_in

# -------------------------------------------------------------------
# API Endpoints
# -------------------------------------------------------------------
@app.get("/", response_class=HTMLResponse)
async def serve_ui(request: Request):
    return templates.TemplateResponse(request=request, name="index.html")

@app.post("/api/login")
async def login(payload: dict):
    username = payload.get("username", "").strip()
    password = payload.get("password", "").strip()
    
    # 1. พยายามยืนยันตัวตนผ่าน Google Sheets (Online)
    res = call_apps_script({"action": "login", "username": username, "password": password})
    if res.get("success"):
        return {"success": True, "user": res.get("user"), "is_offline": False}
    
    # 2. ถ้าเน็ตหลุด (Offline) ให้ตรวจจาก Cache ในเครื่อง
    if res.get("is_offline") and os.path.exists(CACHE_DB_FILE):
        try:
            with open(CACHE_DB_FILE, "r", encoding="utf-8") as f:
                cached_db = json.load(f)
            user_sheet = cached_db.get("User", [])
            for row in user_sheet[1:]:
                if len(row) >= 2 and str(row[0]).strip() == username and str(row[1]).strip() == password:
                    return {"success": True, "user": username, "is_offline": True, "warning": "Logged in via Offline Cache"}
        except Exception:
            pass

    return JSONResponse(status_code=401, content={"success": False, "message": res.get("message", "Login Failed")})

@app.get("/api/database")
async def get_database():
    # 1. ลองดึงจาก Google Sheets ล่าสุด
    res = call_apps_script({"action": "get_database"})
    if res.get("success"):
        # บันทึกลง Local Cache อัตโนมัติ
        with open(CACHE_DB_FILE, "w", encoding="utf-8") as f:
            json.dump(res.get("data", {}), f, ensure_ascii=False, indent=2)
        return {"success": True, "data": res.get("data"), "is_offline": False}
    
    # 2. ถ้าต่อเน็ตไม่ได้ ดึงจาก Offline Cache
    if os.path.exists(CACHE_DB_FILE):
        with open(CACHE_DB_FILE, "r", encoding="utf-8") as f:
            cached_data = json.load(f)
        return {"success": True, "data": cached_data, "is_offline": True}
    
    raise HTTPException(status_code=503, detail="Database Unavailable (No Internet & No Cache)")

@app.post("/api/upload")
async def upload_file(
    topic: str = Form(...),
    phase: str = Form(...),
    room: str = Form(...),
    eq: str = Form(...),
    subject: str = Form(...),
    file: UploadFile = File(...)
):
    contents = await file.read()
    b64_str = base64.b64encode(contents).decode("utf-8")
    
    # ส่งขึ้น Google Drive ผ่าน Apps Script
    payload = {
        "action": "upload_image",
        "topic": topic, "phase": phase, "room": room,
        "eq": eq, "subject": subject,
        "filename": file.filename,
        "base64Data": b64_str
    }
    res = call_apps_script(payload)
    
    # บันทึกสำเนาลงเครื่องไว้ด้วย (Offline Copy)
    local_sub_dir = os.path.join(LOCAL_IMAGES_DIR, topic, phase, room, eq, f"Subject_{subject}")
    os.makedirs(local_sub_dir, exist_ok=True)
    with open(os.path.join(local_sub_dir, file.filename), "wb") as f:
        f.write(contents)

    return res

@app.post("/api/export/word")
async def export_word(payload: dict):
    try:
        topic = payload.get("topic")
        phase = payload.get("phase")
        room = payload.get("room")
        test_date = payload.get("test_date", "13-Jul-2026")
        revision = payload.get("revision", "R0")
        template_name = payload.get("template_name", "Template.docx")
        
        # 1. ดาวน์โหลด Template จาก Google Drive หรือ Local Cache
        tmpl_res = call_apps_script({"action": "get_template", "templateName": template_name})
        if tmpl_res.get("success"):
            tmpl_bytes = base64.b64decode(tmpl_res.get("base64"))
        else:
            local_tmpl = os.path.join(BASE_DIR, template_name)
            if os.path.exists(local_tmpl):
                with open(local_tmpl, "rb") as f: tmpl_bytes = f.read()
            else:
                raise HTTPException(status_code=404, detail="Word Template Not Found")

        tpl = DocxTemplate(BytesIO(tmpl_bytes))
        
        context = {
            'project_name': "THAI DC 1-RYG",
            'location': room,
            'test_date': test_date,
            'R': revision, 'revision': revision,
            'room_name': room,
            'checked_by': payload.get("user", "Tian Jian"),
            'witnessed_by': "", 'approved_by': "",
            'equipment_table': payload.get("equipment_table", []),
            'equipment_list': []
        }
        
        # ประกอบข้อมูลรูปภาพลงใน Context
        for eq_item in payload.get("equipment_data", []):
            eq_code = eq_item.get("code")
            eq_dict = {'equipment_name': eq_item.get("name"), 'equipment_code': eq_code, 'location': room}
            
            for s_num, photos in eq_item.get("subjects", {}).items():
                for p_idx in range(1, 51):
                    tag_u, tag_h = f"S{s_num}_P{p_idx}", f"S{s_num}-P{p_idx}"
                    if p_idx <= len(photos):
                        p_data = photos[p_idx - 1]
                        # โหลดรูปภาพ
                        img_b = base64.b64decode(p_data.split(",")[-1]) if "base64" in p_data else requests.get(p_data).content
                        buf, w_in, h_in = fit_image_to_strict_bounds(img_b, target_width_in=1.9, target_height_in=2.0)
                        img_obj = InlineImage(tpl, buf, width=Inches(w_in), height=Inches(h_in))
                        eq_dict[tag_u] = img_obj
                        eq_dict[tag_h] = img_obj
                        context[tag_u] = img_obj
                        context[tag_h] = img_obj
                    else:
                        eq_dict[tag_u] = ""
                        eq_dict[tag_h] = ""
                        context[tag_u] = ""
                        context[tag_h] = ""
            context['equipment_list'].append(eq_dict)

        tpl.render(context)
        out_stream = BytesIO()
        tpl.save(out_stream)
        out_stream.seek(0)
        
        filename = f"SAT_Report_{topic}_{phase}_{room}.docx"
        return StreamingResponse(
            out_stream,
            media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            headers={"Content-Disposition": f"attachment; filename={filename}"}
        )
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))