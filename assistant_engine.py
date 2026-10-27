import os
import json
from typing import List, Dict, Any, Tuple
from openai import OpenAI
import pandas as pd

DEFAULT_BASE_URL = "https://opencode.ai/zen/go/v1"
DEFAULT_MODEL = "deepseek-v4.1-flash"

ASSISTANT_SYSTEM_PROMPT = """คุณคือ "SPACEMAN Assistant" ผู้ช่วยอัจฉริยะประจำระบบ BOQ & Project Schedule ของ SPACEMAN STUDIO
บทบาทของคุณคือคอยช่วยเหลือนักออกแบบ สถาปนิก ผู้รับเหมา และทีมงานในการ:
1. บริหารจัดการข้อมูล BOQ (ค้นหา, สรุปยอดเงิน, แก้ไขราคา, ปรับปริมาณงาน, เพิ่มลดรายการ)
2. ค้นหารหัสแบบก่อสร้าง, ตำแหน่งหน้าในแบบแปลน PDF, สเปกวัสดุ
3. วางแผนไทม์ไลน์โครงการ (Gantt Chart / Project Schedule) ส่งมอบลูกค้า
4. ให้คำแนะนำเรื่องสเปกวัสดุ (เช่น ไม้ HMR, E0, ลามิเนต, เมลามีน, ท็อปหิน) และราคากลางงานตกแต่งภายใน

กติกาสำคัญ:
- สื่อสารด้วยภาษาไทยอย่างสุภาพ กระชับ เป็นมืออาชีพ สไตล์ทีมงาน SPACEMAN STUDIO
- เมื่อผู้ใช้สั่งให้กระทำการใดๆ เกี่ยวกับระบบ (เช่น สรุปเงิน, แก้ราคา, เพิ่มรายการ, วางไทม์ไลน์) ให้เรียกใช้ Function / Tool ที่เหมาะสมเสมอ
- หลังจากเรียก Tool แล้ว ให้อธิบายผลลัพธ์ที่เกิดขึ้นให้ผู้ใช้ทราบอย่างชัดเจน พร้อมบอกยอดรวมหรือการเปลี่ยนแปลง
"""

ASSISTANT_TOOLS = [
    {
        "type": "function",
        "function": {
            "name": "get_boq_summary",
            "description": "ดูสรุปยอดเงินงบประมาณ BOQ ปัจจุบัน (ค่าวัสดุรวม, ค่าแรงรวม, ค่าดำเนินการ, ยอดรวมทั้งสิ้น และจำนวนรายการ)",
            "parameters": {
                "type": "object",
                "properties": {},
                "required": []
            }
        }
    },
    {
        "type": "function",
        "function": {
            "name": "update_or_add_boq_item",
            "description": "แก้ไข ปรับราคา เปลี่ยนจำนวน หรือเพิ่ม/ลบรายการในตาราง BOQ",
            "parameters": {
                "type": "object",
                "properties": {
                    "code": {
                        "type": "string", 
                        "description": "รหัสแบบ เช่น J1-01, J2-01, D1-01 หรือชื่อย่อรายการ เช่น 'เคาเตอร์ห้องน้ำ', 'กระเบื้อง'"
                    },
                    "description": {
                        "type": "string",
                        "description": "คำอธิบายรายการงานและสเปกวัสดุ"
                    },
                    "category": {
                        "type": "string",
                        "description": "หมวดหมู่งาน เช่น ชั้น 1, ชั้น 2, งานระบบ, Preliminary"
                    },
                    "qty": {
                        "type": "number",
                        "description": "จำนวนหรือปริมาณงานใหม่"
                    },
                    "unit": {
                        "type": "string",
                        "description": "หน่วยนับ เช่น SET, ตร.ม., จุด, LOT"
                    },
                    "material_rate": {
                        "type": "number",
                        "description": "ราคาค่าวัสดุต่อหน่วย (บาท)"
                    },
                    "labour_rate": {
                        "type": "number",
                        "description": "ราคาค่าแรงต่อหน่วย (บาท)"
                    },
                    "action": {
                        "type": "string",
                        "enum": ["update", "add", "delete"],
                        "description": "update = แก้ไขข้อมูลเดิม, add = เพิ่มรายการใหม่, delete = ลบรายการ"
                    }
                },
                "required": ["code"]
            }
        }
    },
    {
        "type": "function",
        "function": {
            "name": "search_drawings",
            "description": "ค้นหาข้อมูลรหัสแบบแปลน หรือหน้าแบบที่เกี่ยวข้องจากไฟล์แบบก่อสร้าง PDF",
            "parameters": {
                "type": "object",
                "properties": {
                    "query": {
                        "type": "string",
                        "description": "คำค้นหา เช่น J1-01, ตู้เสื้อผ้า, ประตู, ห้องน้ำ, หรือเลขหน้า"
                    }
                },
                "required": ["query"]
            }
        }
    },
    {
        "type": "function",
        "function": {
            "name": "create_or_update_schedule",
            "description": "สร้างหรืออัปเดตไทม์ไลน์โครงการ (Project Schedule Gantt Chart)",
            "parameters": {
                "type": "object",
                "properties": {
                    "prompt": {
                        "type": "string",
                        "description": "รายละเอียดโครงการ วันที่เริ่มต้น ขั้นตอน ระยะเวลา และกำหนดเปิดร้าน"
                    }
                },
                "required": ["prompt"]
            }
        }
    },
    {
        "type": "function",
        "function": {
            "name": "load_demo_dataset",
            "description": "โหลดข้อมูลตัวอย่างโครงการ (PLC House หรือ Brewave Aree) เข้าสู่ระบบ",
            "parameters": {
                "type": "object",
                "properties": {
                    "dataset_name": {
                        "type": "string",
                        "enum": ["plc_house_boq", "brewave_schedule", "both"],
                        "description": "plc_house_boq = ตัวอย่าง BOQ, brewave_schedule = ตัวอย่างไทม์ไลน์, both = ทั้งสองอย่าง"
                    }
                },
                "required": ["dataset_name"]
            }
        }
    }
]

def execute_assistant_tool(
    tool_name: str,
    tool_args: Dict[str, Any],
    session_data: Dict[str, Any]
) -> Tuple[Dict[str, Any], str]:
    """
    Executes a tool call locally modifying session_data where appropriate.
    Returns (result_dict, human_readable_summary).
    """
    items_data = session_data.get("items_data", [])
    overhead_pct = session_data.get("overhead_pct", 10.0)
    
    if tool_name == "get_boq_summary":
        if not items_data:
            return {"status": "empty", "message": "ยังไม่มีรายการในตาราง BOQ"}, "ปัจจุบันยังไม่มีข้อมูลรายการ BOQ ในระบบ"
            
        total_mat = sum(float(it.get("material_rate", 0)) * float(it.get("qty", 1)) for it in items_data)
        total_lab = sum(float(it.get("labour_rate", 0)) * float(it.get("qty", 1)) for it in items_data)
        subtotal = total_mat + total_lab
        overhead = subtotal * (overhead_pct / 100.0)
        grand_total = subtotal + overhead

        summary_res = {
            "item_count": len(items_data),
            "total_material": total_mat,
            "total_labour": total_lab,
            "subtotal": subtotal,
            "overhead_pct": overhead_pct,
            "overhead_amount": overhead,
            "grand_total": grand_total
        }
        summary_text = (
            f"ยอดรวม BOQ ปัจจุบัน ({len(items_data)} รายการ):\n"
            f"- ค่าวัสดุรวม: {total_mat:,.2f} บาท\n"
            f"- ค่าแรงรวม: {total_lab:,.2f} บาท\n"
            f"- ค่าดำเนินการ ({overhead_pct}%): {overhead:,.2f} บาท\n"
            f"- ราคารวมทั้งสิ้น (Grand Total): {grand_total:,.2f} บาท"
        )
        return summary_res, summary_text

    elif tool_name == "update_or_add_boq_item":
        code_query = str(tool_args.get("code", "")).strip().upper()
        action = tool_args.get("action", "update")
        
        # Search for existing item
        matched_idx = -1
        for idx, it in enumerate(items_data):
            c = str(it.get("code", "")).strip().upper()
            d = str(it.get("description", "")).strip().upper()
            if code_query and (code_query == c or code_query in d or c in code_query):
                matched_idx = idx
                break
                
        if action == "delete":
            if matched_idx != -1:
                deleted = items_data.pop(matched_idx)
                return {"status": "deleted", "item": deleted}, f"ลบรายการ {deleted.get('code', '')} ออกจากตารางเรียบร้อยแล้ว"
            return {"status": "not_found"}, f"ไม่พบรายการรหัส {code_query} สำหรับลบ"
            
        elif action == "update" and matched_idx != -1:
            target = items_data[matched_idx]
            if "qty" in tool_args and tool_args["qty"] is not None:
                target["qty"] = float(tool_args["qty"])
            if "unit" in tool_args and tool_args["unit"]:
                target["unit"] = str(tool_args["unit"])
            if "material_rate" in tool_args and tool_args["material_rate"] is not None:
                target["material_rate"] = float(tool_args["material_rate"])
            if "labour_rate" in tool_args and tool_args["labour_rate"] is not None:
                target["labour_rate"] = float(tool_args["labour_rate"])
            if "description" in tool_args and tool_args["description"]:
                target["description"] = str(tool_args["description"])
            if "category" in tool_args and tool_args["category"]:
                target["category"] = str(tool_args["category"])
                
            return {"status": "updated", "item": target}, f"อัปเดตรายการ {target.get('code', '')} ({target.get('description', '')}) เรียบร้อยแล้ว"
            
        else: # add or item not found for update
            new_item = {
                "code": tool_args.get("code", "ITEM"),
                "category": tool_args.get("category", "งานเพิ่มเติม"),
                "description": tool_args.get("description", tool_args.get("code", "")),
                "qty": float(tool_args.get("qty", 1.0)),
                "unit": tool_args.get("unit", "SET"),
                "material_rate": float(tool_args.get("material_rate", 0.0)),
                "labour_rate": float(tool_args.get("labour_rate", 0.0)),
                "remark": "เพิ่มโดย AI Assistant"
            }
            items_data.append(new_item)
            return {"status": "added", "item": new_item}, f"เพิ่มรายการใหม่ {new_item['code']}: {new_item['description']} เรียบร้อยแล้ว"

    elif tool_name == "search_drawings":
        q = str(tool_args.get("query", "")).strip().upper()
        detected_codes = session_data.get("detected_codes", [])
        
        matches = []
        for c in detected_codes:
            code_str = str(c.get("code", "")).upper()
            ctx = str(c.get("context", "")).upper()
            if q in code_str or q in ctx:
                matches.append(c)
                
        if matches:
            return {"count": len(matches), "matches": matches[:10]}, f"ตรวจพบรหัสที่เกี่ยวข้อง {len(matches)} รายการ เช่น {[m['code'] for m in matches[:5]]}"
        else:
            return {"count": 0, "message": "ไม่พบรหัสแบบที่ตรงกับคำค้นหาโดยตรง"}, f"ไม่พบรหัสแบบที่ตรงกับคำว่า '{q}' ในสารบัญแบบที่ตรวจพบ"

    elif tool_name == "load_demo_dataset":
        dataset_name = tool_args.get("dataset_name", "both")
        import boq_engine
        import schedule_engine
        
        loaded = []
        if dataset_name in ["plc_house_boq", "both"]:
            session_data["items_data"] = boq_engine.get_demo_precalculated_items()
            loaded.append("BOQ PLC House (25 รายการ)")
        if dataset_name in ["brewave_schedule", "both"]:
            session_data["schedule_data"] = schedule_engine.get_sample_brewave_schedule()
            loaded.append("ไทม์ไลน์ Brewave Aree (7 สเตจ)")
            
        return {"status": "loaded", "datasets": loaded}, f"โหลดข้อมูลตัวอย่างสำเร็จ: {', '.join(loaded)}"

    elif tool_name == "create_or_update_schedule":
        prompt_txt = tool_args.get("prompt", "")
        import schedule_engine
        api_key = session_data.get("api_key", "")
        base_url = session_data.get("base_url", DEFAULT_BASE_URL)
        model_name = session_data.get("model_name", DEFAULT_MODEL)
        
        if api_key:
            try:
                new_sched = schedule_engine.parse_schedule_prompt(api_key, prompt_txt, base_url, model_name)
                session_data["schedule_data"] = new_sched
                return {"status": "success", "project": new_sched.get("project_name")}, f"สร้างไทม์ไลน์โครงการ {new_sched.get('project_name', '')} สำเร็จแล้ว!"
            except Exception as e:
                return {"status": "error", "error": str(e)}, f"ไม่สามารถสร้างไทม์ไลน์ได้: {str(e)}"
        else:
            session_data["schedule_data"] = schedule_engine.get_sample_brewave_schedule()
            return {"status": "demo_loaded"}, "โหลดไทม์ไลน์ตัวอย่างสำเร็จ (กรุณาตั้งค่า API Key เพื่อสร้างไทม์ไลน์แบบกำหนดเอง)"

    return {"status": "unknown_tool"}, "ไม่พบเครื่องมือที่ต้องการ"

def run_chat_turn(
    messages: List[Dict[str, str]],
    api_key: str,
    session_data: Dict[str, Any],
    base_url: str = DEFAULT_BASE_URL,
    model_name: str = DEFAULT_MODEL
) -> Tuple[str, List[str]]:
    """
    Executes a single conversational turn with tool calling enabled.
    Returns (assistant_reply_text, list_of_actions_taken).
    """
    if not api_key:
        return "กรุณาใส่ AI Engine API Key ในการตั้งค่าระบบ (⚙️ ด้านล่าง) ก่อนเริ่มสนทนาครับ", []

    client = OpenAI(base_url=base_url, api_key=api_key, default_headers={"x-opencode-session": "ses_spaceman_chat"})
    
    # Format messages for OpenAI API
    api_messages = [{"role": "system", "content": ASSISTANT_SYSTEM_PROMPT}]
    for msg in messages:
        api_messages.append({"role": msg["role"], "content": msg["content"]})
        
    try:
        response = client.chat.completions.create(
            model=model_name,
            messages=api_messages,
            tools=ASSISTANT_TOOLS,
            tool_choice="auto",
            temperature=0.3
        )
    except Exception as e:
        return f"เกิดข้อผิดพลาดในการเชื่อมต่อกับ AI: {str(e)}", []

    choice = response.choices[0]
    actions_taken = []
    
    # Check if AI decided to call one or more tools
    if choice.message.tool_calls:
        # Append assistant's tool-call request message
        api_messages.append(choice.message)
        
        for tool_call in choice.message.tool_calls:
            t_name = tool_call.function.name
            try:
                t_args = json.loads(tool_call.function.arguments or "{}")
            except:
                t_args = {}
                
            tool_result, human_summary = execute_assistant_tool(t_name, t_args, session_data)
            actions_taken.append(f"🛠️ {human_summary}")
            
            # Append tool response
            api_messages.append({
                "role": "tool",
                "tool_call_id": tool_call.id,
                "content": json.dumps(tool_result, ensure_ascii=False)
            })
            
        # Second call to get final natural language answer
        try:
            second_resp = client.chat.completions.create(
                model=model_name,
                messages=api_messages,
                temperature=0.3
            )
            final_reply = second_resp.choices[0].message.content or "ดำเนินการตามคำสั่งเรียบร้อยแล้วครับ!"
            return final_reply, actions_taken
        except Exception as e:
            return f"ดำเนินการสำเร็จแล้ว: {', '.join(actions_taken)}", actions_taken
            
    else:
        return choice.message.content or "มีอะไรให้ผมช่วยเรื่อง BOQ หรือไทม์ไลน์งานไหมครับ?", []
