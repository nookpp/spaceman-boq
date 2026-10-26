import json
import os
from typing import List, Dict, Any, Optional
from openai import OpenAI

DEFAULT_BASE_URL = "https://opencode.ai/zen/go/v1"
DEFAULT_MODEL = "deepseek-v4.1-flash"

SYSTEM_PROMPT = """คุณคือวิศวกรและผู้เชี่ยวชาญการประมาณราคา (Senior Quantity Surveyor / Interior Cost Estimator)
หน้าที่ของคุณคือ:
1. วิเคราะห์แบบก่อสร้าง, แปลนสถาปัตย์, แบบขยาย Built-in (Joinery / Interior) และสเปกงานที่ได้รับ
2. ถอดรายการปริมาณงาน (Bill of Quantities: BOQ) ออกมาเป็นโครงสร้างตารางที่ชัดเจน
3. ระบุรหัสแบบ (เช่น J1-01, J2-01, D1-01), รายการและสเปกวัสดุ (ขนาด กว้างxยาวxสูง, ชนิดไม้, ปิดผิวลามิเนต), ปริมาณ (Qty), หน่วย (Unit เช่น SET, ตร.ม., ม., จุด, LOT)
4. ประเมินราคาเบื้องต้นตามราคาตลาดงานตกแต่งภายใน (Material Rate และ Labour Rate ต่อหน่วย เป็นเงินบาท)

สำคัญมาก:
- ส่งผลลัพธ์กลับมาเป็น JSON array ภายใต้คีย์ "items" เท่านั้น
- รูปแบบ Schema:
{
  "items": [
    {
      "code": "รหัสแบบ เช่น J1-01",
      "category": "หมวดหมู่งาน เช่น ชั้น 1 หรือ ชั้น 2 หรือ งานระบบ",
      "description": "คำอธิบายรายการ พร้อมขนาดและสเปกวัสดุ",
      "qty": 1.0,
      "unit": "SET",
      "material_rate": 45000,
      "labour_rate": 8000,
      "remark": "หมายเหตุ เช่น ดูแบบขยาย ID3-01"
    }
  ]
}
อย่าพิมพ์ข้อความเกริ่นนำหรือคำลงท้าย ให้ส่งเฉพาะ JSON ล้วนๆ
"""

def extract_boq_from_text_or_drawing(
    api_key: str,
    content_text: str,
    base_url: str = DEFAULT_BASE_URL,
    model_name: str = DEFAULT_MODEL,
    system_prompt: Optional[str] = None
) -> List[Dict[str, Any]]:
    """
    Calls OpenCode Go (OpenAI-compatible API) to extract BOQ items.
    """
    if not api_key:
        raise ValueError("กรุณาระบุ OpenCode Go API Key")

    client = OpenAI(
        base_url=base_url,
        api_key=api_key
    )

    prompt = f"วิเคราะห์ข้อมูลแบบก่อสร้างและสเปกต่อไปนี้ แล้วถอดรายการ BOQ ตามฟอร์แมต JSON ที่กำหนด:\n\n{content_text}"

    response = client.chat.completions.create(
        model=model_name,
        messages=[
            {"role": "system", "content": system_prompt or SYSTEM_PROMPT},
            {"role": "user", "content": prompt}
        ],
        temperature=0.2,
    )

    raw_content = response.choices[0].message.content or "{}"
    
    # Clean potential markdown fences ```json ... ```
    cleaned = raw_content.strip()
    if cleaned.startswith("```"):
        cleaned = cleaned.split("\n", 1)[-1]
    if cleaned.endswith("```"):
        cleaned = cleaned.rsplit("```", 1)[0]
    cleaned = cleaned.strip()

    try:
        data = json.loads(cleaned)
        if isinstance(data, dict) and "items" in data:
            return data["items"]
        elif isinstance(data, list):
            return data
        else:
            return []
    except json.JSONDecodeError as e:
        # Fallback: attempt to find JSON array
        import re
        match = re.search(r'\[.*\]', cleaned, re.DOTALL)
        if match:
            return json.loads(match.group(0))
        raise ValueError(f"AI ตอบกลับมาในรูปแบบที่ไม่ใช่ JSON ถูกต้อง: {raw_content[:200]}")

def get_demo_precalculated_items() -> List[Dict[str, Any]]:
    """
    Returns realistic market-calibrated BOQ items matching PLC HOUSE project.
    Can be used instantly for testing without needing an active API Key.
    """
    return [
        # ชั้น 1
        {"code": "J1-01", "category": "ชั้น 1", "description": "ตู้ TV ไม้ HMR กรุลามิเนตลายไม้และสีขาว ขนาด 1.14*2.62*2.70 ม.", "qty": 1.0, "unit": "SET", "material_rate": 42000, "labour_rate": 8000, "remark": "ID3-01"},
        {"code": "J1-02", "category": "ชั้น 1", "description": "ตู้สูง ไม้ HMR กรุลามิเนตลายไม้และสีขาว ขนาด 0.45*0.37/3.42*2.70 ม.", "qty": 1.0, "unit": "SET", "material_rate": 38000, "labour_rate": 7500, "remark": "ID3-05"},
        {"code": "J1-03", "category": "ชั้น 1", "description": "ตู้ลอย ไม้ HMR กรุลามิเนตลายไม้และสีขาว ขนาด 0.45*0.37/3.42*2.70 ม.", "qty": 1.0, "unit": "SET", "material_rate": 28000, "labour_rate": 6000, "remark": "ID3-04"},
        {"code": "J1-04", "category": "ชั้น 1", "description": "เคาเตอร์ ขนาด 0.60*2.25*2.70 ม.", "qty": 1.0, "unit": "SET", "material_rate": 32000, "labour_rate": 6500, "remark": "ID3-01"},
        {"code": "J1-05", "category": "ชั้น 1", "description": "เคาเตอร์ชุดครัว ขนาด 0.60*7.00*2.70 ม.", "qty": 1.0, "unit": "SET", "material_rate": 115000, "labour_rate": 25000, "remark": "ท็อปหินสังเคราะห์"},
        {"code": "J1-06", "category": "ชั้น 1", "description": "ตู้เตี้ย+ผนังกระจกเงา 0.60*1.10*2.70 ม.", "qty": 1.0, "unit": "SET", "material_rate": 19000, "labour_rate": 4500, "remark": "ID3-07"},
        {"code": "J1-07", "category": "ชั้น 1", "description": "ตู้สูง 0.60*3.00*2.70", "qty": 1.0, "unit": "SET", "material_rate": 48000, "labour_rate": 9500, "remark": "-"},
        {"code": "D1-01", "category": "ชั้น 1", "description": "ประตูส่วนต่อเติม ขนาด 0.80*2.70 ม.", "qty": 1.0, "unit": "SET", "material_rate": 12000, "labour_rate": 2500, "remark": "-"},
        {"code": "D1-03", "category": "ชั้น 1", "description": "ประตูส่วนต่อเติม ขนาด 0.80*2.70 ม.", "qty": 1.0, "unit": "SET", "material_rate": 12000, "labour_rate": 2500, "remark": "-"},
        {"code": "ห้องน้ำ", "category": "ชั้น 1", "description": "เคาเตอร์ห้องน้ำชั้นล่าง 0.60*2.10*0.538", "qty": 1.0, "unit": "SET", "material_rate": 22000, "labour_rate": 4500, "remark": "-"},
        {"code": "หน้าบ้าน", "category": "ชั้น 1", "description": "งานปูกระเบื้องพื้น", "qty": 8.0, "unit": "ตร.ม.", "material_rate": 650, "labour_rate": 350, "remark": "-"},
        
        # ชั้น 2
        {"code": "J2-01", "category": "ชั้น 2", "description": "ผนังตกแต่งหัวเตียง ขนาด 0.20*3.40*2.80 ม.", "qty": 1.0, "unit": "SET", "material_rate": 36000, "labour_rate": 8000, "remark": "ID3-01"},
        {"code": "J2-02", "category": "ชั้น 2", "description": "โต๊ะทำงานห้องนอนใหญ่+ตู้ลอย ขนาด 0.20*2.40*2.80 ม.", "qty": 1.0, "unit": "SET", "material_rate": 29000, "labour_rate": 6000, "remark": "ID3-02"},
        {"code": "J2-03", "category": "ชั้น 2", "description": "ตู้เสื้อผ้า ห้องแต่งตัว ขนาด 1.30*1.30*2.80 ม.", "qty": 1.0, "unit": "SET", "material_rate": 45000, "labour_rate": 9000, "remark": "ID3-01"},
        {"code": "J2-04", "category": "ชั้น 2", "description": "ตู้เสื้อผ้า ห้องแต่งตัว ขนาด 0.60*2.10*2.80 ม.", "qty": 1.0, "unit": "SET", "material_rate": 38000, "labour_rate": 7500, "remark": "ID3-03"},
        {"code": "J2-05", "category": "ชั้น 2", "description": "ตู้เสื้อผ้า ห้องแต่งตัว ขนาด 0.60*1.80/0.68*2.80 ม.", "qty": 1.0, "unit": "SET", "material_rate": 35000, "labour_rate": 7000, "remark": "ID3-05"},
        {"code": "J2-06", "category": "ชั้น 2", "description": "โต๊ะแต่งตัว ขนาด 0.60*4.2*2.80/0.80 ม.", "qty": 1.0, "unit": "SET", "material_rate": 32000, "labour_rate": 6500, "remark": "ID3-07"},
        {"code": "J2-07", "category": "ชั้น 2", "description": "ตู้สูงใส่รองเท้า ขนาด 0.65*1.50*1.06/1.74 ม.", "qty": 1.0, "unit": "SET", "material_rate": 26000, "labour_rate": 5500, "remark": "ID3-10"},
        {"code": "J2-13", "category": "ชั้น 2", "description": "ตู้เสื้อผ้า ขนาด 0.60*1.40*2.80 ม.", "qty": 1.0, "unit": "SET", "material_rate": 28000, "labour_rate": 5500, "remark": "ID3-01"},

        # งานระบบ
        {"code": "ติดตั้งดวงโคม (เฉพาะค่าแรงไม่รวมของ)", "category": "งานระบบ", "description": "ติดตั้งดวงโคม (เฉพาะค่าแรงไม่รวมของ)", "qty": 1.0, "unit": "LOT", "material_rate": 0, "labour_rate": 15000, "remark": "รวมทั้งหลัง"},
        {"code": "ติดตั้งพัดลมดูดอากาศ (เฉพาะค่าแรงไม่รวมของ)", "category": "งานระบบ", "description": "ติดตั้งพัดลมดูดอากาศ (เฉพาะค่าแรงไม่รวมของ)", "qty": 1.0, "unit": "LOT", "material_rate": 0, "labour_rate": 3500, "remark": "-"},
        {"code": "ติดตั้งปลั๊ก (เฉพาะค่าแรงไม่รวมของ)", "category": "งานระบบ", "description": "ติดตั้งปลั๊ก (เฉพาะค่าแรงไม่รวมของ)", "qty": 1.0, "unit": "LOT", "material_rate": 0, "labour_rate": 8000, "remark": "-"},
        {"code": "ติดตั้งสวิตซ์ (เฉพาะค่าแรงไม่รวมของ)", "category": "งานระบบ", "description": "ติดตั้งสวิตซ์ (เฉพาะค่าแรงไม่รวมของ)", "qty": 1.0, "unit": "LOT", "material_rate": 0, "labour_rate": 5000, "remark": "-"},

        # Preliminary
        {"code": "งานโพรเทคชั่น", "category": "Preliminary", "description": "งานโพรเทคชั่นป้องกันพื้นที่", "qty": 1.0, "unit": "LOT", "material_rate": 15000, "labour_rate": 8000, "remark": "-"},
        {"code": "งานทำความสะอาด", "category": "Preliminary", "description": "งานทำความสะอาด Big Cleaning", "qty": 1.0, "unit": "LOT", "material_rate": 5000, "labour_rate": 7000, "remark": "-"},
        {"code": "ค่าเดินทาง", "category": "Preliminary", "description": "ค่าเดินทางและควบคุมงาน", "qty": 1.0, "unit": "LOT", "material_rate": 10000, "labour_rate": 0, "remark": "-"}
    ]
