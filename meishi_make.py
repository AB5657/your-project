import pandas as pd
from reportlab.pdfgen import canvas
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.utils import ImageReader
import os 
 







# --- 設定項目 ---
EXCEL_FILE = '社員一覧.xlsx'
SHEET_NAME = '社員一覧'
OUTPUT_PDF = '名刺一覧.pdf'
LOGO_FILE = '帳票_正和ロゴ.png'
FONT_PATH = "C:/Windows/Fonts/msgothic.ttc"
FONT_NAME = "MS-Gothic"

if os.path.exists(FONT_PATH):
    pdfmetrics.registerFont(TTFont(FONT_NAME, FONT_PATH))
else:
    FONT_NAME = "HeiseiKakuGo-W5" 

def draw_tight_label_value(c, label_base, value, x, y, font_name, size):
    c.setFont(font_name, size)
    c.drawString(x, y, label_base)
    curr_x = x + c.stringWidth(label_base, font_name, size) - 0.4*mm
    c.drawString(curr_x, y, "・")
    curr_x += c.stringWidth("・", font_name, size) - 0.8*mm
    c.drawString(curr_x, y, value)
    return curr_x + c.stringWidth(value, font_name, size)

def draw_tel_fax_line_final(c, tel_val, fax_val, x, y, width, font_name, size):
    if tel_val:
        draw_tight_label_value(c, "TEL", tel_val, x, y, font_name, size)
    if fax_val:
        fax_w = (c.stringWidth("FAX", font_name, size) + 
                 c.stringWidth("・", font_name, size) + 
                 c.stringWidth(fax_val, font_name, size) - 1.2*mm)
        start_x = x + width - fax_w
        draw_tight_label_value(c, "FAX", fax_val, start_x, y, font_name, size)

def create_business_cards():
    df = pd.read_excel(EXCEL_FILE, sheet_name=SHEET_NAME)
    c = canvas.Canvas(OUTPUT_PDF, pagesize=(55*mm, 91*mm))
    card_w, card_h = 55*mm, 91*mm

    logo = None
    if os.path.exists(LOGO_FILE):
        img = ImageReader(LOGO_FILE)
        img_w, img_h = img.getSize()
        aspect = img_h / float(img_w)
        logo = (img, 40*mm, card_h - 4*mm - (12*mm * aspect), 12*mm, 12*mm * aspect)

    for _, row in df.iterrows():
        data = row.fillna('').to_dict()
        if logo:
            c.drawImage(logo[0], logo[1], logo[2], width=logo[3], height=logo[4], mask='auto')

        current_y = 58 * mm
        c.setFont(FONT_NAME, 8) 
        for key in ['部署名１', '部署名２', '役職１']:
            val = str(data.get(key, '')).strip()
            if val:
                c.drawString(10*mm, current_y, val); current_y -= 4 * mm

        name = str(data.get('氏名', '')).strip()
        if name:
            current_y -= 5 * mm 
            c.setFont(FONT_NAME, 18); c.drawCentredString(card_w/2, current_y, name)

        margin = 3 * mm
        draw_w = card_w - (margin * 2)
        contact_y = 5 * mm
        line_s = 3.5 * mm
        BASE_SIZE = 8 

        # 1. Email/URL/携帯
        for label_desc, key in reversed([("E-mail: ", 'e-mail'), ("URL: ", 'URL'), ("携帯", '携帯')]):
            val = str(data.get(key, '')).strip()
            if val:
                if key == '携帯':
                    draw_tight_label_value(c, "携帯", val, margin, contact_y, FONT_NAME, BASE_SIZE)
                else:
                    text = f"{label_desc}{val}"
                    tw = c.stringWidth(text, FONT_NAME, BASE_SIZE)
                    fs = BASE_SIZE * (draw_w/tw) if tw > draw_w else BASE_SIZE
                    c.setFont(FONT_NAME, fs); c.drawString(margin, contact_y, text)
                contact_y += line_s

        # 2. TEL2/FAX2
        t2, f2 = str(data.get('TEL２', '')).strip(), str(data.get('FAX２', '')).strip()
        if t2 or f2:
            draw_tel_fax_line_final(c, t2, f2, margin, contact_y, draw_w, FONT_NAME, BASE_SIZE)
            contact_y += line_s

        # 3. 住所2-1
        p2, a2_1 = str(data.get('郵便番号２', '')).strip(), str(data.get('住所２-１', '')).strip()
        if p2 or a2_1:
            c.setFont(FONT_NAME, BASE_SIZE)
            zip_t2 = f"〒{p2} " if p2 else ""
            zw2 = c.stringWidth(zip_t2, FONT_NAME, BASE_SIZE)
            c.drawString(margin, contact_y, zip_t2)
            addr_w = c.stringWidth(a2_1, FONT_NAME, BASE_SIZE)
            remain_w = draw_w - zw2
            # TextObject内でのみスケールを適用
            to = c.beginText(margin + zw2, contact_y)
            to.setFont(FONT_NAME, BASE_SIZE)
            if addr_w > remain_w:
                to.setHorizScale((remain_w / addr_w) * 100 * 0.99)
            to.textLine(a2_1)
            c.drawText(to) # ここでTextObjectの設定は破棄される
            contact_y += line_s

        # 4. TEL1/FAX1
        t1, f1 = str(data.get('TEL１', '')).strip(), str(data.get('FAX１', '')).strip()
        if t1 or f1:
            draw_tel_fax_line_final(c, t1, f1, margin, contact_y, draw_w, FONT_NAME, BASE_SIZE)
            contact_y += line_s

        # --- 住所1-1 と 1-2 のサイズ連動 ---
        p1, a1_1 = str(data.get('郵便番号１', '')).strip(), str(data.get('住所１-１', '')).strip()
        a1_2 = str(data.get('住所１-２', '')).strip()
        c.setFont(FONT_NAME, BASE_SIZE)
        zip_t1 = f"〒{p1} " if p1 else ""
        zw1 = c.stringWidth(zip_t1, FONT_NAME, BASE_SIZE)
        remain_w1 = draw_w - zw1
        addr_w1 = c.stringWidth(a1_1, FONT_NAME, BASE_SIZE)
        target_scale = (remain_w1 / addr_w1 * 100 * 0.99) if addr_w1 > remain_w1 else 100.0

        # 5. 住所1-2
        if a1_2:
            to = c.beginText(margin, contact_y)
            to.setFont(FONT_NAME, BASE_SIZE)
            to.setHorizScale(target_scale)
            scaled_w = c.stringWidth(a1_2, FONT_NAME, BASE_SIZE) * (target_scale / 100.0)
            to.setTextOrigin(margin + draw_w - scaled_w, contact_y)
            to.textLine(a1_2)
            c.drawText(to)
            contact_y += line_s

        # 6. 住所1-1
        if p1 or a1_1:
            c.setFont(FONT_NAME, BASE_SIZE)
            c.drawString(margin, contact_y, zip_t1)
            to = c.beginText(margin + zw1, contact_y)
            to.setFont(FONT_NAME, BASE_SIZE)
            if addr_w1 > remain_w1:
                to.setHorizScale(target_scale)
            else:
                char_space = (remain_w1 - addr_w1) / (len(a1_1) - 1) if len(a1_1) > 1 else 0
                if char_space > 3*mm: char_space = 0
                to.setCharSpace(char_space)
            to.textLine(a1_1)
            c.drawText(to)
            contact_y += line_s

        # --- 会社名（ここからは通常のキャンバス描画に戻る） ---
        contact_y += 2.0 * mm
        extra = str(data.get('追記事業所', '')).strip()
        if extra:
            # キャンバス（c）の直接描画。TextObjectのスケールの影響を受けない。
            c.setFont(FONT_NAME, 9)
            c.drawCentredString(card_w/2, contact_y, extra)
            contact_y += 4.5 * mm
        
        c.setFont(FONT_NAME, 9); c.drawString(12*mm, contact_y, "株式会社")
        c.setFont(FONT_NAME, 13); c.drawString(28*mm, contact_y, "正　和")

        c.showPage()
    c.save()
    print(f"PDF作成完了: {OUTPUT_PDF}")

if __name__ == "__main__":
    create_business_cards()