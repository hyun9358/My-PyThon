import tkinter as tk
from tkinter import colorchooser, filedialog, ttk
from PIL import Image, ImageDraw, ImageFont, ImageTk
import os
import sys
import platform
import subprocess
import math

# --- 다크모드 감지 함수들 ---
def is_windows_dark_mode():
    try:
        import winreg
        registry = winreg.ConnectRegistry(None, winreg.HKEY_CURRENT_USER)
        key_path = r'SOFTWARE\Microsoft\Windows\CurrentVersion\Themes\Personalize'
        key = winreg.OpenKey(registry, key_path)
        value, _ = winreg.QueryValueEx(key, "AppsUseLightTheme")
        winreg.CloseKey(key)
        return value == 0
    except Exception:
        return False

def is_macos_dark_mode():
    try:
        result = subprocess.run(
            ['defaults', 'read', '-g', 'AppleInterfaceStyle'],
            capture_output=True, text=True)
        return 'Dark' in result.stdout
    except Exception:
        return False

def detect_dark_mode():
    if sys.platform.startswith('win'):
        return is_windows_dark_mode()
    elif sys.platform == 'darwin':
        return is_macos_dark_mode()
    else:
        return False

# --- 리소스 경로 처리 ---
if getattr(sys, 'frozen', False):
    base_path = sys._MEIPASS
else:
    base_path = os.path.dirname(__file__)

ICON_PATH = os.path.join(base_path, "Create Stamp.ico")

# --- 사용자 폰트 목록 및 경로 리스트 ---
fonts = [
    (os.path.join(base_path, "HJ한전서A.ttf"), "HJ한전서A"),
    (os.path.join(base_path, "HJ한전서B.ttf"), "HJ한전서B"),
    (os.path.join(base_path, "malgun.ttf"), "맑은 고딕 (일반)"),
    (os.path.join(base_path, "malgunbd.ttf"), "맑은 고딕 (볼드)"),
    (os.path.join(base_path, "NanumMyeongjo.ttf"), "나눔명조 (일반)"),
    (os.path.join(base_path, "NanumMyeongjoBold.ttf"), "나눔명조 (볼드)"),
    (os.path.join(base_path, "NanumMyeongjoExtraBold.ttf"), "나눔명조 (엑스트라볼드)"),
]

# --- 상수들 ---
STAMP_TYPES = [
    "원형", "정사각형", "엣지 있는 정사각형",
    "직사각형", "엣지 있는 직사각형", "문자 to 도장모양"
]

SUB_STAMP_SHAPES = [
    "원", "정사각", "직사각(가로)", "직사각(세로)", "타원형(가로)", "타원형(세로)"
]

SUB_STAMP_SHAPES_CIRCLE = ["원", "타원형(가로)", "타원형(세로)"]
SUB_STAMP_SHAPES_RECT = ["정사각", "직사각(가로)", "직사각(세로)"]

TEXT_DIRECTIONS = [
    "좌->우", "상->하", "조선체 스타일"
]

# --- 기본 변수들 ---
IMG_SIZE = 400
BORDER_WIDTH = 5
fg_color_ok = "#008800"
fg_color_error = "#CC0000"

# 폰트 크기 1부터 150까지 생성, 60,80,100,120은 (추천) 붙임
font_sizes = []
for i in range(1, 151):
    if i in (60, 80, 100, 120):
        font_sizes.append(f"{i} (추천)")
    else:
        font_sizes.append(str(i))

# --- 전역 상태: 사용자가 선택한 도장 크기와 줄 수, 줄간격 ---
stamp_width = IMG_SIZE
stamp_height = IMG_SIZE
selected_line_count = 1
selected_line_spacing = 1.0  # 줄간격 배율: 0.9(좁게),1.0(보통),1.2(넉넉)

# --- 도장 생성 함수 ---
def generate_seal():
    global preview_img, stamp_width, stamp_height, selected_line_count, selected_line_spacing
    try:
        name = name_entry.get().strip()
        if not name:
            status_label.config(text="이름을 입력하세요.", fg=fg_color_error)
            return

        # 폰트 크기 처리 - 콤보에서 (추천) 제거 후 정수 변환
        font_size_str = font_size_combo.get()
        font_size = int(font_size_str.split()[0])  # "60 (추천)" -> 60

        text_dir = text_direction_combo.get()
        stamp_type = stamp_type_combo.get()
        selected_font_index = font_name_combo.current()
        selected_font_path = fonts[selected_font_index][0]

        # 사용자 선택 도장 사이즈 (콤보에서 반영)
        try:
            size_text = size_combo.get()
            if size_text and size_text != "사용자 지정":
                w, h = map(int, size_text.split('x'))
                stamp_width = w
                stamp_height = h
            else:
                # 사용자 지정이면 엔트리값 읽기
                cw = custom_width_entry.get().strip()
                ch = custom_height_entry.get().strip()
                if cw and ch:
                    stamp_width = max(50, int(cw))
                    stamp_height = max(50, int(ch))
                else:
                    # 기본값 유지
                    stamp_width = IMG_SIZE
                    stamp_height = IMG_SIZE
        except Exception:
            # 실패하면 기본 값 사용
            stamp_width = IMG_SIZE
            stamp_height = IMG_SIZE

        # 줄 수 콤보에서 읽기 (문자 to 도장모양에서만 사용됨)
        try:
            selected_line_count = int(line_count_combo.get())
        except Exception:
            selected_line_count = 1

        # 줄간격 콤보
        try:
            spacing_choice = line_spacing_combo.get()
            if spacing_choice == "좁게":
                selected_line_spacing = 0.9
            elif spacing_choice == "보통":
                selected_line_spacing = 1.0
            else:
                selected_line_spacing = 1.2
        except Exception:
            selected_line_spacing = 1.0

        try:
            # font 객체는 기본값으로 생성하되, draw_text_to_stamp_shape 내부에서 필요하면 자동 조정됩니다.
            font = ImageFont.truetype(selected_font_path, font_size)
        except Exception as e:
            status_label.config(text=f"폰트 로딩 실패: {e}", fg=fg_color_error)
            font = ImageFont.load_default()

        # 문자 to 도장모양 서브메뉴 선택 확인
        if stamp_type == "문자 to 도장모양":
            sub_shape = sub_stamp_shape_combo.get()
            if not sub_shape:
                status_label.config(text="서브 모양을 선택하세요.", fg=fg_color_error)
                return
            # 문자 to 도장모양은 도장 외곽 없음, 글자만 도장형태에 맞게 배치
            width, height = stamp_width, stamp_height
            img = Image.new("RGBA", (width, height), (0, 0, 0, 0))
            draw = ImageDraw.Draw(img)

            color = seal_color.get()
            # 문자 to 도장모양 함수 호출 (줄 수, 줄간격, 폰트 경로/크기 전달)
            draw_text_to_stamp_shape(draw, name, selected_font_path, font.size, width, height, color, sub_shape, selected_line_count, selected_line_spacing)

        else:
            # 기존 도장모양 처리 (외곽 그리기)
            if stamp_type in ["직사각형", "엣지 있는 직사각형"]:
                width = stamp_width
                height = int(stamp_width * 0.7)  # 기존 비율 유지
            else:
                width = stamp_width
                height = stamp_height

            img = Image.new("RGBA", (width, height), (0, 0, 0, 0))
            draw = ImageDraw.Draw(img)

            color = seal_color.get()

            stamp_type = stamp_type_combo.get()
            sub_shape = sub_stamp_shape_combo.get()

            if stamp_type == "원형":
                if sub_shape == "원":
                    draw.ellipse(
                    (BORDER_WIDTH//2, BORDER_WIDTH//2, width-BORDER_WIDTH//2, height-BORDER_WIDTH//2),
                    outline=color, width=BORDER_WIDTH)
                elif sub_shape == "타원형(가로)":
                    draw.ellipse(
                        (BORDER_WIDTH//2, height//4, width-BORDER_WIDTH//2, height*3//4),
                        outline=color, width=BORDER_WIDTH)
                elif sub_shape == "타원형(세로)":
                    draw.ellipse(
                        (width//4, BORDER_WIDTH//2, width*3//4, height-BORDER_WIDTH//2),
                        outline=color, width=BORDER_WIDTH)

            elif stamp_type == "정사각형":
                draw.rectangle(
                    (BORDER_WIDTH//2, BORDER_WIDTH//2, width-BORDER_WIDTH//2, height-BORDER_WIDTH//2),
                    outline=color, width=BORDER_WIDTH)
            elif stamp_type == "엣지 있는 정사각형":
                radius = 30
                draw_rounded_rectangle(draw, (BORDER_WIDTH//2, BORDER_WIDTH//2, width-BORDER_WIDTH//2, height-BORDER_WIDTH//2),
                                       radius, outline=color, width=BORDER_WIDTH)
            elif stamp_type == "직사각형":
                draw.rectangle(
                    (BORDER_WIDTH//2, BORDER_WIDTH//2, width-BORDER_WIDTH//2, height-BORDER_WIDTH//2),
                    outline=color, width=BORDER_WIDTH)
            elif stamp_type == "엣지 있는 직사각형":
                radius = 20
                draw_rounded_rectangle(draw, (BORDER_WIDTH//2, BORDER_WIDTH//2, width-BORDER_WIDTH//2, height-BORDER_WIDTH//2),
                                       radius, outline=color, width=BORDER_WIDTH)

            # 글자 방향에 따라 텍스트 배치 (기존 방식 유지)
            if text_direction_combo.get() == "좌->우":
                draw_text_left_to_right(draw, name, ImageFont.truetype(selected_font_path, font.size), width, height, color)
            elif text_direction_combo.get() == "상->하":
                draw_text_top_to_bottom(draw, name, ImageFont.truetype(selected_font_path, font.size), width, height, color)
            else:
                draw_text_joseon_style(draw, name, ImageFont.truetype(selected_font_path, font.size), width, height, color)

        preview_img = ImageTk.PhotoImage(img)
        preview_label.config(image=preview_img)
        preview_label.image = preview_img
        preview_label.image_obj = img  # 알파 채널 유지용
        status_label.config(text="도장이 생성되었습니다.", fg=fg_color_ok)

    except Exception as e:
        status_label.config(text=f"도장 생성 실패: {e}", fg=fg_color_error)
        print(f"도장 생성 실패: {e}")

# --- 문자 to 도장모양 글자 배치 함수 ---
def draw_text_to_stamp_shape(draw, text, font_path, base_font_size, width, height, color, shape, line_count=1, line_spacing=1.0):
    """
    draw: ImageDraw.Draw 인스턴스
    text: 출력할 문자열
    font_path: truetype 폰트 파일 경로 (문자 크기 자동 조정을 위해 필요)
    base_font_size: 사용자가 선택한 기본 폰트 크기 (정수)
    width, height: 이미지 크기
    color: 문자열 색상
    shape: "원", "정사각", "직사각(가로)" 등
    line_count: 1..6
    line_spacing: 줄간격 배율 (예: 0.9, 1.0, 1.2)
    """

    n = len(text)
    BORDER_WIDTH = 20

    if n == 0:
        return

    # 안전한 줄 수 범위
    try:
        line_count = int(line_count)
    except Exception:
        line_count = 1
    line_count = max(1, min(6, line_count))

    # --- 초기 폰트 생성 (base) 및 문자 크기 측정 ---
    try:
        font = ImageFont.truetype(font_path, base_font_size)
    except Exception:
        font = ImageFont.load_default()
    # 측정(현재 폰트)
    max_w, max_h = 0, 0
    for ch in text:
        try:
            bbox = draw.textbbox((0,0), ch, font=font)
            w = bbox[2] - bbox[0]
            h = bbox[3] - bbox[1]
        except AttributeError:
            w, h = draw.textsize(ch, font=font)
        max_w = max(max_w, w)
        max_h = max(max_h, h)

    # --- shape 정규화 ---
    shape = shape.strip()
    shape = {"원": "원형", "원형": "원형",
             "정사각": "정사각", "정사각형": "정사각",
             "직사각(가로)": "직사각(가로)", "직사각(세로)": "직사각(세로)",
             "타원형(가로)": "타원형(가로)", "타원형(세로)": "타원형(세로)"}.get(shape, shape)

    # --- 폰트 자동 조절 로직 ---
    # 도형 / 배치별로 한 문자에 허용되는 최대 폭/높이를 계산하고,
    # 필요하면 폰트 크기를 작게 만들어 겹침(overflow)을 방지합니다.
    def get_truetype(size):
        try:
            return ImageFont.truetype(font_path, size)
        except Exception:
            return ImageFont.load_default()

    # helper: 재계산용 측정 함수
    def measure_max_char(font_obj):
        mw, mh = 0, 0
        for ch in text:
            try:
                b = draw.textbbox((0,0), ch, font=font_obj)
                w = b[2] - b[0]
                h = b[3] - b[1]
            except AttributeError:
                w, h = draw.textsize(ch, font=font_obj)
            mw = max(mw, w)
            mh = max(mh, h)
        return mw, mh

    # --- 각 도형별로 '한 문자당 허용 너비/높이'를 계산 ---
    if shape == "원형":
        # 외곽 반지름 기준으로 한 줄에 들어갈 문자 수를 예상 -> 한 문자당 허용 각도에 따른 호 길이 고려
        radius = min(width, height) / 2 - BORDER_WIDTH - 4
        # 외곽(가장 바깥 줄)에 들어갈 문자 수 최대값 예측 (간단히 n)
        # 한 문자에 허용되는 호 길이(대략) = 2 * pi * radius / max(chars_in_line)
        # 그러나 우리는 chars_per_line을 문자 분배 단계에서 결정하므로 여기에서는 안전한 축소만 계산
        # 단순 안전 측정: 원 전체 둘레를 고려해서 한 문자에 허용되는 호 길이 사용
        circumference = 2 * math.pi * max(1, radius)
        # 최소 허용 문자폭 = circumference / max(1, n) * 0.9(여유)
        safe_char_w = max(4, circumference / max(1, n) * 0.9)
        safe_char_h = max(4, (height / 2) * 0.8)
    elif shape.startswith("타원형"):
        # 타원은 수평/수직 반지름 차이를 고려
        rx = width / 2 - BORDER_WIDTH - 4
        ry = height / 2 - BORDER_WIDTH - 4
        # 근사 둘레 (Ramanujan)
        h_val = ((rx - ry)**2) / ((rx + ry)**2) if (rx + ry) != 0 else 0
        approx_perim = math.pi * (rx + ry) * (1 + 3*h_val/(10+math.sqrt(4-3*h_val)))
        safe_char_w = max(4, approx_perim / max(1, n) * 0.9)
        safe_char_h = max(4, min(rx, ry) * 0.8)
    elif shape == "정사각":
        inner_w = width - BORDER_WIDTH*4
        inner_h = height - BORDER_WIDTH*4
        # 한 줄당 문자 수는 나중에 결정되므로 안전폭은 inner_w / max(1,n)
        safe_char_w = max(4, inner_w / max(1, n) * 1.0)
        safe_char_h = max(4, inner_h / max(1, line_count) * 0.9 * line_spacing)
    elif shape == "직사각(가로)":
        inner_w = width - BORDER_WIDTH*4
        inner_h = height - BORDER_WIDTH*4
        safe_char_w = max(4, inner_w / max(1, n) * 1.0)
        safe_char_h = max(4, inner_h / max(1, line_count) * 0.9 * line_spacing)
    elif shape == "직사각(세로)":
        inner_w = width - BORDER_WIDTH*4
        inner_h = height - BORDER_WIDTH*4
        # 세로는 컬럼당 행 수가 line_count 해석에 따라 달라짐; 안전값은 세로 기준
        safe_char_w = max(4, inner_w / max(1, line_count) * 0.9)
        safe_char_h = max(4, inner_h / max(1, n) * 1.0 * line_spacing)
    else:
        # 기본 안전값
        safe_char_w = max(4, (width - BORDER_WIDTH*4) / max(1, n) * 0.9)
        safe_char_h = max(4, (height - BORDER_WIDTH*4) / max(1, line_count) * 0.9 * line_spacing)

    # 현재 font의 문자 크기
    cur_mw, cur_mh = measure_max_char(font)

    # 비율로 축소가 필요한지 판단
    scale_w = safe_char_w / cur_mw if cur_mw > 0 else 1.0
    scale_h = safe_char_h / cur_mh if cur_mh > 0 else 1.0
    scale = min(1.0, scale_w, scale_h)

    # 추가 여유 마진
    scale *= 0.98

    if scale < 0.99:
        # 새로운 폰트 크기 계산
        new_size = max(6, int(base_font_size * scale))
        font = get_truetype(new_size)
        # 재측정
        max_w, max_h = measure_max_char(font)
    else:
        # 기존 font 사용, max_w,max_h 이미 계산되어 있을 수 있음
        try:
            bbox = draw.textbbox((0,0), text[0], font=font)
            max_w = bbox[2] - bbox[0]
            max_h = bbox[3] - bbox[1]
        except Exception:
            max_w, max_h = max_w, max_h

    # 이제 실제 배치 로직 (원래 구성을 유지하되 줄수/줄간격 적용 및 겹침 방지)
    if shape == "원형":
        cx, cy = width // 2, height // 2
        base_radius = min(width, height) // 2 - BORDER_WIDTH - max(max_w, max_h)//2
        base_radius = max(10, int(base_radius))

        remaining = n
        start_idx = 0
        for line_idx in range(line_count):
            remain_lines = line_count - line_idx
            chars_this_line = max(1, math.ceil(remaining / remain_lines))
            remaining -= chars_this_line

            radius = base_radius - line_idx * int(max_h * 1.6 * line_spacing)
            if radius < 8:
                radius = 8

            angle_step = 360.0 / chars_this_line
            start_angle = -90.0
            for i in range(chars_this_line):
                if start_idx + i >= n:
                    break
                ch = text[start_idx + i]
                angle_deg = start_angle + i * angle_step
                angle_rad = math.radians(angle_deg)
                x = cx + int(radius * math.cos(angle_rad))
                y = cy + int(radius * math.sin(angle_rad))
                try:
                    bbox = draw.textbbox((0,0), ch, font=font)
                    w = bbox[2] - bbox[0]
                    h = bbox[3] - bbox[1]
                except AttributeError:
                    w, h = draw.textsize(ch, font=font)
                draw.text((x - w//2, y - h//2), ch, font=font, fill=color)
            start_idx += chars_this_line

    elif shape == "타원형(가로)":
        cx, cy = width // 2, height // 2
        rx = width // 2 - BORDER_WIDTH - max_w//2
        ry = height // 2 - BORDER_WIDTH - max_h//2
        remaining = n
        start_idx = 0
        for line_idx in range(line_count):
            chars_this_line = max(1, math.ceil(remaining / (line_count - line_idx)))
            remaining -= chars_this_line
            inner_rx = rx - line_idx * int(max_h * 1.5 * line_spacing)
            inner_ry = ry - line_idx * int(max_h * 1.5 * line_spacing)
            if inner_rx < 4: inner_rx = 4
            if inner_ry < 4: inner_ry = 4
            angle_step = 360.0 / chars_this_line
            start_angle = -90.0
            for i in range(chars_this_line):
                if start_idx + i >= n:
                    break
                ch = text[start_idx + i]
                angle_deg = start_angle + i * angle_step
                angle_rad = math.radians(angle_deg)
                x = cx + int(inner_rx * math.cos(angle_rad))
                y = cy + int(inner_ry * math.sin(angle_rad))
                try:
                    bbox = draw.textbbox((0,0), ch, font=font)
                    w = bbox[2] - bbox[0]
                    h = bbox[3] - bbox[1]
                except AttributeError:
                    w, h = draw.textsize(ch, font=font)
                draw.text((x - w//2, y - h//2), ch, font=font, fill=color)
            start_idx += chars_this_line

    elif shape == "타원형(세로)":
        # 동일하지만 rx/ry 교환 처리
        cx, cy = width // 2, height // 2
        rx = width // 2 - BORDER_WIDTH - max_w//2
        ry = height // 2 - BORDER_WIDTH - max_h//2
        remaining = n
        start_idx = 0
        for line_idx in range(line_count):
            chars_this_line = max(1, math.ceil(remaining / (line_count - line_idx)))
            remaining -= chars_this_line
            inner_rx = rx - line_idx * int(max_h * 1.5 * line_spacing)
            inner_ry = ry - line_idx * int(max_h * 1.5 * line_spacing)
            if inner_rx < 4: inner_rx = 4
            if inner_ry < 4: inner_ry = 4
            angle_step = 360.0 / chars_this_line
            start_angle = -90.0
            for i in range(chars_this_line):
                if start_idx + i >= n:
                    break
                ch = text[start_idx + i]
                angle_deg = start_angle + i * angle_step
                angle_rad = math.radians(angle_deg)
                x = cx + int(inner_rx * math.cos(angle_rad))
                y = cy + int(inner_ry * math.sin(angle_rad))
                try:
                    bbox = draw.textbbox((0,0), ch, font=font)
                    w = bbox[2] - bbox[0]
                    h = bbox[3] - bbox[1]
                except AttributeError:
                    w, h = draw.textsize(ch, font=font)
                draw.text((x - w//2, y - h//2), ch, font=font, fill=color)
            start_idx += chars_this_line

    elif shape == "정사각":
        inner_left = BORDER_WIDTH * 2
        inner_right = width - BORDER_WIDTH * 2
        inner_top = BORDER_WIDTH * 2
        inner_bottom = height - BORDER_WIDTH * 2

        available_w = max(10, inner_right - inner_left)
        available_h = max(10, inner_bottom - inner_top)

        chars_per_line = max(1, math.ceil(n / line_count))
        # 한 줄당 허용 문자수(가로방향) 계산
        max_chars_fit = max(1, int(available_w / (max_w * 1.05)))  # 5% 여유
        if max_chars_fit < chars_per_line:
            chars_per_line = max_chars_fit

        idx = 0
        # 줄간격 반영
        line_h = available_h / line_count * line_spacing
        for row in range(line_count):
            if idx >= n:
                break
            this_count = min(chars_per_line, n - idx)
            if this_count == 1:
                gap = 0
            else:
                gap = available_w / (this_count - 1)
                gap = max(gap, max_w * 1.05)
            y_center = inner_top + row * (available_h / line_count) + (available_h / line_count) / 2
            total_row_width = (this_count - 1) * gap if this_count > 1 else max_w
            start_x = inner_left + (available_w - total_row_width) / 2
            for i in range(this_count):
                if idx >= n:
                    break
                ch = text[idx]
                x_center = start_x + i * gap
                try:
                    bbox = draw.textbbox((0,0), ch, font=font)
                    w = bbox[2] - bbox[0]
                    h = bbox[3] - bbox[1]
                except AttributeError:
                    w, h = draw.textsize(ch, font=font)
                draw.text((x_center - w/2, y_center - h/2), ch, font=font, fill=color)
                idx += 1

    elif shape == "직사각(가로)":
        inner_left = BORDER_WIDTH * 2
        inner_right = width - BORDER_WIDTH * 2
        inner_top = BORDER_WIDTH * 2
        inner_bottom = height - BORDER_WIDTH * 2

        available_w = max(10, inner_right - inner_left)
        available_h = max(10, inner_bottom - inner_top)

        chars_per_line = max(1, math.ceil(n / line_count))
        max_chars_fit = max(1, int(available_w / (max_w * 1.05)))
        if max_chars_fit < chars_per_line:
            chars_per_line = max_chars_fit

        idx = 0
        for row in range(line_count):
            if idx >= n:
                break
            this_count = min(chars_per_line, n - idx)
            if this_count == 1:
                gap = 0
            else:
                gap = available_w / (this_count - 1)
                gap = max(gap, max_w * 1.05)
            y_center = inner_top + row * (available_h / line_count) + (available_h / line_count) / 2
            total_row_width = (this_count - 1) * gap if this_count > 1 else max_w
            start_x = inner_left + (available_w - total_row_width) / 2
            for i in range(this_count):
                if idx >= n:
                    break
                ch = text[idx]
                x_center = start_x + i * gap
                try:
                    bbox = draw.textbbox((0,0), ch, font=font)
                    w = bbox[2] - bbox[0]
                    h = bbox[3] - bbox[1]
                except AttributeError:
                    w, h = draw.textsize(ch, font=font)
                draw.text((x_center - w/2, y_center - h/2), ch, font=font, fill=color)
                idx += 1

    elif shape == "직사각(세로)":
        inner_left = BORDER_WIDTH * 2
        inner_right = width - BORDER_WIDTH * 2
        inner_top = BORDER_WIDTH * 2
        inner_bottom = height - BORDER_WIDTH * 2

        available_w = max(10, inner_right - inner_left)
        available_h = max(10, inner_bottom - inner_top)

        cols = line_count
        chars_per_col = max(1, math.ceil(n / cols))
        # 세로에 맞게 fit 행수 계산
        max_rows_fit = max(1, int(available_h / (max_h * 1.05)))
        if chars_per_col > max_rows_fit:
            chars_per_col = max_rows_fit

        idx = 0
        col_w = available_w / cols
        for col in range(cols):
            if idx >= n:
                break
            this_count = min(chars_per_col, n - idx)
            if this_count == 1:
                vgap = 0
            else:
                vgap = available_h / (this_count - 1)
                vgap = max(vgap, max_h * 1.05)
            x_center = inner_left + col * col_w + col_w / 2
            total_col_height = (this_count - 1) * vgap if this_count > 1 else max_h
            start_y = inner_top + (available_h - total_col_height) / 2
            for r in range(this_count):
                if idx >= n:
                    break
                ch = text[idx]
                y_center = start_y + r * vgap
                try:
                    bbox = draw.textbbox((0,0), ch, font=font)
                    w = bbox[2] - bbox[0]
                    h = bbox[3] - bbox[1]
                except AttributeError:
                    w, h = draw.textsize(ch, font=font)
                draw.text((x_center - w/2, y_center - h/2), ch, font=font, fill=color)
                idx += 1

# --- 도장 기본 텍스트 배치 함수들 ---
def draw_text_left_to_right(draw, text, font, width, height, color):
    # 도장 중앙에 좌->우 방향 텍스트 배치
    try:
        bbox = draw.textbbox((0,0), text, font=font)
        w = bbox[2] - bbox[0]
        h = bbox[3] - bbox[1]
    except Exception:
        w, h = draw.textsize(text, font=font)
    x = (width - w) // 2
    y = (height - h) // 2
    draw.text((x, y), text, font=font, fill=color)

def draw_text_top_to_bottom(draw, text, font, width, height, color):
    # 도장 중앙에 상->하 방향 텍스트 배치
    total_height = 0
    char_sizes = []
    for ch in text:
        try:
            bbox = draw.textbbox((0,0), ch, font=font)
            w = bbox[2] - bbox[0]
            h = bbox[3] - bbox[1]
        except Exception:
            w, h = draw.textsize(ch, font=font)
        char_sizes.append((w, h))
        total_height += h
    y = (height - total_height) // 2
    x = width // 2
    for i, ch in enumerate(text):
        w, h = char_sizes[i]
        draw.text((x - w//2, y), ch, font=font, fill=color)
        y += h

def draw_text_joseon_style(draw, text, font, width, height, color):
    # 조선체 스타일 (약간 왼쪽 정렬 좌->우, 세로 간격 넉넉히)
    x = BORDER_WIDTH*3
    try:
        y = (height - len(text) * font.size) // 2
    except Exception:
        y = (height - len(text) * 12) // 2
    for ch in text:
        try:
            bbox = draw.textbbox((0,0), ch, font=font)
            w = bbox[2] - bbox[0]
            h = bbox[3] - bbox[1]
        except Exception:
            w, h = draw.textsize(ch, font=font)
        draw.text((x, y), ch, font=font, fill=color)
        y += h + 2

# --- 둥근 사각형 그리기 보조 함수 ---
def draw_rounded_rectangle(draw, box, radius, outline, width):
    left, top, right, bottom = box
    draw.line([(left+radius, top), (right-radius, top)], fill=outline, width=width)
    draw.line([(left+radius, bottom), (right-radius, bottom)], fill=outline, width=width)
    draw.line([(left, top+radius), (left, bottom-radius)], fill=outline, width=width)
    draw.line([(right, top+radius), (right, bottom-radius)], fill=outline, width=width)
    draw.arc([left, top, left+2*radius, top+2*radius], 180, 270, fill=outline, width=width)
    draw.arc([right-2*radius, top, right, top+2*radius], 270, 360, fill=outline, width=width)
    draw.arc([right-2*radius, bottom-2*radius, right, bottom], 0, 90, fill=outline, width=width)
    draw.arc([left, bottom-2*radius, left+2*radius, bottom], 90, 180, fill=outline, width=width)

# --- 저장 함수 ---
def save_stamp():
    try:
        img_obj = preview_label.image_obj
        if img_obj is None:
            status_label.config(text="저장할 이미지가 없습니다.", fg=fg_color_error)
            return
        file_path = filedialog.asksaveasfilename(defaultextension=".png",
                                                filetypes=[("PNG 파일", "*.png")])
        if not file_path:
            return
        img_obj.save(file_path)
        status_label.config(text=f"저장 성공: {file_path}", fg=fg_color_ok)
    except Exception as e:
        status_label.config(text=f"저장 실패: {e}", fg=fg_color_error)

# --- 도장 모양 콤보박스 변경 이벤트 ---
def on_stamp_type_change(event=None):
    selected = stamp_type_combo.get()
    if selected == "원형":
        sub_stamp_shape_combo.config(state="readonly")
        sub_stamp_shape_combo['values'] = ["원", "타원형(가로)", "타원형(세로)"]
        sub_stamp_shape_combo.current(0)  # 기본값 "원"
        line_count_combo.set("1")
        line_count_combo.config(state="disabled")
        line_spacing_combo.set("보통")
        line_spacing_combo.config(state="disabled")
    elif selected == "문자 to 도장모양":
        sub_stamp_shape_combo.config(state="readonly")
        sub_stamp_shape_combo['values'] = SUB_STAMP_SHAPES_RECT  # 기존 값 유지
        sub_stamp_shape_combo.current(0)
        line_count_combo.config(state="readonly")
        line_count_combo.current(0)
        line_spacing_combo.config(state="readonly")
        line_spacing_combo.current(1)  # 보통
    else:
        sub_stamp_shape_combo.set("")
        sub_stamp_shape_combo.config(state="disabled")
        line_count_combo.set("1")
        line_count_combo.config(state="disabled")
        line_spacing_combo.set("보통")
        line_spacing_combo.config(state="disabled")

# --- 도장 크기 콤보 변경 이벤트 (사용자 지정 활성화) ---
def on_size_change(event=None):
    sel = size_combo.get()
    if sel == "사용자 지정":
        custom_width_entry.config(state="normal")
        custom_height_entry.config(state="normal")
    else:
        custom_width_entry.delete(0, tk.END)
        custom_height_entry.delete(0, tk.END)
        custom_width_entry.config(state="disabled")
        custom_height_entry.config(state="disabled")

# --- 윈도우 위치: 해상도 자동 감지해서 대각선 1 2 교차 지점(화면 중앙)으로 이동 ---
def position_window_at_diagonals_intersection():
    """
    화면의 해상도를 자동으로 감지하여
    '대각선 1과 2의 교차 지점' 즉 화면의 중앙으로 창의 중심을 맞춰 배치합니다.
    기존 코드 구조와 기능은 절대 훼손하지 않고, 위치 지정만 수행합니다.
    """
    try:
        # 먼저 레이아웃이 계산되도록 처리
        root.update_idletasks()

        # 화면 해상도(가로/세로)
        screen_w = root.winfo_screenwidth()
        screen_h = root.winfo_screenheight()

        # 대각선 교차점 = 화면 중앙
        cx = screen_w // 2
        cy = screen_h // 2

        # 창의 요청된 크기(또는 현재 크기)를 얻음
        # winfo_width/height는 아직 1일 수 있으므로 요청 크기를 우선 사용
        win_w = root.winfo_reqwidth()
        win_h = root.winfo_reqheight()

        # 너무 작게 잡히는 경우를 방지하기 위한 최소값 보장
        win_w = max(win_w, 400)
        win_h = max(win_h, 300)

        # 창의 좌상단 좌표를 중앙 기준으로 계산
        x = cx - win_w // 2
        y = cy - win_h // 2

        # 화면 밖으로 나가지 않도록 클램프
        x = max(0, x)
        y = max(0, y)

        root.geometry(f"{win_w}x{win_h}+{x}+{y}")
    except Exception as e:
        # 위치 지정 실패해도 앱은 정상 동작해야 하므로 상태 표시만 함
        try:
            status_label.config(text=f"창 위치 지정 실패: {e}", fg=fg_color_error)
        except Exception:
            pass
        print(f"창 위치 지정 실패: {e}")

# --- 기본 UI 구성 ---
root = tk.Tk()
root.title("도장 생성기")
try:
    root.iconbitmap(ICON_PATH)
except Exception:
    pass

# 다크 모드 감지 후 배경색 설정
if detect_dark_mode():
    bg_color = "#222222"
    fg_color = "#DDDDDD"
else:
    bg_color = "#FFFFFF"
    fg_color = "#000000"

root.config(bg=bg_color)

# 입력 프레임
input_frame = tk.Frame(root, bg=bg_color)
input_frame.pack(padx=10, pady=10, fill="x")

tk.Label(input_frame, text="이름:", bg=bg_color, fg=fg_color).grid(row=0, column=0, sticky="w")
name_entry = tk.Entry(input_frame, font=("맑은 고딕", 14))
name_entry.grid(row=0, column=1, sticky="ew", padx=5)
input_frame.grid_columnconfigure(1, weight=1)

tk.Label(input_frame, text="폰트:", bg=bg_color, fg=fg_color).grid(row=1, column=0, sticky="w")
font_name_combo = ttk.Combobox(input_frame, values=[f[1] for f in fonts], state="readonly")
font_name_combo.grid(row=1, column=1, sticky="ew", padx=5)
font_name_combo.current(0)

tk.Label(input_frame, text="폰트 크기:", bg=bg_color, fg=fg_color).grid(row=2, column=0, sticky="w")
font_size_combo = ttk.Combobox(input_frame, values=font_sizes, state="readonly")
font_size_combo.grid(row=2, column=1, sticky="ew", padx=5)
font_size_combo.current(font_sizes.index("80 (추천)"))

tk.Label(input_frame, text="텍스트 방향:", bg=bg_color, fg=fg_color).grid(row=3, column=0, sticky="w")
text_direction_combo = ttk.Combobox(input_frame, values=TEXT_DIRECTIONS, state="readonly")
text_direction_combo.grid(row=3, column=1, sticky="ew", padx=5)
text_direction_combo.current(0)

tk.Label(input_frame, text="도장 모양:", bg=bg_color, fg=fg_color).grid(row=4, column=0, sticky="w")
stamp_type_combo = ttk.Combobox(input_frame, values=STAMP_TYPES, state="readonly")
stamp_type_combo.grid(row=4, column=1, sticky="ew", padx=5)
stamp_type_combo.current(0)
stamp_type_combo.bind("<<ComboboxSelected>>", on_stamp_type_change)

tk.Label(input_frame, text="서브 도장 모양:", bg=bg_color, fg=fg_color).grid(row=5, column=0, sticky="w")
sub_stamp_shape_combo = ttk.Combobox(input_frame, values=SUB_STAMP_SHAPES, state="disabled")
sub_stamp_shape_combo.grid(row=5, column=1, sticky="ew", padx=5)

# 도장 크기 콤보박스 (사용자 선택 가능)
tk.Label(input_frame, text="도장 크기 (WxH):", bg=bg_color, fg=fg_color).grid(row=6, column=0, sticky="w")
size_options = ["200x200", "300x300", "400x400 (기본)", "500x400", "600x450", "사용자 지정"]
size_combo = ttk.Combobox(input_frame, values=size_options, state="readonly")
size_combo.grid(row=6, column=1, sticky="ew", padx=5)
# 기본값 선택 (400x400)
size_combo.current(2)
size_combo.bind("<<ComboboxSelected>>", on_size_change)

# 사용자 지정 입력란 (비활성 기본)
custom_size_frame = tk.Frame(input_frame, bg=bg_color)
custom_size_frame.grid(row=7, column=1, sticky="ew", padx=5, pady=(2,6))
custom_width_entry = tk.Entry(custom_size_frame, width=8, state="disabled")
custom_width_entry.pack(side="left", padx=(0,4))
custom_width_entry.insert(0, "")
custom_height_entry = tk.Entry(custom_size_frame, width=8, state="disabled")
custom_height_entry.pack(side="left", padx=(0,4))
custom_height_entry.insert(0, "")

tk.Label(input_frame, text="줄 수 (문자 to 도장모양):", bg=bg_color, fg=fg_color).grid(row=8, column=0, sticky="w")
line_count_values = [str(i) for i in range(1, 7)]
line_count_combo = ttk.Combobox(input_frame, values=line_count_values, state="disabled")
line_count_combo.grid(row=8, column=1, sticky="ew", padx=5)
line_count_combo.current(0)

tk.Label(input_frame, text="줄간격:", bg=bg_color, fg=fg_color).grid(row=9, column=0, sticky="w")
line_spacing_combo = ttk.Combobox(input_frame, values=["좁게","보통","넉넉"], state="disabled")
line_spacing_combo.grid(row=9, column=1, sticky="ew", padx=5)
line_spacing_combo.current(1)

seal_color = tk.StringVar(value="#FF0000")
tk.Label(input_frame, text="도장 색상:", bg=bg_color, fg=fg_color).grid(row=10, column=0, sticky="w")
color_button = tk.Button(input_frame, text="색상 선택", command=lambda: choose_color())
color_button.grid(row=10, column=1, sticky="w", padx=5)

def choose_color():
    color_code = colorchooser.askcolor(seal_color.get(), title="도장 색상 선택")
    if color_code[1]:
        seal_color.set(color_code[1])
        generate_seal()

# 미리보기 프레임
preview_frame = tk.Frame(root, bg=bg_color)
preview_frame.pack(padx=10, pady=10)

preview_label = tk.Label(preview_frame, bg=bg_color)
preview_label.pack()

# 상태 표시줄
status_label = tk.Label(root, text="", bg=bg_color, fg=fg_color)
status_label.pack(pady=(0,10))

# 버튼 프레임
button_frame = tk.Frame(root, bg=bg_color)
button_frame.pack(padx=10, pady=10, fill="x")

generate_button = tk.Button(button_frame, text="도장 생성", command=generate_seal)
generate_button.pack(side="left", padx=5)

save_button = tk.Button(button_frame, text="저장", command=save_stamp)
save_button.pack(side="left", padx=5)

# --- 초기 도장 생성 ---
preview_label.image_obj = None
generate_seal()

# --- 창을 화면의 대각선 교차점(중앙)으로 위치시키기 ---
position_window_at_diagonals_intersection()

root.mainloop()