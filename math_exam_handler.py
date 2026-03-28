# -*- coding: utf-8 -*-
"""
math_exam_handler.py
====================
Module xử lý đặc thù cho môn TOÁN trong hệ thống Check Đề.

Giải quyết các vấn đề:
1. WMF/Equation 3.0 trong docx (dùng PDF render + AI vision)
2. Subject-aware parsing: phân biệt 3 phần PHẦN I/II/III
3. Normalize so sánh đáp án điền số (2,4 vs 2.4)
4. Prompt riêng cho Toán (khác Tiếng Anh)

Cách dùng:
    from math_exam_handler import (
        detect_subject,
        parse_math_exam_docx,
        parse_math_answer_key,
        build_math_matching_prompt,
        normalize_math_answer,
        verify_math_answers,
    )
"""

import os
import re
import zipfile
from dataclasses import dataclass, field
from typing import Optional

# Đọc prompt từ file .txt (nếu có)
try:
    from prompt_loader import (
        build_prompt_header,
        get_output_format,
        get_vision_prompt,
        get_part_labels,
        get_max_pages,
        get_subject_label,
        get_parts_config,
        get_part_boundary_patterns,
        get_part_type,
    )
    PROMPT_LOADER_AVAILABLE = True
except ImportError:
    PROMPT_LOADER_AVAILABLE = False

# ─────────────────────────────────────────────
# 1. DATA CLASSES
# ─────────────────────────────────────────────

@dataclass
class MathQuestion:
    """Một câu hỏi trong đề Toán (có thể là trắc nghiệm, đúng/sai hoặc điền số)."""
    number: int                    # Số thứ tự câu trong phần (1, 2, ...)
    global_number: int             # Số thứ tự tuyệt đối trong toàn đề (1..18)
    part: int                      # Phần: 1=Trắc nghiệm, 2=Đúng/Sai, 3=Điền số, 4=Tự luận
    q_type: str                    # 'mc' | 'tf' | 'fill' | 'essay'
    text: str                      # Nội dung câu hỏi (sau khi strip công thức stub)
    options: dict = field(default_factory=dict)  # {'A': '...', 'B': '...', 'C': '...', 'D': '...'}
    has_equation: bool = False     # True nếu câu này có công thức (WMF/OMML)
    raw_block: str = ''            # Raw text block của câu (để debug)

    def __repr__(self):
        return f"[Part{self.part}/Q{self.global_number}/{self.q_type}] {self.text[:60]}"


# ─────────────────────────────────────────────
# 2. SUBJECT DETECTION
# ─────────────────────────────────────────────

def detect_subject(filepath: str, lines: list[str] = None) -> str:
    """
    Phát hiện môn thi dựa trên tên file và nội dung.

    Returns:
        'math'    — Toán (có 3 phần trắc nghiệm / đúng-sai / điền số)
        'english' — Tiếng Anh (chỉ trắc nghiệm, có reading/listening)
        'other'   — Các môn khác (trắc nghiệm đơn thuần)
    """
    basename = os.path.basename(filepath).lower()

    # Phát hiện qua tên file
    math_keywords = ['toan', 'toán', 'math']
    english_keywords = ['ta', 'tiengAnh', 'english', 'anh']

    for kw in math_keywords:
        if kw in basename:
            return 'math'
    for kw in english_keywords:
        if kw in basename:
            return 'english'

    # Phát hiện qua nội dung (nếu có)
    if lines:
        header = ' '.join(lines[:30]).lower()
        if any(k in header for k in ['toán', 'toan', 'math']):
            return 'math'
        if any(k in header for k in ['tiếng anh', 'english', 'reading', 'listening']):
            return 'english'
        # Đặc trưng Toán: có phần đúng/sai và điền ngắn
        full = '\n'.join(lines)
        has_tf = bool(re.search(r'PHẦN\s+II|đúng\s+hoặc\s+sai|đúng/sai', full, re.I))
        has_fill = bool(re.search(r'PHẦN\s+III|trả\s+lời\s+ngắn|điền', full, re.I))
        if has_tf and has_fill:
            return 'math'

    return 'other'


# ─────────────────────────────────────────────
# 3. WMF DETECTION
# ─────────────────────────────────────────────

def docx_has_wmf_equations(filepath: str) -> bool:
    """
    Kiểm tra file docx có dùng Equation 3.0 (WMF embed) không.
    Nếu True → cần dùng PDF render thay vì OMML extraction.
    """
    try:
        with zipfile.ZipFile(filepath) as z:
            names = z.namelist()
            # Có file WMF/EMF trong media
            has_wmf = any(n.endswith('.wmf') or n.endswith('.emf') for n in names)
            if not has_wmf:
                return False
            # Kiểm tra document.xml: có w:object (OLE embed) không
            if 'word/document.xml' in names:
                xml = z.read('word/document.xml').decode('utf-8', errors='ignore')
                has_ole = 'w:object' in xml
                has_omml = 'oMath' in xml or 'm:oMath' in xml
                # WMF+OLE nhưng không có OMML → Equation 3.0
                return has_wmf and has_ole and not has_omml
    except Exception:
        pass
    return False


# ─────────────────────────────────────────────
# 4. MATH EXAM PARSER
# ─────────────────────────────────────────────

# Regex nhận diện tiêu đề phần
_PART_PATTERNS = {
    1: re.compile(r'PHẦN\s+I\b.*?(?:nhiều\s+phương\s+án|trắc\s+nghiệm\s+nhiều)', re.I | re.S),
    2: re.compile(r'PHẦN\s+II\b.*?(?:đúng\s+sai|đúng\s+hoặc\s+sai)', re.I | re.S),
    3: re.compile(r'PHẦN\s+III\b.*?(?:trả\s+lời\s+ngắn|điền|ngắn)', re.I | re.S),
    4: re.compile(r'PHẦN\s+IV\b.*?(?:tự\s+luận)', re.I | re.S),
}

_QUESTION_PATTERN = re.compile(r'(?:^|\n)\s*Câu\s+(\d+)\s*[.:]?\s*', re.I)
_OPTION_STANDALONE = re.compile(r'^\s*([A-D])\.\s+(.+)', re.I | re.M)
_OPTION_INLINE = re.compile(r'([A-D])\.\s+(.+?)(?=\s+[A-D]\.\s|\s*$)', re.I)

# Nhận biết stub công thức (pandoc output hoặc text extract thiếu)
_EQUATION_STUB = re.compile(
    r'!\[.*?\]\(media/image\d+\.(wmf|emf|png|jpg)\)'  # pandoc WMF stub
    r'|<math[^>]*>.*?</math>'                           # MathML
    r'|\[EQUATION\]'                                    # placeholder
    r'|\[FORMULA\]',
    re.I | re.S
)


def _detect_part_boundaries(full_text: str, subject: str = 'math') -> dict:
    """
    Tìm vị trí bắt đầu của từng phần trong đề dựa trên các tiêu đề cố định.
    Sử dụng PHẦN IV làm điểm kết thúc cho PHẦN III.
    """
    boundaries = {}

    # Sử dụng Regex với độ tuỳ biến khoảng trắng (đề phòng file Word bị dư dấu cách)
    # Khớp chính xác với cấu trúc bạn yêu cầu
    part_markers = [
        (1, re.compile(r'PHẦN\s+I\.\s*Câu\s+trắc\s+nghiệm\s+nhiều\s+phương\s+án', re.I)),
        (2, re.compile(r'PHẦN\s+II\.\s*Câu\s+trắc\s+nghiệm\s+đúng\s+sai', re.I)),
        (3, re.compile(r'PHẦN\s+III\.\s*Câu\s+trắc\s+nghiệm\s+trả\s+lời\s+ngắn', re.I)),
        (4, re.compile(r'PHẦN\s+IV\.\s*Tự\s+luận', re.I)),
    ]

    positions = []
    
    # Quét văn bản để tìm vị trí bắt đầu của các phần
    for part_num, pattern in part_markers:
        m = pattern.search(full_text)
        if m:
            positions.append((part_num, m.start()))

    # Fallback dự phòng: Nếu docx bị lỗi format mất chữ, tìm theo "PHẦN I/II/III/IV"
    if not positions:
        fallback_markers = [
            (1, re.compile(r'PHẦN\s+I\b', re.I)),
            (2, re.compile(r'PHẦN\s+II\b', re.I)),
            (3, re.compile(r'PHẦN\s+III\b', re.I)),
            (4, re.compile(r'PHẦN\s+IV\b', re.I)),
        ]
        for part_num, pattern in fallback_markers:
            m = pattern.search(full_text)
            if m:
                positions.append((part_num, m.start()))

    # Sắp xếp các phần theo thứ tự xuất hiện trong văn bản
    positions.sort(key=lambda x: x[1])

    # Cắt ranh giới: Phần hiện tại sẽ bắt đầu từ vị trí của nó, và kết thúc ở vị trí của Phần tiếp theo
    for i, (part_num, start) in enumerate(positions):
        end = positions[i + 1][1] if i + 1 < len(positions) else len(full_text)
        boundaries[part_num] = (start, end)

    return boundaries


def _parse_mc_questions(block: str, part_num: int, global_offset: int = 0) -> list[MathQuestion]:
    """Parse câu trắc nghiệm nhiều phương án (PHẦN I)."""
    questions = []
    q_matches = list(_QUESTION_PATTERN.finditer(block))

    for i, match in enumerate(q_matches):
        q_num = int(match.group(1))
        q_start = match.end()
        q_end = q_matches[i + 1].start() if i + 1 < len(q_matches) else len(block)
        q_block = block[q_start:q_end].strip()

        # Phát hiện stub công thức
        has_eq = bool(_EQUATION_STUB.search(q_block))
        # Xóa stub để lấy text thuần
        clean_block = _EQUATION_STUB.sub('[EQ]', q_block)

        options = {}
        opt_matches = list(_OPTION_STANDALONE.finditer(clean_block))

        if opt_matches:
            q_text = clean_block[:opt_matches[0].start()].strip()
            for j, om in enumerate(opt_matches):
                letter = om.group(1).upper()
                end = opt_matches[j + 1].start() if j + 1 < len(opt_matches) else len(clean_block)
                opt_val = clean_block[om.start():end].strip()
                opt_val = re.sub(r'^[A-D]\.\s*', '', opt_val).strip()
                options[letter] = opt_val
        else:
            # Thử inline options
            lines = clean_block.split('\n')
            q_text = clean_block
            for line in lines:
                inline = list(_OPTION_INLINE.finditer(line))
                if len(inline) >= 2:
                    for m in inline:
                        options[m.group(1).upper()] = m.group(2).strip()
                    idx = clean_block.find(line)
                    q_text = clean_block[:idx].strip()
                    break

        questions.append(MathQuestion(
            number=q_num,
            global_number=q_num + global_offset,
            part=part_num,
            q_type='mc',
            text=q_text[:500],
            options=options,
            has_equation=has_eq,
            raw_block=q_block[:300],
        ))

    return questions


def _parse_tf_questions(block: str, part_num: int, global_offset: int = 0) -> list[MathQuestion]:
    """
    Parse câu Đúng/Sai (PHẦN II).
    Mỗi câu có 4 ý a), b), c), d).
    """
    questions = []
    q_matches = list(_QUESTION_PATTERN.finditer(block))

    for i, match in enumerate(q_matches):
        q_num = int(match.group(1))
        q_start = match.end()
        q_end = q_matches[i + 1].start() if i + 1 < len(q_matches) else len(block)
        q_block = block[q_start:q_end].strip()

        has_eq = bool(_EQUATION_STUB.search(q_block))
        clean_block = _EQUATION_STUB.sub('[EQ]', q_block)

        # Tìm các ý a), b), c), d)
        sub_pattern = re.compile(r'([a-d])\)', re.I)
        sub_matches = list(sub_pattern.finditer(clean_block))

        options = {}
        if sub_matches:
            q_text = clean_block[:sub_matches[0].start()].strip()
            for j, sm in enumerate(sub_matches):
                letter = sm.group(1).lower()
                end = sub_matches[j + 1].start() if j + 1 < len(sub_matches) else len(clean_block)
                val = clean_block[sm.end():end].strip()
                options[letter] = val[:200]
        else:
            q_text = clean_block

        questions.append(MathQuestion(
            number=q_num,
            global_number=q_num + global_offset,
            part=part_num,
            q_type='tf',
            text=q_text[:500],
            options=options,
            has_equation=has_eq,
            raw_block=q_block[:300],
        ))

    return questions


def _parse_fill_questions(block: str, part_num: int, global_offset: int = 0) -> list[MathQuestion]:
    """Parse câu điền số ngắn (PHẦN III)."""
    questions = []
    q_matches = list(_QUESTION_PATTERN.finditer(block))

    for i, match in enumerate(q_matches):
        q_num = int(match.group(1))
        q_start = match.end()
        q_end = q_matches[i + 1].start() if i + 1 < len(q_matches) else len(block)
        q_block = block[q_start:q_end].strip()

        has_eq = bool(_EQUATION_STUB.search(q_block))
        clean_block = _EQUATION_STUB.sub('[EQ]', q_block).strip()

        questions.append(MathQuestion(
            number=q_num,
            global_number=q_num + global_offset,
            part=part_num,
            q_type='fill',
            text=clean_block[:500],
            options={},
            has_equation=has_eq,
            raw_block=q_block[:300],
        ))

    return questions


def parse_math_exam_from_text(full_text: str, source_name: str = '', subject: str = 'math') -> list[MathQuestion]:
    """
    Parse toàn bộ đề Toán từ text đã extract.
    Tự động phát hiện và phân loại PHẦN I/II/III/IV.

    Args:
        full_text: Toàn bộ text của đề (từ docx hoặc pdf)
        source_name: Tên file để log

    Returns:
        Danh sách MathQuestion theo thứ tự toàn đề
    """
    all_questions: list[MathQuestion] = []
    boundaries = _detect_part_boundaries(full_text, subject=subject)

    if not boundaries:
        # Fallback: parse như trắc nghiệm thông thường (không có phần)
        print(f"  ⚠️  [{source_name}] Không phát hiện cấu trúc PHẦN I/II/III → parse as MC")
        q_matches = list(_QUESTION_PATTERN.finditer(full_text))
        for i, match in enumerate(q_matches):
            q_num = int(match.group(1))
            q_start = match.end()
            q_end = q_matches[i + 1].start() if i + 1 < len(q_matches) else len(full_text)
            q_block = full_text[q_start:q_end].strip()
            has_eq = bool(_EQUATION_STUB.search(q_block))
            clean = _EQUATION_STUB.sub('[EQ]', q_block)
            all_questions.append(MathQuestion(
                number=q_num, global_number=q_num, part=1,
                q_type='mc', text=clean[:300], has_equation=has_eq,
            ))
        return all_questions

    # Lấy q_type từ PARTS_CONFIG trong file .txt nếu có, fallback hardcode
    if PROMPT_LOADER_AVAILABLE:
        _pcfg = get_parts_config(subject)
        part_info = {
            cfg['part']: (cfg['q_type'].upper(), cfg['q_type'])
            for cfg in _pcfg
        } if _pcfg else {
            1: ('Trắc nghiệm MC', 'mc'),
            2: ('Đúng/Sai', 'tf'),
            3: ('Điền số ngắn', 'fill'),
            4: ('Tự luận', 'essay'),
        }
    else:
        part_info = {
            1: ('Trắc nghiệm MC', 'mc'),
            2: ('Đúng/Sai', 'tf'),
            3: ('Điền số ngắn', 'fill'),
            4: ('Tự luận', 'essay'),
        }

    # Đếm câu hỏi ở phần trước để tính global_number
    # Đếm câu hỏi ở phần trước để tính global_number
    cumulative = 0

    for part_num in sorted(boundaries.keys()):
        # CHỈ ĐẠO QUAN TRỌNG: Bỏ qua hoàn toàn Phần 4 (Tự luận) và các phần sau đó
        if part_num >= 4:
            print(f"  ⏭️  [{source_name}] Đã phát hiện và BỎ QUA hoàn toàn PHẦN {part_num} (Tự luận)")
            continue 

        start, end = boundaries[part_num]
        block = full_text[start:end]
        label, qtype = part_info.get(part_num, ('?', 'mc'))

        if part_num == 1:
            qs = _parse_mc_questions(block, part_num, global_offset=0)
        elif part_num == 2:
            qs = _parse_tf_questions(block, part_num, global_offset=cumulative)
        elif part_num == 3:
            qs = _parse_fill_questions(block, part_num, global_offset=cumulative)
        else:
            qs = []

        eq_count = sum(1 for q in qs if q.has_equation)
        print(f"  📐 [{source_name}] PHẦN {part_num} ({label}): {len(qs)} câu"
              + (f", {eq_count} câu có công thức" if eq_count else ""))

        all_questions.extend(qs)
        if qs:
            cumulative = max(q.global_number for q in qs)

    return all_questions


# ─────────────────────────────────────────────
# 5. ANSWER KEY NORMALIZATION
# ─────────────────────────────────────────────

def normalize_math_answer(raw_val) -> str:
    """
    Chuẩn hóa đáp án Toán để so sánh nhất quán.

    Xử lý:
    - MC: 'A', 'B', 'C', 'D' → uppercase
    - Đúng/Sai: 'SSSĐ', 'ĐSS' → uppercase, chuẩn hóa Đ/S
    - Điền số: '2,4' → '2.4', '13,4' → '13.4'
    - Mixed/None → ''
    """
    if raw_val is None:
        return ''

    val = str(raw_val).strip()

    # Empty
    if not val:
        return ''

    # MC: single letter
    if len(val) == 1 and val.upper() in 'ABCD':
        return val.upper()

    # Đúng/Sai: chuỗi gồm Đ/S/đ/s (có thể 4 ký tự)
    val_upper = val.upper()
    # Chuẩn hóa: Đ → D (để compare), nhưng giữ nguyên chữ để display
    ds_normalized = val_upper.replace('Đ', 'D')
    if re.fullmatch(r'[DS]{2,6}', ds_normalized):
        return val.upper()

    # Số thập phân: thay dấu phẩy thành chấm
    num_val = val.replace(',', '.')
    try:
        f = float(num_val)
        # Bỏ phần thập phân không cần thiết: 4.0 → '4'
        if f == int(f):
            return str(int(f))
        # Làm tròn 1 chữ số thập phân để tránh float precision issues
        return str(round(f, 4)).rstrip('0').rstrip('.')
    except ValueError:
        pass

    return val.upper()


def answers_match(a1: str, a2: str) -> bool:
    """
    So sánh 2 đáp án Toán đã được normalize.
    Đặc biệt: số thập phân compare as float nếu có thể.
    """
    n1 = normalize_math_answer(a1)
    n2 = normalize_math_answer(a2)

    if n1 == n2:
        return True

    # Thử so sánh float
    try:
        return abs(float(n1) - float(n2)) < 1e-6
    except (ValueError, TypeError):
        pass

    # Đúng/Sai: so sánh sau khi chuẩn hóa Đ→D
    def ds_norm(s):
        return s.upper().replace('Đ', 'D').replace('đ', 'D')

    return ds_norm(n1) == ds_norm(n2)


# Removed parse_math_answer_key, pass raw dict directly



# ─────────────────────────────────────────────
# 6. AI PROMPT BUILDER (MATH-AWARE)
# ─────────────────────────────────────────────

def build_math_matching_prompt(
    original_questions: list[MathQuestion],
    shuffled_questions: list[MathQuestion],
    exam_code: str,
    has_wmf: bool = False,
    subject: str = 'math',
) -> str:
    """
    Tạo prompt AI để match câu hỏi Toán.
    Nếu có prompt_loader + file prompts/prompt_toan.txt → dùng nội dung từ file.
    Nếu không → dùng prompt hardcode fallback.
    """

    def format_question(q: MathQuestion) -> str:
        prefix = "[⚠️ CÓ CÔNG THỨC - xem file đính kèm] " if q.has_equation else ""
        if q.q_type == 'mc':
            opts = ' | '.join(f"{k}. {v[:60]}" for k, v in sorted(q.options.items()))
            return f"Câu {q.global_number}: {prefix}{q.text[:200]}\n  {opts}"
        elif q.q_type == 'tf':
            sub_items = ' | '.join(f"{k}) {v[:60]}" for k, v in sorted(q.options.items()))
            return f"Câu {q.global_number} [ĐS]: {prefix}{q.text[:200]}\n  {sub_items}"
        elif q.q_type == 'fill':
            return f"Câu {q.global_number} [Điền]: {prefix}{q.text[:200]}"
        return f"Câu {q.global_number}: {q.text[:200]}"

    def group_by_part(qs):
        groups = {}
        for q in qs:
            groups.setdefault(q.part, []).append(q)
        return groups

    orig_groups = group_by_part(original_questions)
    shuf_groups = group_by_part(shuffled_questions)

    # ── Phần đầu: instruction từ file .txt hoặc fallback hardcode ──
    if PROMPT_LOADER_AVAILABLE:
        header = build_prompt_header(subject, exam_code, has_wmf=has_wmf)
        part_labels = get_part_labels(subject) or {
            1: "PHẦN I — Trắc nghiệm nhiều phương án (đáp án: A/B/C/D)",
            2: "PHẦN II — Trắc nghiệm Đúng/Sai (đáp án: chuỗi Đ/S)",
            3: "PHẦN III — Điền số ngắn (đáp án: số thực)",
        }
        output_fmt = get_output_format(subject)
    else:
        wmf_note = ""
        if has_wmf:
            wmf_note = ("\n⚠️ File đề gốc dùng WMF — công thức có thể bị thiếu ([EQ])."
                        " Tham khảo file PDF đính kèm.\n")
        header = (
            f"Bạn là chuyên gia kiểm tra đề thi Toán. Nhiệm vụ: khớp câu hỏi "
            f"ĐỀ TRỘN (mã {exam_code}) với ĐỀ GỐC.{wmf_note}\n"
            "- Câu hỏi ĐÃ BỊ ĐẢO THỨ TỰ trong cùng một PHẦN\n"
            "- Phương án A/B/C/D của câu MC CŨNG ĐÃ BỊ ĐẢO\n"
            "- PHẦN II: ý a/b/c/d thường KHÔNG bị đảo\n"
            "- PHẦN III: chỉ xáo thứ tự câu, KHÔNG xáo đáp án\n"
            "🛑 LƯU Ý TỐI QUAN TRỌNG: NẾU THẤY 'PHẦN IV. Tự luận', HÃY BỎ QUA HOÀN TOÀN, KHÔNG SO SÁNH HAY TRÍCH XUẤT NỘI DUNG PHẦN NÀY.\n"
        )
        part_labels = {
            1: "PHẦN I — Trắc nghiệm nhiều phương án (đáp án: A/B/C/D)",
            2: "PHẦN II — Trắc nghiệm Đúng/Sai (đáp án: chuỗi Đ/S)",
            3: "PHẦN III — Điền số ngắn (đáp án: số thực)",
        }
        output_fmt = (
            "\n=== YÊU CẦU ===\n"
            "Trả về JSON array, mỗi phần tử: "
            '{"shuffled_q": N, "original_q": N, "part": N, "q_type": "mc"|"tf"|"fill", "option_mapping": {...}}.\n'
            "CHỈ TRẢ VỀ JSON ARRAY THUẦN TÚY. KHÔNG THÊM TEXT NÀO KHÁC.\n"
        )

    # ── Nội dung câu hỏi ──
    prompt = header + "\n\n=== ĐỀ GỐC ===\n"
    for part_num in sorted(orig_groups.keys()):
        label = part_labels.get(part_num, f"PHẦN {part_num}")
        prompt += f"\n{label}:\n"
        for q in orig_groups[part_num]:
            prompt += format_question(q) + "\n"

    prompt += f"\n=== ĐỀ TRỘN (Mã {exam_code}) ===\n"
    for part_num in sorted(shuf_groups.keys()):
        label = part_labels.get(part_num, f"PHẦN {part_num}")
        prompt += f"\n{label}:\n"
        for q in shuf_groups[part_num]:
            prompt += format_question(q) + "\n"

    prompt += output_fmt
    return prompt


def build_vision_matching_prompt(
    exam_code: str,
    subject: str = 'math',
    original_name: str = '',
    shuffled_name: str = '',
) -> str:
    """
    Tạo prompt nhẹ cho vision mode: AI đọc file đính kèm trực tiếp.
    Không gửi extracted text — dùng khi file có công thức phức tạp.
    """
    if PROMPT_LOADER_AVAILABLE:
        prompt = get_vision_prompt(subject, exam_code)
        if prompt:
            if original_name:
                prompt += f"\n\nFile 1 (ĐỀ GỐC): {original_name}"
            if shuffled_name:
                prompt += f"\nFile 2 (ĐỀ TRỘN mã {exam_code}): {shuffled_name}"
            return prompt

    # Fallback hardcode
    return (
        f"Bạn là chuyên gia kiểm tra đề thi Toán. "
        f"Tôi đính kèm 2 file: ĐỀ GỐC và ĐỀ TRỘN (mã {exam_code}).\n\n"
        f"NHIỆM VỤ: Đọc cả 2 file, khớp từng câu hỏi.\n\n"
        f"Đề Toán CHỈ KIỂM TRA 3 PHẦN SAU ĐÂY:\n"
        f"1. PHẦN I — Trắc nghiệm (MC): đảo câu + đảo A/B/C/D\n"
        f"2. PHẦN II — Đúng/Sai (TF): đảo câu, ý a/b/c/d thường không đảo\n"
        f"3. PHẦN III — Điền số (Fill): đảo câu, đáp án số không đảo\n\n"
        f"🛑 TUYỆT ĐỐI BỎ QUA 'PHẦN IV. Tự luận' (nếu có). KHÔNG trả về bất kỳ kết quả nào thuộc Phần IV.\n\n"
        f"Trả về JSON array:\n"
        '{"shuffled_q": N, "original_q": N, "part": N, "q_type": "mc"|"tf"|"fill", "option_mapping": {...}}\n\n'
        f"MC: option_mapping bắt buộc. TF: mapping ý nếu đảo. Fill: để {{}}.\n"
        f"CHỈ TRẢ VỀ JSON ARRAY.\n"
        f"\nFile 1 (ĐỀ GỐC): {original_name}\n"
        f"File 2 (ĐỀ TRỘN mã {exam_code}): {shuffled_name}\n"
    )

# ─────────────────────────────────────────────
# 7. ANSWER VERIFIER (MATH-AWARE)
# ─────────────────────────────────────────────

def verify_math_answers(
    ai_matching: list[dict],
    answer_keys: dict,   # đã normalize qua parse_math_answer_key()
    exam_code: str,
) -> list[dict]:
    """
    Xác minh đáp án Toán dựa trên kết quả matching từ AI.
    Xử lý đúng cả 3 loại: MC, Đúng/Sai, Điền số.

    Returns:
        list[dict] — danh sách lỗi, format giống verify_answers() gốc
    """
    errors = []
    original_answers = answer_keys.get('gốc', {})
    shuffled_answers = answer_keys.get(exam_code, {})

    if not original_answers:
        print(f"  ⚠️  Không có đáp án đề gốc!")
        return errors
    if not shuffled_answers:
        print(f"  ⚠️  Không có đáp án mã {exam_code}!")
        return errors

    for item in ai_matching:
        shuffled_q = item['shuffled_q']
        original_q = item['original_q']
        q_type = item.get('q_type', 'mc')
        option_mapping = item.get('option_mapping', {})

        correct_orig = original_answers.get(original_q, '')
        if not correct_orig:
            print(f"  ⚠️  Không có đáp án gốc câu {original_q}")
            continue

        current = shuffled_answers.get(shuffled_q, '')
        if not current:
            print(f"  ⚠️  Không có đáp án mã {exam_code} câu {shuffled_q}")
            continue

        # Tính đáp án đúng trong đề trộn dựa theo loại câu
        correct_shuffled = _compute_correct_shuffled_answer(
            correct_orig, q_type, option_mapping
        )

        if not correct_shuffled:
            print(f"  ⚠️  Không thể tính đáp án đúng cho câu {shuffled_q} ({q_type})")
            continue

        # Debug: hiển thị chi tiết matching
        print(f"  🔍 Câu {shuffled_q} ({q_type}): gốc={original_q} "
              f"| đáp án gốc={correct_orig} → đáp án trộn={correct_shuffled} "
              f"| hiện tại={current} | mapping={option_mapping}")

        # So sánh
        if not answers_match(current, correct_shuffled):
            errors.append({
                'exam_code': f"Mã {exam_code}",
                'shuffled_q': shuffled_q,
                'current_answer': current,
                'correct_answer': correct_shuffled,
                'original_q': original_q,
                'q_type': q_type,
            })
            print(f"  ❌ SAI: Mã {exam_code} Câu {shuffled_q} ({q_type})"
                  f" [gốc: câu {original_q}]"
                  f" | Hiện tại: {current} → Đúng: {correct_shuffled}")

    return errors


def _compute_correct_shuffled_answer(
    correct_orig: str,
    q_type: str,
    option_mapping: dict,
) -> str:
    """
    Tính đáp án đúng trong đề trộn từ đáp án gốc + option_mapping.

    option_mapping: keys là chữ cái ĐỀ TRỘN, values là chữ cái ĐỀ GỐC
    Ví dụ: {'A': 'C', 'B': 'A', 'C': 'D', 'D': 'B'}
    → Đáp án gốc là 'C' → Đáp án trộn là 'A'
    """
    if q_type == 'fill':
        # Điền số không bị đảo đáp án
        return correct_orig

    if q_type == 'mc':
        # Tìm key đề trộn mà value (đề gốc) = correct_orig
        for shuffled_letter, orig_letter in option_mapping.items():
            if orig_letter.upper() == correct_orig.upper():
                return shuffled_letter.upper()
        # Fallback nếu không có mapping
        return correct_orig

    if q_type == 'tf':
        # Đáp án dạng ĐSSS, SĐSD...
        # option_mapping: {'a': 'a', 'b': 'c', ...} (ý trộn → ý gốc)
        if not option_mapping:
            # Không bị đảo ý → giữ nguyên
            return correct_orig

        # Chuẩn hóa đáp án gốc thành dict ý → Đ/S
        ds_str = correct_orig.upper()
        orig_sub_answers = {}
        for idx, letter in enumerate('ABCD'):
            if idx < len(ds_str):
                orig_sub_answers[letter.lower()] = ds_str[idx]
        # Xử lý cả ký hiệu a/b/c/d thường
        for idx, letter in enumerate('abcd'):
            if idx < len(ds_str):
                orig_sub_answers[letter] = ds_str[idx]

        # Tạo chuỗi đáp án trộn
        result_chars = []
        for shuf_sub in ['a', 'b', 'c', 'd']:
            orig_sub = option_mapping.get(shuf_sub, shuf_sub)
            result_chars.append(orig_sub_answers.get(orig_sub, '?'))

        return ''.join(result_chars)

    return correct_orig



# ─────────────────────────────────────────────
# 8. PDF-FALLBACK HELPER (khi docx có WMF)
# ─────────────────────────────────────────────

def get_pdf_companion(docx_path: str) -> Optional[str]:
    """
    Tìm file PDF cùng tên với docx (để dùng làm fallback visual).
    Ví dụ: Đề_gốc.docx → Đề_gốc.pdf

    Returns: đường dẫn PDF nếu tồn tại, None nếu không
    """
    pdf_path = os.path.splitext(docx_path)[0] + '.pdf'
    if os.path.exists(pdf_path):
        return pdf_path
    # Cũng tìm trong cùng thư mục
    dirname = os.path.dirname(docx_path)
    basename = os.path.splitext(os.path.basename(docx_path))[0]
    for fname in os.listdir(dirname):
        if fname.lower().endswith('.pdf') and basename.lower() in fname.lower():
            return os.path.join(dirname, fname)
    return None


def prepare_files_for_ai(
    original_file: str,
    shuffled_file: str,
    has_wmf: bool = False,
) -> list[str]:
    """
    Chuẩn bị danh sách file để gửi kèm AI.

    Logic theo từng trường hợp:

    ┌─────────────────────────┬──────────────────────────────────────────┐
    │ Đề gốc                  │ Gửi AI                                   │
    ├─────────────────────────┼──────────────────────────────────────────┤
    │ .pdf                    │ PDF trực tiếp (AI đọc được hình ảnh)     │
    │ .docx + OMML (mới)      │ docx (AI đọc OMML qua inline)            │
    │ .docx + WMF (Eq. 3.0)   │ PDF companion (cùng tên) nếu có, else   │
    │                         │ docx (AI chỉ thấy text, thiếu công thức)│
    └─────────────────────────┴──────────────────────────────────────────┘

    Đề trộn: ưu tiên PDF (nếu có cả pdf và docx cùng mã đề), fallback docx.
    """
    files = []
    orig_ext = os.path.splitext(original_file)[1].lower()

    # ── Đề gốc ──
    if orig_ext == '.pdf':
        # Đề gốc đã là PDF → gửi thẳng, AI nhìn được công thức
        files.append(original_file)
        print(f"  📎 Đề gốc PDF: {os.path.basename(original_file)}")

    elif has_wmf:
        # Docx dùng Equation 3.0 → cần PDF companion
        pdf = get_pdf_companion(original_file)
        if pdf:
            print(f"  📎 Đề gốc WMF → dùng PDF companion: {os.path.basename(pdf)}")
            files.append(pdf)
        else:
            print(f"  ⚠️  Đề gốc WMF nhưng KHÔNG tìm thấy PDF companion!")
            print(f"      → Export '{os.path.basename(original_file)}' sang PDF (cùng tên) để tăng độ chính xác.")
            files.append(original_file)  # gửi docx, AI chỉ thấy text stub
    else:
        # Docx dùng OMML mới → gửi thẳng docx
        files.append(original_file)
        print(f"  📎 Đề gốc OMML docx: {os.path.basename(original_file)}")

    # ── Đề trộn ──
    if os.path.exists(shuffled_file):
        shuf_ext = os.path.splitext(shuffled_file)[1].lower()

        # Nếu đề trộn là docx, kiểm tra xem có PDF cùng tên không (ưu tiên PDF)
        if shuf_ext == '.docx':
            pdf_alt = get_pdf_companion(shuffled_file)
            if pdf_alt:
                print(f"  📎 Đề trộn → dùng PDF: {os.path.basename(pdf_alt)}")
                files.append(pdf_alt)
            else:
                files.append(shuffled_file)
        else:
            # Đã là PDF
            files.append(shuffled_file)

    return files


# ─────────────────────────────────────────────
# 9. INTEGRATION HELPER
# ─────────────────────────────────────────────

def process_math_exam_v2(
    original_file: str,
    shuffled_file: str,
    answer_keys_raw: dict,
    exam_code: str,
    ai_client,
    extract_text_fn,
    use_vision: bool = False,
):
    import json
    import time
    import re
    import os

    # Dùng nguyên dict đáp án (đã .upper() từ parse_answer_key_excel)
    answer_keys = answer_keys_raw

    orig_ext = os.path.splitext(original_file)[1].lower()
    has_wmf = False
    if orig_ext == '.docx':
        has_wmf = docx_has_wmf_equations(original_file)
        if has_wmf:
            print(f"  ⚠️  Đề gốc dùng Equation 3.0 (WMF)")

    # Detect subject
    if orig_ext == '.pdf':
        subject = detect_subject(original_file, lines=None)
    else:
        orig_lines = extract_text_fn(original_file)
        subject = detect_subject(original_file, orig_lines)

    print(f"  🎯 Môn: {subject} | WMF: {has_wmf} | Vision: {use_vision}")

    should_use_vision = use_vision or has_wmf or (orig_ext == '.pdf')
    ai_files = prepare_files_for_ai(original_file, shuffled_file, has_wmf)

    # Parse text để lấy danh sách câu hỏi
    orig_lines = extract_text_fn(original_file)
    shuf_lines = extract_text_fn(shuffled_file)
    orig_text = '\n'.join(orig_lines)
    shuf_text = '\n'.join(shuf_lines)

    if subject == 'math':
        orig_qs = parse_math_exam_from_text(orig_text, os.path.basename(original_file), subject=subject)
        shuf_qs = parse_math_exam_from_text(shuf_text, os.path.basename(shuffled_file), subject=subject)
    else:
        from check_de import parse_questions_from_text as legacy_parse
        orig_qs_legacy = legacy_parse(orig_text, os.path.basename(original_file))
        shuf_qs_legacy = legacy_parse(shuf_text, os.path.basename(shuffled_file))
        orig_qs = [MathQuestion(
            number=q.number, global_number=q.number, part=1, q_type='mc',
            text=q.text, options=q.options,
        ) for q in orig_qs_legacy]
        shuf_qs = [MathQuestion(
            number=q.number, global_number=q.number, part=1, q_type='mc',
            text=q.text, options=q.options,
        ) for q in shuf_qs_legacy]

    if should_use_vision:
        print(f"  👁️  Vision mode: gửi file trực tiếp cho AI")
        prompt = build_vision_matching_prompt(
            exam_code=exam_code,
            subject=subject,
            original_name=os.path.basename(original_file),
            shuffled_name=os.path.basename(shuffled_file),
        )
    else:
        prompt = build_math_matching_prompt(
            orig_qs, shuf_qs, exam_code,
            has_wmf=has_wmf, subject=subject,
        )
        if has_wmf:
            prompt += "\n\nHÃY THAM KHẢO FILE PDF ĐÍNH KÈM để đọc công thức.\n"


    # ─────────────────────────────────────────────
    # CALL AI - BẬT CACHE FILE + PARSER JSON BẤT TỬ
    # ─────────────────────────────────────────────
    max_retries = 3
    ai_matching = []

    for attempt in range(1, max_retries + 1):
        try:
            if attempt > 1:
                wait_time = 20 * attempt # Nghỉ 40s, 60s để nhả Rate limit
                print(f"  ⏳ Hệ thống tạm nghỉ {wait_time}s để API hồi phục...")
                time.sleep(wait_time)

            print(f"  🤖 Gọi AI matching cho mã {exam_code}... (Lần {attempt}, prompt: {len(prompt):,} chars, files: {len(ai_files)})")
            
            # CHÚ Ý: Đã bật use_file_api=True để cache file PDF, chống bị đánh dấu là Spam!
            response = ai_client.send_data_to_AI(
                prompt=prompt,
                file_paths=ai_files,
                temperature=0.05,
                top_p=0.8,
                max_output_tokens=65000,
                response_mime_type="application/json",
                use_file_api=True 
            )

            # 1. Bắt trường hợp bị Google chặn Safety
            if "⚠️ API trả về rỗng" in response:
                raise ValueError("API bị Google đánh dấu Spam/Safety. AI không sinh ra kết quả.")

            # 2. Làm sạch chuỗi
            cleaned = response.strip()
            if cleaned.startswith('```'):
                cleaned = re.sub(r'^```\w*\n?', '', cleaned)
                cleaned = re.sub(r'\n?```$', '', cleaned)
                cleaned = cleaned.strip()

            ai_matching = None
            
            # 3. Thử parse trực tiếp (cứu các case AI trả về Object thay vì Array)
            try:
                parsed_data = json.loads(cleaned)
                if isinstance(parsed_data, dict):
                    # Tìm mảng nằm trong object (vd: {"matches": [...]})
                    for key, val in parsed_data.items():
                        if isinstance(val, list):
                            ai_matching = val
                            break
                    if ai_matching is None:
                        ai_matching = [parsed_data]
                elif isinstance(parsed_data, list):
                    ai_matching = parsed_data
            except json.JSONDecodeError:
                # 4. Nếu dính text linh tinh, xài Regex vét cạn
                match = re.search(r'\[.*\]', cleaned, re.DOTALL)
                if match:
                    try:
                        ai_matching = json.loads(match.group(0))
                    except json.JSONDecodeError:
                        raise ValueError(f"Regex tìm được mảng nhưng bị lỗi parse. Nội dung AI trả: {cleaned[:150]}...")
                else:
                    # NẾU LỖI, IN THẲNG CHUỖI RA LOG ĐỂ BẮT TẬN TAY!
                    raise ValueError(f"Không tìm thấy mảng JSON. Nội dung AI trả về thực sự là: {cleaned[:150]}...")

            # 5. Ép KPI chất lượng
            if not isinstance(ai_matching, list):
                raise TypeError("Dữ liệu cuối cùng không phải là mảng.")
            
            if len(ai_matching) < len(shuf_qs) * 0.90:
                raise ValueError(f"AI trả thiếu câu ({len(ai_matching)}/{len(shuf_qs)}).")

            mc_questions = [q for q in ai_matching if q.get('q_type', 'mc') == 'mc']
            for q in mc_questions:
                if 'option_mapping' not in q or not q['option_mapping']:
                    raise ValueError(f"AI quên map A,B,C,D ở câu số {q.get('shuffled_q')}.")

            print(f"  ✅ AI đã làm chuẩn {len(ai_matching)} câu.")
            break

        except Exception as e:
            print(f"  ⚠️ Lỗi AI ở lần {attempt}: {e}")
            ai_matching = []
            
            if attempt == max_retries:
                print(f"  ❌ Thất bại sau {max_retries} lần. Bỏ qua mã {exam_code}.")
                return [], [], {}, orig_qs, shuf_qs
    # ─────────────────────────────────────────────

    # Map lại thứ tự global cho các câu (bản vá cũ của bạn)
    if shuf_qs and orig_qs and ai_matching:
        unique_matches = {}
        for item in ai_matching:
            if 'part' not in item:
                qt = item.get('q_type', 'mc').lower()
                if qt == 'tf': item['part'] = 2
                elif qt == 'fill': item['part'] = 3
                else: item['part'] = 1
                
            part = item.get('part', 1)
            sq = item.get('shuffled_q', 0)
            oq = item.get('original_q', 0)
            
            mapped_sq = sq
            for q in shuf_qs:
                if getattr(q, 'part', 1) == part and getattr(q, 'number', 0) == sq:
                    mapped_sq = getattr(q, 'global_number', sq)
                    break
            item['shuffled_q'] = mapped_sq
            
            mapped_oq = oq
            for q in orig_qs:
                if getattr(q, 'part', 1) == part and getattr(q, 'number', 0) == oq:
                    mapped_oq = getattr(q, 'global_number', oq)
                    break
            item['original_q'] = mapped_oq
            
            if mapped_sq not in unique_matches:
                unique_matches[mapped_sq] = item
                
        ai_matching = list(unique_matches.values())

    errors = verify_math_answers(ai_matching, answer_keys, exam_code)

    debug = {
        'exam_code': exam_code,
        'subject': subject,
        'has_wmf': has_wmf,
        'vision_mode': should_use_vision,
        'ai_matches': len(ai_matching),
        'errors': errors,
    }

    return errors, ai_matching, debug, orig_qs, shuf_qs

