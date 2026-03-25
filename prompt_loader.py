# -*- coding: utf-8 -*-
"""
prompt_loader.py
================
Module load prompt từ file .txt theo subject (math / english / other).
Cung cấp các hàm helper cho check_de.py và math_exam_handler.py.
"""

import os
import re

# Thư mục chứa file prompt
_PROMPTS_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'prompts')

# Cache: subject → parsed dict
_cache: dict = {}


def _parse_prompt_file(filepath: str) -> dict:
    """
    Parse file prompt .txt theo format [SECTION_NAME] ... nội dung ...
    Returns: dict { section_name_lower: content_str }
    """
    sections = {}
    current_section = None
    current_lines = []

    with open(filepath, 'r', encoding='utf-8') as f:
        for line in f:
            # Detect section header: [SECTION_NAME]
            m = re.match(r'^\[([A-Z_]+)\]\s*$', line.strip())
            if m:
                # Save previous section
                if current_section is not None:
                    sections[current_section] = '\n'.join(current_lines).strip()
                current_section = m.group(1).lower()
                current_lines = []
            elif current_section is not None:
                # Bỏ qua dòng comment (#) ở đầu section
                if not (line.strip().startswith('#') and current_section in ('parts_config',)):
                    current_lines.append(line.rstrip('\n'))

    # Save last section
    if current_section is not None:
        sections[current_section] = '\n'.join(current_lines).strip()

    return sections


def _load_subject(subject: str) -> dict:
    """Load và cache prompt file cho subject."""
    if subject in _cache:
        return _cache[subject]

    # Tìm file prompt
    filename = f'prompt_{subject}.txt'
    filepath = os.path.join(_PROMPTS_DIR, filename)

    if not os.path.exists(filepath):
        print(f"  ⚠️  Không tìm thấy file prompt: {filepath}")
        _cache[subject] = {}
        return {}

    try:
        data = _parse_prompt_file(filepath)
        _cache[subject] = data
        return data
    except Exception as e:
        print(f"  ❌ Lỗi đọc file prompt {filepath}: {e}")
        _cache[subject] = {}
        return {}


# ============================================================
# PUBLIC API
# ============================================================

def build_prompt_header(subject: str, exam_code: str, has_wmf: bool = False) -> str:
    """Trả về phần instruction header cho prompt AI."""
    data = _load_subject(subject)

    header = data.get('prompt_header', '')
    if not header:
        return ''

    # Format exam_code vào prompt
    header = header.replace('{exam_code}', str(exam_code))

    # Thêm WMF warning nếu cần
    if has_wmf:
        wmf_note = data.get('prompt_header_wmf', '')
        if wmf_note:
            header += '\n' + wmf_note

    return header


def get_output_format(subject: str) -> str:
    """Trả về phần yêu cầu output JSON."""
    data = _load_subject(subject)
    return data.get('output_format', '')


def get_vision_prompt(subject: str, exam_code: str) -> str:
    """Trả về prompt cho vision mode (gửi file trực tiếp)."""
    data = _load_subject(subject)
    prompt = data.get('vision_prompt', '')
    if prompt:
        prompt = prompt.replace('{exam_code}', str(exam_code))
    return prompt


def get_part_labels(subject: str) -> dict:
    """
    Trả về dict {part_num: label_str}.
    Ví dụ: {1: 'PHẦN I — Trắc nghiệm ...', 2: 'PHẦN II — Đúng/Sai ...'}
    """
    data = _load_subject(subject)
    raw = data.get('part_labels', '')
    if not raw:
        return {}

    labels = {}
    for line in raw.strip().split('\n'):
        line = line.strip()
        if '=' in line:
            key, val = line.split('=', 1)
            try:
                labels[int(key.strip())] = val.strip()
            except ValueError:
                pass
    return labels


def get_max_pages(subject: str) -> int:
    """Trả về số trang tối đa cho phép."""
    data = _load_subject(subject)
    raw = data.get('max_pages', '4')
    try:
        return int(raw.strip())
    except ValueError:
        return 4


def get_subject_label(subject: str) -> str:
    """Trả về tên hiển thị của môn (ví dụ: 'Toán học', 'Tiếng Anh')."""
    data = _load_subject(subject)
    return data.get('subject_label', subject.capitalize())


def get_parts_config(subject: str) -> list[dict]:
    """
    Trả về cấu hình các PHẦN.
    Returns: [{'part': 1, 'q_type': 'mc', 'label': '...', 'keywords': ['...']}, ...]
    """
    data = _load_subject(subject)
    raw = data.get('parts_config', '')
    if not raw:
        return []

    configs = []
    for line in raw.strip().split('\n'):
        line = line.strip()
        if not line or line.startswith('#'):
            continue
        parts = line.split('|')
        if len(parts) >= 3:
            try:
                cfg = {
                    'part': int(parts[0].strip()),
                    'q_type': parts[1].strip(),
                    'label': parts[2].strip(),
                    'keywords': [kw.strip() for kw in parts[3].split(',')] if len(parts) > 3 else [],
                }
                configs.append(cfg)
            except (ValueError, IndexError):
                pass
    return configs


def get_part_boundary_patterns(subject: str) -> dict:
    """
    Trả về dict {part_num: [keyword_list]} để nhận diện tiêu đề phần trong văn bản.
    """
    configs = get_parts_config(subject)
    if not configs:
        return {}

    return {cfg['part']: cfg['keywords'] for cfg in configs if cfg['keywords']}


def get_part_type(subject: str, part_num: int) -> str:
    """Trả về q_type cho part_num. Mặc định 'mc'."""
    configs = get_parts_config(subject)
    for cfg in configs:
        if cfg['part'] == part_num:
            return cfg['q_type']
    return 'mc'
