"""
한글 이름 → 로마자 변환 → 암호화 코드 생성 유틸
"""
import re

# ── 한글 로마자 변환 테이블 (국립국어원 표준) ────────────
_CHOSUNG = [
    'g','kk','n','d','tt','r','m','b','pp',
    's','ss','','j','jj','ch','k','t','p','h'
]
_JUNGSUNG = [
    'a','ae','ya','yae','eo','e','yeo','ye','o',
    'wa','wae','oe','yo','u','wo','we','wi','yu','eu','ui','i'
]
_JONGSUNG = [
    '','g','kk','gs','n','nj','nh','l','lg','lm',
    'lb','ls','lt','lp','lh','m','b','bs','s','ss',
    'ng','j','ch','k','t','p','h'
]


def _romanize_syllable(char: str) -> str:
    """한글 음절 1자 → 로마자"""
    code = ord(char) - 0xAC00
    if code < 0 or code > 11171:
        return char.lower()
    cho  = code // (21 * 28)
    jung = (code % (21 * 28)) // 28
    jong = code % 28
    return _CHOSUNG[cho] + _JUNGSUNG[jung] + _JONGSUNG[jong]


def romanize_korean(text: str) -> str:
    """한글 단어 → 로마자 (예: 홍길동 → honggildong)"""
    result = ''
    for char in str(text):
        if '가' <= char <= '힣':
            result += _romanize_syllable(char)
        elif char.isalpha():
            result += char.lower()
        # 숫자/특수문자는 제외
    return result


def _extract_suffix(name: str):
    """
    이름 끝의 비한글 접미사 분리
    홍길동a  → ('홍길동', 'a')
    홍길동88 → ('홍길동', '88')
    홍길동   → ('홍길동', '')
    """
    name = str(name).strip()
    i = len(name) - 1
    suffix = ''
    while i >= 0 and not ('가' <= name[i] <= '힣'):
        suffix = name[i] + suffix
        i -= 1
    return name[:i + 1], suffix


def _split_name_parts(korean_name: str) -> list:
    """
    한글 이름 → 파트 분리
    '홍 길 동'  → ['홍', '길', '동']
    '홍길동'    → ['홍', '길동']  (첫 글자=성)
    '남궁 길동' → ['남궁', '길동']
    """
    name = korean_name.strip()
    if ' ' in name:
        return [p for p in name.split() if p]
    elif len(name) >= 2:
        return [name[0], name[1:]]
    return [name]


def generate_code_from_korean(name: str) -> str:
    """
    한글 이름 → 암호화 코드
    홍 길 동  → h4g3d411
    홍길동a   → h4g3d411a
    홍길동88  → h4g3d41188
    """
    name = str(name).strip()
    if not name:
        return ''

    korean_part, suffix = _extract_suffix(name)

    if not korean_part:
        # 이미 영문이름인 경우
        return generate_code_from_english(name)

    parts = _split_name_parts(korean_part)
    code = ''
    total_letters = 0

    for part in parts:
        roman = romanize_korean(part)
        if roman:
            code += roman[0].lower() + str(len(roman))
            total_letters += len(roman)

    code += str(total_letters)
    code += suffix.lower()
    return code


def generate_code_from_english(name: str) -> str:
    """
    영문 이름 → 암호화 코드
    Hong Gil Dong → h4g3d411
    """
    words = str(name).strip().split()
    code = ''
    total = 0
    for word in words:
        clean = re.sub(r'[^a-zA-Z]', '', word)
        if clean:
            code += clean[0].lower() + str(len(clean))
            total += len(clean)
    code += str(total)
    return code
