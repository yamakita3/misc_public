#!/usr/bin/env python3
from pathlib import Path
import argparse
import re
import sys
from typing import Optional, Tuple

from PIL import Image, ImageOps, ImageFilter
import pytesseract

TS_PATTERN = re.compile(
    r'(\d{1,2})\s*[Cc℃]\s*(\d{1,3})\s*[Ff℉°]?\s*'
    r'(\d{4})\D?(\d{2})\D?(\d{2})\s+'
    r'(\d{2})\D?(\d{2})\D?(\d{2})\s+(\d{4})'
)


def natural_key(path: Path):
    return [int(s) if s.isdigit() else s.lower() for s in re.split(r'(\d+)', path.name)]



def crop_bottom_band(img: Image.Image, band_ratio: float = 0.14) -> Image.Image:
    w, h = img.size
    top = int(h * (1 - band_ratio))
    return img.crop((0, top, w, h))



def preprocess(img: Image.Image, scale: int = 3) -> Image.Image:
    gray = ImageOps.grayscale(img)
    w, h = gray.size
    gray = gray.resize((w * scale, h * scale))
    gray = gray.filter(ImageFilter.SHARPEN)
    gray = ImageOps.autocontrast(gray)
    bw = gray.point(lambda p: 255 if p > 160 else 0)
    return bw



def ocr_text(img: Image.Image) -> str:
    configs = [
        '--psm 6',
        '--psm 7',
        '--psm 11',
    ]
    texts = []
    for cfg in configs:
        try:
            txt = pytesseract.image_to_string(img, lang='eng', config=cfg)
            texts.append(txt)
        except Exception:
            pass
    return '\n'.join(texts)



def normalize(text: str) -> str:
    t = text
    t = t.replace('O', '0').replace('o', '0')
    t = t.replace('I', '1').replace('l', '1').replace('|', '1')
    t = t.replace('S', '5')
    t = t.replace('B', '8')
    t = t.replace('“', ' ').replace('”', ' ')
    t = t.replace('—', '-').replace('_', ' ')
    t = re.sub(r'\s+', ' ', t)
    return t.strip()



def extract_structured(text: str) -> Optional[str]:
    t = normalize(text)
    m = TS_PATTERN.search(t)
    if not m:
        digits = re.sub(r'[^0-9]', '', t)
        if len(digits) >= 21:
            digits = digits[-21:]
            c, f = digits[0:2], digits[2:4]
            yyyy, mm, dd = digits[4:8], digits[8:10], digits[10:12]
            hh, mi, ss = digits[12:14], digits[14:16], digits[16:18]
            idx = digits[18:22] if len(digits) >= 22 else digits[18:21]
            return f'{c}℃{f}℉{yyyy}/{mm}/{dd} {hh}:{mi}:{ss} {idx}'
        return None
    c, f, yyyy, mm, dd, hh, mi, ss, idx = m.groups()
    return f'{int(c)}℃{int(f)}℉{yyyy}/{mm}/{dd} {hh}:{mi}:{ss} {idx}'



def process_image(path: Path) -> Tuple[str, Optional[str]]:
    img = Image.open(path)
    band = crop_bottom_band(img)
    prep = preprocess(band)
    text = ocr_text(prep)
    parsed = extract_structured(text)
    return path.name, parsed



def main():
    parser = argparse.ArgumentParser(description='Trail camera style timestamp OCR batch extractor')
    parser.add_argument('input_dir', nargs='?', default='.', help='画像フォルダ')
    parser.add_argument('-o', '--output', default='ocr_results.txt', help='追記先テキスト')
    parser.add_argument('--glob', default='*.jpg', help='画像glob。例: *.jpg')
    args = parser.parse_args()

    input_dir = Path(args.input_dir)
    files = sorted(input_dir.glob(args.glob), key=natural_key)
    if not files:
        exts = ('*.jpg', '*.jpeg', '*.png', '*.JPG', '*.JPEG', '*.PNG')
        files = []
        for ext in exts:
            files.extend(input_dir.glob(ext))
        files = sorted(set(files), key=natural_key)

    if not files:
        print('画像が見つかりませんでした。', file=sys.stderr)
        sys.exit(1)

    with open(args.output, 'a', encoding='utf-8') as fw:
        for path in files:
            name, parsed = process_image(path)
            line = f'{name}\t{parsed if parsed else "OCR失敗"}'
            fw.write(line + '\n')
            print(line)


if __name__ == '__main__':
    main()
