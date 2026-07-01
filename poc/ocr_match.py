#!/usr/bin/env python3
"""
ネイルホリック 提出画像 OCR照合 PoC (v1)

提出クリエイティブ画像をOCRで文字化し、CRシート由来の「正解」色番号セットと
機械照合する。設計方針どおり、AIには正誤を判断させず、OCRは「文字を読む」だけ、
一致判定は機械が文字単位で行う。

やること:
  1. 画像をOCR（tesseract, 日本語＋英語）で文字抽出
  2. 正規化（記号・空白・全角半角。色番号の1文字差は吸収しない）
  3. 色番号の抽出と、正解セットとの集合照合（過不足＝取り違え）
  4. 指定した必須文字列（注釈・CTA等）の有無チェック
  5. 証拠つきレポート

使い方:
  python3 ocr_match.py <画像> --truth BL920,BL921,WT045R [--require "SPF50+,PA++++"]
  python3 ocr_match.py <画像> --truth-file codes.txt

依存: tesseract-ocr(+jpn), pytesseract, pillow
"""
import sys
import re
import argparse
import unicodedata

CODE_RE = re.compile(r"[A-Z]{2}\d{3}[A-Z]?")


def norm(s: str) -> str:
    s = unicodedata.normalize("NFKC", s)
    s = s.replace("　", " ")
    s = re.sub(r"[ \t]+", " ", s)
    return s.strip()


def ocr_passes(image_path):
    """複数パスでOCRし、(全文テキスト, 色番号→{count, conf}) を返す。
    色番号は複数パスでの検出回数と最大確信度を持たせ、OCRノイズ(1パスだけ・低確信)を選別できるようにする。"""
    from PIL import Image, ImageOps
    import pytesseract
    img = Image.open(image_path).convert("RGB")
    w, h = img.size
    scale = 2 if max(w, h) < 1600 else 1
    if scale > 1:
        img = img.resize((w * scale, h * scale), Image.LANCZOS)
    g = ImageOps.autocontrast(ImageOps.grayscale(img))

    variants = [g, img]
    texts, code_stat = [], {}
    n_pass = 0
    for src in variants:
        for psm in (11, 6, 3):
            n_pass += 1
            cfg = f"--oem 1 --psm {psm}"
            texts.append(pytesseract.image_to_string(src, lang="jpn+eng", config=cfg))
            data = pytesseract.image_to_data(src, lang="jpn+eng", config=cfg,
                                             output_type=pytesseract.Output.DICT)
            for word, conf in zip(data["text"], data["conf"]):
                try:
                    c = float(conf)
                except (TypeError, ValueError):
                    c = -1
                for code in CODE_RE.findall(norm(word)):
                    st = code_stat.setdefault(code, {"count": 0, "conf": 0.0})
                    st["count"] += 1
                    st["conf"] = max(st["conf"], c)
    return "\n".join(texts), code_stat, n_pass


def confident_codes(code_stat, min_conf=60.0, min_count=2):
    """確信度・検出回数のどちらかを満たす色番号を『読めた』とみなす。"""
    return {c for c, st in code_stat.items()
            if st["conf"] >= min_conf or st["count"] >= min_count}


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("image")
    ap.add_argument("--truth", default="", help="正解の色番号(カンマ区切り)")
    ap.add_argument("--truth-file", default="", help="正解の色番号(1行1件 or カンマ区切り)")
    ap.add_argument("--require", default="", help="必須文字列(カンマ区切り)")
    ap.add_argument("--show-ocr", action="store_true", help="OCR生テキストを表示")
    args = ap.parse_args()

    truth = set()
    if args.truth:
        truth |= {norm(x) for x in args.truth.split(",") if x.strip()}
    if args.truth_file:
        with open(args.truth_file, encoding="utf-8") as f:
            for line in f:
                truth |= {norm(x) for x in re.split(r"[,\s]+", line) if x.strip()}

    raw, code_stat, n_pass = ocr_passes(args.image)
    norm_text = norm(raw)
    found = confident_codes(code_stat)

    line = "=" * 72
    print(line)
    print(" ネイルホリック 提出画像 OCR照合 PoC (v1)")
    print(line)
    print(f"  画像: {args.image}    OCRパス数: {n_pass}")

    if args.show_ocr:
        print("\n--- OCR生テキスト（正規化前）---")
        print(raw.strip()[:2000])
        print("--- ここまで ---")

    def stat(c):
        st = code_stat.get(c, {"count": 0, "conf": 0})
        return f"{c}(conf{int(st['conf'])}/{st['count']}回)"

    print(f"\n【OCRで読めた色番号(確信)】 {len(found)} 件: "
          f"{', '.join(stat(c) for c in sorted(found)) or '（なし）'}")
    low = set(code_stat) - found
    if low:
        print(f"  （参考）低確信で除外: {', '.join(stat(c) for c in sorted(low))}")

    if truth:
        print("\n【色番号の照合（正解 vs OCR）】")
        missing = truth - found           # 正解にあるがOCRで見つからない（未反映 or OCR読み落とし）
        unknown = found - truth           # OCRにあるが正解に無い（取り違え疑い）
        ok = truth & found
        print(f"  一致: {len(ok)}/{len(truth)}  {', '.join(sorted(ok)) or ''}")
        if missing:
            print(f"  ⚠ 正解にあるがOCR未検出: {', '.join(sorted(missing))}")
            print("     → 未反映の可能性、またはOCRの読み落とし（写真上の細字など）。要目視。")
        if unknown:
            print(f"  ❌ OCRにあるが正解に無い(取り違え疑い): "
                  f"{', '.join(stat(c) for c in sorted(unknown))}")
        if not missing and not unknown:
            print("  ✅ 色番号は正解と完全一致")

    if args.require:
        print("\n【必須文字列の有無】")
        for token in [t for t in args.require.split(",") if t.strip()]:
            nt = norm(token)
            hit = nt in norm_text
            print(f"  {'✅' if hit else '❌'} {token}")

    print(line)


if __name__ == "__main__":
    main()
