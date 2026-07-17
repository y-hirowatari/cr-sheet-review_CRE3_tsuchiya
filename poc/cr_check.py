#!/usr/bin/env python3
"""
ネイルホリック CRシート 整合チェック PoC (v0)

提出物（クリエイティブ画像）が無くても、CRシート単体で検証できる範囲の
「機械的な照合」を実行する。設計方針は docs/AIレビューの仕組み_設計書_チーム共有用.md、
仕様は docs/照合ロジック仕様_たたき台.md を参照。

v0 でやること（提出物なしで検証可能）:
  1. 色番号↔色ライン名の整合チェック（シート全体）
       同じ色番号が別のライン名で書かれていないか＝取り違え/誤記の検出
  2. キャプション使用色 ⊆（文字入れ指示 ∪ 対象商材）の包含チェック（投稿単位）
       キャプションにあるのにシートのどこにも無い色番号＝キャプション側の取り違え
  3. インベントリ出力（投稿一覧・抽出色番号・対象商材カバレッジ）
  --demo-typo: 正解に1文字だけ違う色番号を混ぜ、機械照合が検出できることを実演

v0 でやらないこと:
  - 提出画像のOCR照合（v1。実際の提出画像が必要）
  - 配色/アート出典の判定（参考どまり）

使い方:
  python3 cr_check.py <CRシート.pdf> [<CRシート.pdf> ...] [--demo-typo]
"""
import sys
import re
import unicodedata
from collections import defaultdict

CODE_RE = re.compile(r"[A-Z]{2}\d{3}[A-Z]?")
POST_RE = re.compile(r"投稿no,\s*(\d+)_\s*(\d+)月\s*(\d+)日")
TITLE_RE = re.compile(r"投稿no,.*?【(.+?)】")
CODE_LINE_RE = re.compile(r"^([A-Z]{2}\d{3}[A-Z]?)[ 　]+(.+)$")

# ライン名として扱わない語（対象商材欄の定型句など）
NON_LINE = ("商材画像", "格納先", "色玉", "必要素材", "背景画像")


def norm(s: str) -> str:
    """記号・空白・全角半角の表記ゆれのみ吸収（色番号そのものは変えない）。"""
    s = unicodedata.normalize("NFKC", s)
    s = s.replace("　", " ")
    s = re.sub(r"[ \t]+", " ", s)
    return s.strip()


def norm_line(name: str) -> str:
    """ライン名の正規化：囲み記号・前後の記号を落とす。"""
    name = norm(name)
    name = name.strip("()（）「」[]【】 　:：・")
    return name.strip()


def codes_in(text: str):
    return CODE_RE.findall(norm(text))


def parse(pdf_path):
    import fitz
    doc = fitz.open(pdf_path)
    posts = {}
    medium = None

    def get_post(medium, text):
        m = POST_RE.search(norm(text))
        if not m:
            return None
        key = (medium, int(m.group(1)), int(m.group(2)), int(m.group(3)))
        if key not in posts:
            posts[key] = {
                "key": key, "medium": medium, "title": None,
                "instr_codes": set(), "shohin_codes": set(), "caption_codes": set(),
                "code_line": {}, "has_caption": False, "has_instr": False,
            }
        return posts[key]

    for pg in doc:
        raw = pg.get_text()
        n = norm(raw)
        stripped = n.strip()

        # 媒体の切り替え（"IG" / "X" の区切りページ）
        if stripped in ("IG", "X"):
            medium = stripped
            continue
        if medium is None:
            continue

        is_cap = "投稿キャプション" in n
        is_ins = "文字入れ指示内容" in n
        if not (is_cap or is_ins):      # 目次・一覧ページは投稿化しない
            continue

        post = get_post(medium, raw)
        if post is None:
            continue

        tm = TITLE_RE.search(n)
        if tm and not post["title"]:
            post["title"] = tm.group(1)

        if is_cap:
            post["has_caption"] = True
            body = n.split("【修正版】", 1)[1] if "【修正版】" in n else n
            post["caption_codes"] |= set(codes_in(body))
        elif is_ins:
            post["has_instr"] = True
            instr_part = n.split("デザイン", 1)[0] if "デザイン" in n else n
            shohin_part = n.split("対象商材", 1)[1] if "対象商材" in n else ""
            for raw_line in instr_part.splitlines():
                line = norm(raw_line)
                m = CODE_LINE_RE.match(line)
                if m:
                    code, name = m.group(1), norm_line(m.group(2))
                    post["instr_codes"].add(code)
                    if name and not any(w in name for w in NON_LINE):
                        post["code_line"].setdefault(code, set()).add(name)
                else:
                    post["instr_codes"] |= set(codes_in(line))
            post["shohin_codes"] |= set(codes_in(shohin_part))

    return posts


def collect(pdfs):
    all_posts, gmap, gsrc = [], defaultdict(set), defaultdict(lambda: defaultdict(set))
    for pdf in pdfs:
        posts = parse(pdf)
        src = pdf.split("/")[-1]
        # ファイル名から "6月"/"7月" を推定できないので通し番号ラベル
        for key in sorted(posts, key=lambda k: (k[0], k[2], k[3], k[1])):
            p = posts[key]
            p["src"] = src
            all_posts.append(p)
            lbl = f'{p["medium"]} no,{key[1]} {key[2]}/{key[3]}'
            for code, names in p["code_line"].items():
                for nm in names:
                    gmap[code].add(nm)
                    gsrc[code][nm].add(lbl)
    return all_posts, gmap, gsrc


def report(all_posts, gmap, gsrc):
    line = "=" * 72
    print(line)
    print(" ネイルホリック CRシート 整合チェック PoC (v0)")
    print(line)

    # 【1】色番号↔ライン名
    print("\n【1】色番号↔色ライン名の整合チェック（シート全体）")
    conflicts = {c: v for c, v in gmap.items() if len(v) > 1}
    if not conflicts:
        print(f"  OK: 色番号 {len(gmap)} 種すべてでライン名の矛盾なし")
    else:
        print(f"  NG: {len(conflicts)} 件の色番号でライン名が食い違い（取り違えの可能性）")
        for c in sorted(conflicts):
            print(f"    - {c}:")
            for nm in sorted(conflicts[c]):
                print(f"        「{nm}」 ← {', '.join(sorted(gsrc[c][nm]))}")

    # 【2】キャプション ⊆ 指示∪商材
    print("\n【2】キャプション使用色 ⊆（文字入れ指示 ∪ 対象商材）（投稿単位）")
    print("     ※キャプションにあるがシートのどこにも無い色番号＝キャプション側の取り違え疑い")
    issues = 0
    for p in all_posts:
        if not p["has_caption"]:
            continue
        sheet = p["instr_codes"] | p["shohin_codes"]
        if not sheet:                       # 照合対象なし（X の転載キャプション等）
            continue
        missing = p["caption_codes"] - sheet
        if missing:
            issues += 1
            k = p["key"]
            print(f'  NG: {p["medium"]} no,{k[1]} {k[2]}/{k[3]} 【{p["title"]}】')
            print(f"        シートに無い色番号: {', '.join(sorted(missing))}")
    if not issues:
        print("  OK: 照合対象の全投稿で、キャプションの色番号はシート側に存在")

    # 【3】インベントリ
    print("\n【3】投稿インベントリ")
    print(f"  {'媒体':<4}{'投稿':<26}{'指示':>4}{'商材':>4}{'ｷｬﾌﾟ':>5}  商材カバレッジ")
    for p in all_posts:
        k = p["key"]
        lbl = f'no,{k[1]} {k[2]}/{k[3]} {p["title"] or ""}'[:25]
        instr = p["instr_codes"]
        cov = f"{len(instr & p['shohin_codes'])}/{len(instr)}" if instr else "-"
        print(f"  {p['medium']:<4}{lbl:<26}{len(instr):>4}{len(p['shohin_codes']):>4}"
              f"{len(p['caption_codes']):>5}  {cov}")
    print(f"\n  ユニーク色番号（全シート合計）: {len(gmap)} 種")
    print(line)


def demo_typo(all_posts):
    """正解セットの1つを1文字だけ書き換えた『提出物/キャプション』を作り、検出できることを実演。"""
    target = max((p for p in all_posts if p["has_instr"]),
                 key=lambda p: len(p["instr_codes"]), default=None)
    if not target:
        return
    truth = sorted(target["instr_codes"] | target["shohin_codes"])
    victim = truth[0]
    # 末尾の数字を1つずらす（BL920 -> BL921 のような一字違いを模擬）
    m = re.match(r"^([A-Z]{2})(\d{3})([A-Z]?)$", victim)
    d = int(m.group(2))
    typo = f"{m.group(1)}{(d+1)%1000:03d}{m.group(3)}"
    submitted = set(truth) - {victim} | {typo}

    k = target["key"]
    print("\n" + "-" * 72)
    print(" 【デモ】正解を1文字だけ書き換えたら検出できるか")
    print("-" * 72)
    print(f"  対象投稿: {target['medium']} no,{k[1]} {k[2]}/{k[3]} 【{target['title']}】")
    print(f"  模擬ミス: 提出物側で {victim} → {typo} と1文字取り違え")
    missing = submitted - (target["instr_codes"] | target["shohin_codes"])
    extra = (target["instr_codes"] | target["shohin_codes"]) - submitted
    if missing or extra:
        print("  ✅ 検出成功:")
        if missing:
            print(f"       提出物にある不明な色番号: {', '.join(sorted(missing))}（正解に存在しない）")
        if extra:
            print(f"       提出物に欠けている色番号: {', '.join(sorted(extra))}（正解に存在）")
        print(f"       → 正しい表記(コピペ用): {victim}")
    else:
        print("  ❌ 検出できませんでした")
    print("-" * 72)


def main():
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    flags = {a for a in sys.argv[1:] if a.startswith("--")}
    if not args:
        print(__doc__); sys.exit(1)
    all_posts, gmap, gsrc = collect(args)
    report(all_posts, gmap, gsrc)
    if "--demo-typo" in flags:
        demo_typo(all_posts)


if __name__ == "__main__":
    main()
