"""エアレジ・Airペイから月次のCSVを自動でダウンロードする（試作）。

  python airregi_download.py                    先月分を格納先フォルダに保存する
  python airregi_download.py --test             先月分を「取得テスト」フォルダに保存し、ブラウザの動きを表示する
  python airregi_download.py --month 2026-09    対象月を指定する

ダウンロードするもの:
  日別売上（売上集計）CSV、商品別売上CSV、Airペイ振込明細CSV（対象月と翌月の振込分）
うまくいかなかった画面は「取得ログ」フォルダに画像とHTMLで保存する。
"""
import calendar
import configparser
import datetime as dt
import re
import shutil
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
LOG_DIR = HERE / "取得ログ"
DAILY_URL = "https://airregi.jp/CLP/view/salesList/"
PRODUCT_URL = "https://airregi.jp/CLP/view/salesListByMenu/"
PAYOUT_URL = "https://merchant.airpayment.jp/rpycl/transfer/details/"
CSV_ENC = "cp932"


def load_config():
    cfg = configparser.ConfigParser(interpolation=None)
    cfg.read(HERE / "設定.ini", encoding="utf-8-sig")
    return cfg


def parse_month(argv, today):
    if "--month" in argv:
        y, m = argv[argv.index("--month") + 1].split("-")
        return int(y), int(m)
    last = today.replace(day=1) - dt.timedelta(days=1)
    return last.year, last.month


def next_month(y, m):
    return (y + 1, 1) if m == 12 else (y, m + 1)


def save_trouble(page, name):
    """止まった画面を画像とHTMLで残す（修正するときに送ってもらう用）。"""
    LOG_DIR.mkdir(exist_ok=True)
    try:
        page.screenshot(path=str(LOG_DIR / f"{name}.png"), full_page=True)
        (LOG_DIR / f"{name}.html").write_text(page.content(), encoding="utf-8")
    except Exception:
        pass


def visible(locator):
    return locator.filter(visible=True)


def click_text(page, text, exact=False):
    pattern = re.compile(r"^\s*" + re.escape(text) + r"\s*$") if exact else re.compile(re.escape(text))
    for c in (page.get_by_role("button", name=pattern), page.get_by_role("link", name=pattern),
              page.locator(f"input[value*='{text}']"), page.get_by_text(pattern)):
        c = visible(c)
        if c.count() > 0:
            c.first.click()
            return
    raise RuntimeError(f"「{text}」が画面に見つかりませんでした")


def download(page, text, timeout=30000):
    with page.expect_download(timeout=timeout) as info:
        click_text(page, text)
    return info.value


def set_input_value(page, locator, value):
    """入力欄に値を入れる。読み取り専用の欄にはJavaScriptで入れる。"""
    try:
        locator.fill(value, timeout=3000)
    except Exception:
        locator.evaluate(
            """(el, v) => {
                const setter = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set;
                setter.call(el, v);
                el.dispatchEvent(new Event('input', {bubbles: true}));
                el.dispatchEvent(new Event('change', {bubbles: true}));
            }""", value)
    page.keyboard.press("Escape")  # カレンダーが開いたら閉じる


def select_year_month(page, y, m):
    """年と月のプルダウンがあれば選ぶ。選べたら True。"""
    done_y = done_m = False
    for sel in visible(page.locator("select")).all():
        labels = [o.strip() for o in sel.locator("option").all_inner_texts()]
        values = sel.locator("option").evaluate_all("os => os.map(o => o.value)")
        for want, kind in ((str(y), "y"), (str(m), "m"), (f"{m:02d}", "m")):
            if (kind == "y" and done_y) or (kind == "m" and done_m):
                continue
            for cand in (want, f"{want}年", f"{want}月"):
                if cand in labels or cand in values:
                    if kind == "m" and any(len(v) == 4 for v in values if v.isdigit()):
                        break  # 年のプルダウンに月を入れない
                    try:
                        sel.select_option(label=cand) if cand in labels else sel.select_option(value=cand)
                    except Exception:
                        continue
                    done_y, done_m = (True, done_m) if kind == "y" else (done_y, True)
                    break
    return done_y and done_m


# ---------- 各画面 ----------

def login(page, cfg):
    page.goto(DAILY_URL)
    page.wait_for_load_state("networkidle")
    pw_box = visible(page.locator("input[type=password]"))
    if pw_box.count() == 0:
        return  # 前回のログインが残っている
    a = cfg["エアレジ"]
    visible(page.locator("input:not([type=password]):not([type=hidden]):not([type=checkbox])"
                         ":not([type=radio]):not([type=submit]):not([type=button])")).first.fill(a["ログインID"].strip())
    pw_box.first.fill(a["パスワード"].strip())
    only_login = re.compile(r"^\s*ログイン(する)?\s*$")
    btn = visible(page.get_by_role("button", name=only_login))
    if btn.count() == 0:
        btn = visible(page.locator("input[type=submit], button[type=submit]"))
    btn.last.click() if btn.count() else pw_box.first.press("Enter")
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(2000)
    if visible(page.locator("input[type=password]")).count() > 0:
        raise RuntimeError("ログインできませんでした（IDとパスワード、または確認コードの画面を確認してください）")
    if "salesList" not in page.url:
        page.goto(DAILY_URL)
        page.wait_for_load_state("networkidle")


def get_daily(page, y, m, out_dir):
    page.goto(DAILY_URL)
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(1500)
    if not select_year_month(page, y, m):
        boxes = visible(page.locator("input")).all()
        box = next((b for b in boxes if re.search(r"\d{4}\s*[/年]\s*\d{1,2}", b.input_value() or "")), None)
        if box is None:
            raise RuntimeError("対象月を入れる欄が見つかりませんでした")
        old = box.input_value()
        new = re.sub(r"(\d{4})(\s*[/年]\s*)(\d{1,2})",
                     lambda mt: f"{y}{mt.group(2)}{m:0{len(mt.group(3))}d}", old, count=1)
        set_input_value(page, box, new)
    click_text(page, "表示する")
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(2000)
    d = download(page, "CSVデータをダウンロードする")
    path = out_dir / d.suggested_filename
    d.save_as(path)
    first = path.read_bytes().decode(CSV_ENC, errors="ignore").splitlines()[1].split(",")[0]
    if not first.startswith(f"{y}{m:02d}"):
        path.unlink()
        raise RuntimeError(f"違う月のデータがダウンロードされました（{first}）")
    return path


def get_product(page, y, m, out_dir):
    page.goto(PRODUCT_URL)
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(1500)
    last = calendar.monthrange(y, m)[1]
    boxes = visible(page.locator("input")).all()
    box = next((b for b in boxes if re.search(r"\d{4}/\d{2}/\d{2}", b.input_value() or "")), None)
    if box is None:
        raise RuntimeError("期間を入れる欄が見つかりませんでした")
    old = box.input_value()
    sep = re.search(r"\d{4}/\d{2}/\d{2}(\s*\S\s*)\d{4}/\d{2}/\d{2}", old)
    new = f"{y}/{m:02d}/01{sep.group(1) if sep else ' ~ '}{y}/{m:02d}/{last:02d}"
    set_input_value(page, box, new)
    click_text(page, "表示する")
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(2000)
    d = download(page, "商品単位の売上")
    name = d.suggested_filename
    path = out_dir / name
    d.save_as(path)
    want = f"{y}{m:02d}01-{y}{m:02d}{last:02d}"
    if re.search(r"\d{8}-\d{8}", name) and want not in name:
        path.unlink()
        raise RuntimeError(f"違う期間のデータがダウンロードされました（{name}）")
    return path


def get_payout(ctx, page, y, m, out_dir):
    page.goto(PAYOUT_URL)
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(1500)
    if not select_year_month(page, y, m):
        page.goto(f"{PAYOUT_URL}?pullDownYearCdFrom={y}&pullDownMonthCdFrom={m}&search=検索")
    else:
        click_text(page, "検索", exact=True)
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(2000)
    target = "決済金額・手数料・振込金額"
    if visible(page.get_by_text(target)).count() == 0:
        click_text(page, "ダウンロード")  # メニューを開く
        page.wait_for_timeout(1000)
    d = download(page, target)
    path = out_dir / f"振込明細_{y}年{m:02d}月振込分.csv"
    d.save_as(path)
    dates = re.findall(r'^"(\d{4}-\d{2})-\d{2}"', path.read_bytes().decode(CSV_ENC, errors="ignore"), re.M)
    if dates and any(dd != f"{y}-{m:02d}" for dd in dates):
        path.unlink()
        raise RuntimeError(f"違う月の振込明細がダウンロードされました（{sorted(set(dates))}）")
    return path


def main():
    cfg = load_config()
    test = "--test" in sys.argv
    y, m = parse_month(sys.argv, dt.date.today())
    out_dir = HERE / "取得テスト" if test else Path(cfg["フォルダ"]["格納先"].strip())
    out_dir.mkdir(parents=True, exist_ok=True)
    if LOG_DIR.exists():
        shutil.rmtree(LOG_DIR)
    if not cfg.has_section("エアレジ") or not cfg["エアレジ"].get("パスワード", "").strip():
        sys.exit("設定.ini の [エアレジ] にログインIDとパスワードを書いてください。")
    try:
        from playwright.sync_api import sync_playwright
    except ImportError:
        sys.exit("playwright が入っていません。「エアレジ取得の準備.bat」を実行してください。")

    print(f"エアレジ・Airペイから {y}年{m}月分 を取得します（保存先: {out_dir}）")
    results = []
    with sync_playwright() as pw:
        # 前回のログイン状態を残すため、ツール専用のブラウザ設定フォルダを使う
        ctx = None
        for channel in ("msedge", "chrome", None):
            try:
                ctx = pw.chromium.launch_persistent_context(
                    str(HERE / "browser_profile"), headless=not test, accept_downloads=True,
                    viewport={"width": 1400, "height": 1000}, **({"channel": channel} if channel else {}))
                break
            except Exception:
                continue
        if ctx is None:
            sys.exit("Edge も Chrome も起動できませんでした")
        page = ctx.pages[0] if ctx.pages else ctx.new_page()
        try:
            login(page, cfg)
            print("  ログイン: OK")
        except Exception as e:
            save_trouble(page, "1_ログイン")
            print(f"  ログイン: 失敗（{e}）")
            ctx.close()
            return 1
        ny, nm = next_month(y, m)
        steps = [
            ("2_日別売上", lambda: get_daily(page, y, m, out_dir)),
            ("3_商品別売上", lambda: get_product(page, y, m, out_dir)),
            (f"4_振込明細_{m}月振込分", lambda: get_payout(ctx, page, y, m, out_dir)),
            (f"5_振込明細_{nm}月振込分", lambda: get_payout(ctx, page, ny, nm, out_dir)),
        ]
        for name, step in steps:
            try:
                path = step()
                print(f"  {name[2:]}: OK（{path.name}）")
                results.append(True)
            except Exception as e:
                save_trouble(page, name)
                print(f"  {name[2:]}: 失敗（{str(e).splitlines()[0]}）")
                results.append(False)
        ctx.close()
    if all(results):
        print("\nすべて取得できました。")
        return 0
    print(f"\nうまくいかなかった画面を次のフォルダに保存しました:\n  {LOG_DIR}")
    return 1


if __name__ == "__main__":
    sys.exit(main())
