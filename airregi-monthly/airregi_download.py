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

def fill_login_form(page, cfg):
    """ログイン画面・本人確認画面が出ていれば入力して進む。出ていなければ何もしない。"""
    pw_box = visible(page.locator("input[name=password], input[type=password]:not([name^=dummy])"))
    if pw_box.count() == 0:
        return
    a = cfg["エアレジ"]
    id_box = visible(page.locator("input[name=username]:not([type=hidden])"))
    if id_box.count() == 0:
        id_box = visible(page.locator("input:not([type=password]):not([type=hidden]):not([type=checkbox])"
                                      ":not([type=radio]):not([type=submit]):not([type=button]):not([name^=dummy])"))
    if id_box.count() > 0:  # 本人確認画面にはIDの欄がない
        id_box.first.fill(a["ログインID"].strip())
    pw_box.first.fill(a["パスワード"].strip())
    btn = visible(page.locator("input[type=submit][value='ログイン']"))
    if btn.count() == 0:
        btn = visible(page.get_by_role("button", name=re.compile(r"^\s*ログイン(する)?\s*$")))
    btn.last.click() if btn.count() else pw_box.first.press("Enter")
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(2000)
    if visible(page.locator("input[name=password], input[type=password]:not([name^=dummy])")).count() > 0:
        raise RuntimeError("ログインできませんでした（IDとパスワード、または確認コードの画面を確認してください）")


def open_page(page, cfg, url):
    """ページを開く。ログインや本人確認を求められたら入力して、もう一度開く。"""
    page.goto(url)
    page.wait_for_load_state("networkidle")
    if visible(page.locator("input[name=password], input[type=password]:not([name^=dummy])")).count() > 0:
        fill_login_form(page, cfg)
        if not page.url.startswith(url.split("?")[0]):
            page.goto(url)
            page.wait_for_load_state("networkidle")
    page.wait_for_timeout(1500)


def login(page, cfg):
    open_page(page, cfg, DAILY_URL)


def get_daily(page, cfg, y, m, out_dir):
    open_page(page, cfg, DAILY_URL)
    if not select_year_month(page, y, m):
        raise RuntimeError("対象月のプルダウンが見つかりませんでした")
    click_text(page, "表示する")
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(2000)
    click_text(page, "CSVデータをダウンロードする")
    page.wait_for_timeout(1000)
    with page.expect_download(timeout=30000) as info:
        page.locator("[data-testid=salesList-csv-download-tooltip] button", has_text="売上集計").first.click()
    d = info.value
    path = out_dir / d.suggested_filename
    d.save_as(path)
    first = path.read_bytes().decode(CSV_ENC, errors="ignore").splitlines()[1].split(",")[0]
    if not first.startswith(f"{y}{m:02d}"):
        path.unlink()
        raise RuntimeError(f"違う月のデータがダウンロードされました（{first}）")
    return path


def pick_range(page, y, m, last, test):
    """「期間選択」カレンダーを開いて、対象月の1日と末日を選んで「確定」を押す。"""
    box = page.locator(".input-date .datepicker input, .datepicker input").first
    box.click()
    cal = page.locator(".picker .datepicker-container").first
    cal.wait_for(state="visible", timeout=10000)
    if test:
        save_trouble(page, "3a_商品別売上_カレンダー")  # カレンダーの作りを確認する用
    header = cal.locator("tr.movement td.switch").first

    def shown():
        nums = re.findall(r"\d+", header.inner_text())
        return (int(nums[0]), int(nums[1])) if len(nums) >= 2 else None

    for _ in range(36):  # 対象月が表示されるまで「«」「»」で移動（「2026年09月」のような0付きにも対応）
        now = shown()
        if now is None:
            raise RuntimeError(f"カレンダーの年月が読めませんでした（{header.inner_text()!r}）")
        if now == (y, m):
            break
        row = cal.locator("tr.movement td")
        arrow = row.first if now > (y, m) else row.last
        arrow.click()
        page.wait_for_timeout(500)
        if shown() == now:  # 押しても月が変わらない
            arrow.locator("*").first.click() if arrow.locator("*").count() else None
            page.wait_for_timeout(500)
            if shown() == now:
                raise RuntimeError(f"カレンダーの「«」「»」で月を移動できませんでした（{header.inner_text().strip()}のまま）")
    else:
        raise RuntimeError(f"カレンダーに{y}年{m}月が出てきませんでした（{header.inner_text().strip()}）")
    for day in (1, last):
        # 前後の月の日付（class="old" / "new"）は除く
        cell = cal.locator("tbody td:not(.old):not(.new)").filter(has_text=re.compile(rf"^\s*{day}\s*$"))
        if cell.count() == 0:
            raise RuntimeError(f"カレンダーで{day}日を押せませんでした")
        cell.first.click()
        page.wait_for_timeout(300)
    cal.locator("button.btn-confirm, button:has-text('確定')").first.click()
    page.wait_for_timeout(500)
    value = box.input_value()
    if f"{y}/{m:02d}/01" not in value or f"{y}/{m:02d}/{last:02d}" not in value:
        raise RuntimeError(f"期間を設定できませんでした（{value}）")


def get_product(page, cfg, y, m, out_dir, test):
    open_page(page, cfg, PRODUCT_URL)
    last = calendar.monthrange(y, m)[1]
    pick_range(page, y, m, last, test)
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


def get_payout(page, cfg, y, m, out_dir):
    open_page(page, cfg, PAYOUT_URL)
    if not select_year_month(page, y, m):
        open_page(page, cfg, f"{PAYOUT_URL}?pullDownYearCdFrom={y}&pullDownMonthCdFrom={m}&search=検索")
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


def run(cfg, y, m, out_dir, show=False, skip=()):
    """ダウンロードして、失敗した手順の名前の一覧を返す。skip に入れた手順（daily / product /
    payout / payout_next）は、もうファイルがあるものとして飛ばす。"""
    out_dir.mkdir(parents=True, exist_ok=True)
    if LOG_DIR.exists():
        shutil.rmtree(LOG_DIR)
    if not cfg.has_section("エアレジ") or not cfg["エアレジ"].get("パスワード", "").strip():
        print("  設定.ini の [エアレジ] にログインIDとパスワードがありません")
        return ["ログイン"]
    ny, nm = next_month(y, m)
    steps = [
        ("daily", "2_日別売上", lambda page: get_daily(page, cfg, y, m, out_dir)),
        ("product", "3_商品別売上", lambda page: get_product(page, cfg, y, m, out_dir, show)),
        ("payout", f"4_振込明細_{m}月振込分", lambda page: get_payout(page, cfg, y, m, out_dir)),
        ("payout_next", f"5_振込明細_{nm}月振込分", lambda page: get_payout(page, cfg, ny, nm, out_dir)),
    ]
    steps = [st for st in steps if st[0] not in skip]
    print(f"エアレジ・Airペイから {y}年{m}月分 を取得します（保存先: {out_dir}）")
    if not steps:
        print("  必要なファイルはそろっているので、ダウンロードしませんでした")
        return []
    try:
        from playwright.sync_api import sync_playwright
    except ImportError:
        print("  playwright が入っていません。「エアレジ取得の準備.bat」を実行してください。")
        return ["準備"]

    failed = []
    with sync_playwright() as pw:
        # 前回のログイン状態を残すため、ツール専用のブラウザ設定フォルダを使う
        ctx = None
        for channel in ("msedge", "chrome", None):
            try:
                ctx = pw.chromium.launch_persistent_context(
                    str(HERE / "browser_profile"), headless=not show, accept_downloads=True,
                    viewport={"width": 1400, "height": 1000}, **({"channel": channel} if channel else {}))
                break
            except Exception:
                continue
        if ctx is None:
            print("  Edge も Chrome も起動できませんでした")
            return ["ブラウザ"]
        page = ctx.pages[0] if ctx.pages else ctx.new_page()
        try:
            login(page, cfg)
            print("  ログイン: OK")
        except Exception as e:
            save_trouble(page, "1_ログイン")
            print(f"  ログイン: 失敗（{e}）")
            ctx.close()
            return ["ログイン"]
        for _, name, step in steps:
            try:
                path = step(page)
                print(f"  {name[2:]}: OK（{path.name}）")
            except Exception as e:
                save_trouble(page, name)
                print(f"  {name[2:]}: 失敗（{str(e).splitlines()[0]}）")
                failed.append(name[2:])
        ctx.close()
    if failed:
        print(f"  うまくいかなかった画面を次のフォルダに保存しました: {LOG_DIR}")
    return failed


def main():
    cfg = load_config()
    test = "--test" in sys.argv
    y, m = parse_month(sys.argv, dt.date.today())
    out_dir = HERE / "取得テスト" if test else Path(cfg["フォルダ"]["格納先"].strip())
    failed = run(cfg, y, m, out_dir, show=test)
    print("\nすべて取得できました。" if not failed else "\n一部取得できませんでした。")
    return 1 if failed else 0


if __name__ == "__main__":
    sys.exit(main())
