# -*- coding: utf-8 -*-
"""
請求書の自動まとめ＆税理士さんへの送信ツール

毎日1回（タスクスケジューラから）実行されると:
  1. ヤフーメールから、登録した取引先の請求書メールを探して
     添付ファイルを「請求書\\今回送る分」フォルダに保存します。
  2. 毎月20日以降で、その月にまだ送っていなければ、
     フォルダを圧縮し、確認メッセージを出します。
     「はい」を押したときだけ税理士さんへメールで送ります。

使い方（手動実行）:
  py invoice_tool.py            … 受信と（20日以降なら）送信確認
  py invoice_tool.py --collect  … 受信だけ
  py invoice_tool.py --send-now … 日付に関係なく、今すぐ送信確認を出す（テスト用）
"""
import configparser
import datetime as dt
import email
import email.policy
import imaplib
import json
import os
import re
import shutil
import smtplib
import ssl
import sys
import traceback
import zipfile
from email.message import EmailMessage
from email.utils import parseaddr
from pathlib import Path

TOOL_DIR = Path(__file__).resolve().parent
CONFIG_PATH = TOOL_DIR / "config.ini"
STATE_PATH = TOOL_DIR / "state.json"
LOG_PATH = TOOL_DIR / "log.txt"

IMAP_HOST = "imap.mail.yahoo.co.jp"
SMTP_HOST = "smtp.mail.yahoo.co.jp"
SMTP_PORT = 465
MAX_ATTACH_MB = 20  # ヤフーメールの送信上限（25MB）より少し小さめ
VENDOR_PREFIX = "取引先:"


# ---------------------------------------------------------------- 共通
def log(msg):
    line = f"{dt.datetime.now():%Y-%m-%d %H:%M:%S} {msg}"
    print(line)
    with open(LOG_PATH, "a", encoding="utf-8") as f:
        f.write(line + "\n")


def load_config():
    if not CONFIG_PATH.exists():
        raise RuntimeError(
            "config.ini が見つかりません。\n"
            "config.ini.sample をコピーして config.ini を作ってください。"
        )
    cfg = configparser.ConfigParser(interpolation=None)
    cfg.read(CONFIG_PATH, encoding="utf-8-sig")
    return cfg


def load_state():
    if STATE_PATH.exists():
        return json.loads(STATE_PATH.read_text(encoding="utf-8"))
    return {"processed": [], "last_sent_month": ""}


def save_state(state):
    state["processed"] = state["processed"][-2000:]
    STATE_PATH.write_text(json.dumps(state, ensure_ascii=False, indent=1), encoding="utf-8")


def folders(cfg):
    base = TOOL_DIR / cfg.get("フォルダ", "保存先", fallback="請求書")
    pending = base / "今回送る分"
    sent = base / "送信済み"
    pending.mkdir(parents=True, exist_ok=True)
    sent.mkdir(parents=True, exist_ok=True)
    return pending, sent


def safe_name(name):
    name = re.sub(r'[\\/:*?"<>|\r\n\t]', "_", name).strip(" .")
    return name or "file"


def unique_path(path):
    if not path.exists():
        return path
    i = 2
    while True:
        p = path.with_name(f"{path.stem}({i}){path.suffix}")
        if not p.exists():
            return p
        i += 1


def vendors(cfg):
    result = []
    for sec in cfg.sections():
        if not sec.startswith(VENDOR_PREFIX):
            continue
        s = cfg[sec]
        senders = [x.strip().lower() for x in s.get("差出人", "").split(",") if x.strip()]
        if not senders:
            continue
        result.append({
            "name": sec[len(VENDOR_PREFIX):].strip(),
            "senders": senders,
            "keyword": s.get("件名キーワード", "").strip(),
            "pdf_password": s.get("PDFパスワード", "").strip(),
            "site": s.get("サイト", "").strip().lower(),
            "login_id": s.get("ログインID", "").strip(),
            "login_password": s.get("ログインパスワード", "").strip(),
        })
    return result


# ---------------------------------------------------------------- 受信
def decrypt_pdf(path, password):
    """パスワード付きPDFを解除して上書きする。失敗したら元のまま残す。"""
    try:
        import pikepdf
    except ImportError:
        log("  ※pikepdf が未インストールのため、PDFのパスワードは解除せず保存しました")
        return
    try:
        with pikepdf.open(path, password=password, allow_overwriting_input=True) as pdf:
            pdf.save(path)
        log("  PDFのパスワードを解除しました")
    except Exception as e:
        log(f"  ※PDFのパスワード解除に失敗（パスワードを確認してください）: {e}")


def collect(cfg, state):
    pending, _ = folders(cfg)
    vlist = vendors(cfg)
    if not vlist:
        log("取引先が登録されていません（config.ini の [取引先:○○] を追加してください）")
        return 0

    user = cfg.get("ヤフーメール", "メールアドレス")
    password = cfg.get("ヤフーメール", "パスワード")
    mailbox = cfg.get("ヤフーメール", "検索フォルダ", fallback="INBOX")
    days = cfg.getint("ヤフーメール", "さかのぼる日数", fallback=45)
    # IMAPの日付は英語の月名が必要（日本語Windowsの strftime では日本語になるため固定表記）
    months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
    d = dt.date.today() - dt.timedelta(days=days)
    since = f"{d.day:02d}-{months[d.month - 1]}-{d.year}"

    processed = set(state["processed"])
    saved = 0
    site_jobs = {}  # 取引先名 -> [(請求書番号, 発行日, 処理済みキー)]
    imap = imaplib.IMAP4_SSL(IMAP_HOST, 993, ssl_context=ssl.create_default_context())
    try:
        imap.login(user, password)
        imap.select(mailbox, readonly=True)  # 既読にしない
        for v in vlist:
            for sender in v["senders"]:
                typ, data = imap.search(None, "SINCE", since, "FROM", f'"{sender}"')
                if typ != "OK" or not data or not data[0]:
                    continue
                for num in data[0].split():
                    typ, msg_data = imap.fetch(num, "(BODY.PEEK[])")
                    if typ != "OK":
                        continue
                    raw = next((p[1] for p in msg_data if isinstance(p, tuple)), None)
                    if raw is None:
                        continue
                    msg = email.message_from_bytes(raw, policy=email.policy.default)
                    msg_id = (msg.get("Message-ID") or f"{sender}-{num.decode()}").strip()
                    key = f"{v['name']}|{'site|' if v['site'] else ''}{msg_id}"
                    if key in processed:
                        continue
                    from_addr = parseaddr(str(msg.get("From", "")))[1].lower()
                    if sender not in from_addr:
                        continue
                    subject = str(msg.get("Subject", ""))
                    if v["keyword"] and v["keyword"] not in subject:
                        continue

                    try:
                        date = email.utils.parsedate_to_datetime(msg["Date"]).date()
                    except Exception:
                        date = dt.date.today()

                    if v["site"]:
                        # 添付なし：件名から請求書番号を読み取り、あとでサイトからダウンロードする
                        m = re.search(r"(\d{4,})\s*[:：]", subject)
                        if not m:
                            log(f"※件名から請求書番号を読み取れませんでした: {v['name']}／{subject}")
                            continue
                        dm = re.search(r"\[(\d{4})/(\d{1,2})/(\d{1,2})\]", subject)
                        if dm:
                            date = dt.date(*map(int, dm.groups()))
                        site_jobs.setdefault(v["name"], []).append((m.group(1), date, key))
                        continue

                    count = 0
                    for part in msg.iter_attachments():
                        fname = part.get_filename()
                        content = part.get_payload(decode=True)
                        if not fname or not content:
                            continue
                        out = unique_path(pending / safe_name(f"{date:%Y-%m-%d}_{v['name']}_{fname}"))
                        out.write_bytes(content)
                        log(f"保存: {out.name}（件名: {subject}）")
                        if v["pdf_password"] and out.suffix.lower() == ".pdf":
                            decrypt_pdf(out, v["pdf_password"])
                        count += 1
                    if count == 0:
                        log(f"※添付ファイルのない請求書メールがありました: {v['name']}／{subject}"
                            "（サイトからのダウンロードが必要かもしれません）")
                    saved += count
                    processed.add(key)
                    state["processed"].append(key)
    finally:
        try:
            imap.logout()
        except Exception:
            pass

    for v in vlist:
        jobs = site_jobs.get(v["name"])
        if not jobs:
            continue
        done = download_from_site(v, jobs, pending)
        saved += len(done)
        state["processed"].extend(done)
    save_state(state)
    log(f"受信チェック完了: 新しく保存したファイル {saved} 件")
    return saved


# ---------------------------------------------------------------- サイトからのダウンロード
SHOW_BROWSER = False  # --show-browser を付けると、ブラウザの動きが画面に表示されます


def open_browser(pw):
    """パソコンに入っている Edge（なければ Chrome）を使う。"""
    for channel in ("msedge", "chrome", None):
        try:
            kw = {"channel": channel} if channel else {}
            return pw.chromium.launch(headless=not SHOW_BROWSER, **kw)
        except Exception:
            continue
    raise RuntimeError("Edge も Chrome も起動できませんでした")


def click_and_download(ctx, page, target, out):
    """ボタンを押して、ダウンロード（または新しいタブで開いたPDF）を保存する。"""
    got = []

    def on_download(d):
        got.append(d)

    def hook(p):
        p.on("download", on_download)

    for p in ctx.pages:
        hook(p)
    ctx.on("page", hook)
    before = len(ctx.pages)
    target.click()
    for _ in range(60):
        if got:
            break
        page.wait_for_timeout(500)
    ctx.remove_listener("page", hook)
    if got:
        got[0].save_as(out)
    elif len(ctx.pages) > before:  # PDFが新しいタブで表示された場合
        url = ctx.pages[-1].url
        resp = ctx.request.get(url)
        if not resp.ok or not resp.body().startswith(b"%PDF"):
            raise RuntimeError("PDFを取得できませんでした")
        out.write_bytes(resp.body())
        ctx.pages[-1].close()
    else:
        raise RuntimeError("ダウンロードが始まりませんでした")


def fill_field(page, label, fallback_css, value):
    field = page.get_by_label(label, exact=True)
    if field.count() != 1:
        field = page.locator(fallback_css).first
    field.fill(value)


MEIRAKU_LOGIN_URL = "https://meiraku-invoice.jp/customer/customer-users/login"


def download_meiraku(page, ctx, v, jobs, pending, done):
    """スジャータめいらく請求書システム（meiraku-invoice.jp）"""
    page.goto(MEIRAKU_LOGIN_URL)
    fill_field(page, "ログインID", "input[type=text], input[type=email], input:not([type])", v["login_id"])
    fill_field(page, "パスワード", "input[type=password]", v["login_password"])
    page.get_by_role("button", name="ログイン", exact=True).click()
    page.wait_for_load_state("networkidle")
    if "login" in page.url:
        raise RuntimeError("ログインできませんでした。ログインID／ログインパスワードを確認してください")

    for number, date, key in jobs:
        row = page.locator("tr", has_text=number)
        if row.count() == 0:
            # 一覧に見当たらなければ、請求書番号で検索する
            fill_field(page, "請求書番号", "input[type=text]", number)
            page.get_by_role("button", name="検索", exact=True).click()
            page.wait_for_load_state("networkidle")
            row = page.locator("tr", has_text=number)
        if row.count() == 0:
            log(f"※{v['name']}: 請求書番号 {number} が一覧に見つかりませんでした")
            continue
        button = row.first.locator("a, button", has_text="PDF").first
        out = unique_path(pending / safe_name(f"{date:%Y-%m-%d}_{v['name']}_請求書{number}.pdf"))
        click_and_download(ctx, page, button, out)
        log(f"保存: {out.name}（サイトからダウンロード）")
        done.append(key)


SITES = {"meiraku": download_meiraku}


def download_from_site(v, jobs, pending):
    """サイトにログインして請求書をダウンロード。成功した分の処理済みキーを返す。"""
    func = SITES.get(v["site"])
    if func is None:
        log(f"※{v['name']}: サイト = {v['site']} には対応していません")
        return []
    if not v["login_id"] or not v["login_password"]:
        log(f"※{v['name']}: ログインID／ログインパスワードが config.ini にありません")
        return []
    try:
        from playwright.sync_api import sync_playwright
    except ImportError:
        log(f"※{v['name']}: playwright が未インストールのため、サイトからのダウンロードができません")
        return []
    done = []
    try:
        with sync_playwright() as pw:
            browser = open_browser(pw)
            try:
                ctx = browser.new_context(accept_downloads=True)
                func(ctx.new_page(), ctx, v, jobs, pending, done)
            finally:
                browser.close()
    except Exception as e:
        log(f"※{v['name']}: サイトからのダウンロードでエラー: {e}")
        notify("請求書ツール", f"{v['name']} のサイトから請求書をダウンロードできませんでした。\n\n{e}\n\n"
               "明日もう一度試します。詳しくは log.txt を見てください。", error=True)
    return done


# ---------------------------------------------------------------- 送信
def target_label(cfg, today):
    mode = cfg.get("税理士さん", "何月分", fallback="前月").strip()
    if mode == "当月":
        y, m = today.year, today.month
    else:
        first = today.replace(day=1) - dt.timedelta(days=1)
        y, m = first.year, first.month
    return y, m


def ask(title, text, cancel=False):
    """Windowsの確認メッセージを表示。はい=True, いいえ=False, キャンセル=None"""
    import tkinter as tk
    from tkinter import messagebox
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    try:
        if cancel:
            return messagebox.askyesnocancel(title, text, parent=root)
        return messagebox.askyesno(title, text, parent=root)
    finally:
        root.destroy()


def notify(title, text, error=False):
    try:
        import tkinter as tk
        from tkinter import messagebox
        root = tk.Tk()
        root.withdraw()
        root.attributes("-topmost", True)
        (messagebox.showerror if error else messagebox.showinfo)(title, text, parent=root)
        root.destroy()
    except Exception:
        pass


def send_flow(cfg, state, force=False):
    today = dt.date.today()
    send_day = cfg.getint("税理士さん", "送信日", fallback=20)
    this_month = f"{today:%Y-%m}"
    if not force:
        if today.day < send_day or state.get("last_sent_month") == this_month:
            return
    # 送信日より前のテスト送信（--send-now）は「今月送った」と記録しない（20日の確認が出なくなるため）
    record = today.day >= send_day

    pending, sent = folders(cfg)
    files = sorted(p for p in pending.iterdir() if p.is_file())
    y, m = target_label(cfg, today)
    label = f"{m}月分請求書"

    if not files:
        skip = ask("請求書の送信",
                   f"「今回送る分」フォルダに請求書がありません（{label}）。\n\n"
                   "今月は送らずに終わりにしますか？\n\n"
                   "はい … 今月は送りません\n"
                   "いいえ … 明日もう一度確認します")
        if skip:
            if record:
                state["last_sent_month"] = this_month
            save_state(state)
            log(f"{label}: 請求書が0件のため、今月は送信しませんでした")
        return

    to_addr = cfg.get("税理士さん", "宛先")
    subject = cfg.get("税理士さん", "件名", fallback="{月分}").replace("{月分}", label)
    body = cfg.get("税理士さん", "本文", fallback="{月分}をお送りします。").replace("{月分}", label)
    body = body.replace("\\n", "\n")

    zip_path = sent / f"請求書_{y}年{m}月分.zip"
    zip_path = unique_path(zip_path)
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as z:
        for p in files:
            z.write(p, arcname=p.name)
    size_mb = zip_path.stat().st_size / 1024 / 1024

    if hasattr(os, "startfile"):
        os.startfile(pending)  # フォルダを開いて中身を見られるように
    listing = "\n".join(f"・{p.name}" for p in files[:25])
    if len(files) > 25:
        listing += f"\n…ほか {len(files) - 25} 件"
    warn = ""
    if size_mb > MAX_ATTACH_MB:
        warn = f"\n\n⚠ 圧縮ファイルが {size_mb:.1f}MB あり、大きすぎて送れない可能性があります。"

    answer = ask("税理士さんへ送信してよいですか？",
                 f"宛先: {to_addr}\n件名: {subject}\n添付: {zip_path.name}（{size_mb:.1f}MB）\n\n"
                 f"入っている請求書（{len(files)}件）:\n{listing}{warn}\n\n"
                 "はい … 今すぐ送信します\n"
                 "いいえ … 今日は送らず、明日もう一度確認します\n"
                 "キャンセル … 今月は送りません",
                 cancel=True)
    if answer is None:
        zip_path.unlink(missing_ok=True)
        if record:
            state["last_sent_month"] = this_month
        save_state(state)
        log(f"{label}: 今月は送らないことにしました（ファイルはそのまま残っています）")
        return
    if answer is False:
        zip_path.unlink(missing_ok=True)
        log(f"{label}: 送信を見送りました（明日もう一度確認します）")
        return

    user = cfg.get("ヤフーメール", "メールアドレス")
    password = cfg.get("ヤフーメール", "パスワード")
    msg = EmailMessage()
    msg["From"] = user
    msg["To"] = to_addr
    if cfg.get("税理士さん", "控えを自分にも送る", fallback="はい").strip() in ("はい", "yes", "true", "1"):
        msg["Bcc"] = user
    msg["Subject"] = subject
    msg.set_content(body)
    msg.add_attachment(zip_path.read_bytes(), maintype="application", subtype="zip",
                       filename=zip_path.name)
    try:
        with smtplib.SMTP_SSL(SMTP_HOST, SMTP_PORT, context=ssl.create_default_context()) as s:
            s.login(user, password)
            s.send_message(msg)
    except Exception:
        zip_path.unlink(missing_ok=True)  # 送れなかったら作りかけの圧縮ファイルは消す（明日また確認）
        raise

    # 送ったファイルは「送信済み\\○年○月分」へ移動
    done_dir = unique_path(sent / f"{y}年{m}月分")
    done_dir.mkdir(parents=True)
    for p in files:
        shutil.move(str(p), str(done_dir / p.name))
    if record:
        state["last_sent_month"] = this_month
    save_state(state)
    log(f"{label}: {to_addr} へ送信しました（{len(files)}件）")
    notify("送信完了", f"{label}を税理士さんへ送信しました。\n\n送ったファイルは\n{done_dir}\nに移動しました。")


# ---------------------------------------------------------------- 入口
def main():
    global SHOW_BROWSER
    args = sys.argv[1:]
    SHOW_BROWSER = "--show-browser" in args
    try:
        cfg = load_config()
        state = load_state()
        collect(cfg, state)
        if "--collect" not in args:
            send_flow(cfg, state, force="--send-now" in args)
    except Exception as e:
        log("エラー: " + "".join(traceback.format_exception(e)))
        hint = ""
        if isinstance(e, (imaplib.IMAP4.error, smtplib.SMTPAuthenticationError)):
            hint = ("\n\nヤフーメールにログインできませんでした。\n"
                    "・config.ini のメールアドレス／パスワード\n"
                    "・ヤフーメールの「IMAP/POP/SMTPアクセス」設定\n"
                    "を確認してください。")
        notify("請求書ツールでエラーが起きました", f"{e}{hint}\n\n詳しくは log.txt を見てください。", error=True)
        sys.exit(1)


if __name__ == "__main__":
    main()
