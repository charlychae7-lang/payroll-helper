"""급여명세서 메일 발송 도구.

급여앱의 [메일 발송용 파일 저장]으로 만든 zip을 열어, 직원마다 본인의 PDF만 첨부해 보냅니다.
앱 비밀번호는 실행 중에만 사용하고 어디에도 저장하지 않습니다. 추가 설치 없이 Python 표준 기능만 씁니다.
"""
import csv
import io
import json
import os
import queue
import re
import shutil
import smtplib
import ssl
import tempfile
import threading
import time
import zipfile
from datetime import datetime
from email.message import EmailMessage
from email.utils import formataddr, formatdate, make_msgid
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

BASE = Path(__file__).resolve().parent
LIST_NAME = "발송목록.csv"
LOG_HEAD = ["보낸시각", "연월", "성명", "이메일", "파일", "결과"]
SERVICES = {
    "Gmail": ("smtp.gmail.com", 465, "ssl"),
    "네이버": ("smtp.naver.com", 465, "ssl"),
    "다음·한메일": ("smtp.daum.net", 465, "ssl"),
    "카카오메일": ("smtp.kakao.com", 465, "ssl"),
    "직접 입력": ("", 465, "ssl"),
}
DEFAULTS = {
    "service": "Gmail", "host": "smtp.gmail.com", "port": 465, "security": "ssl",
    "sender": "", "name": "",
    "subject": "[{회사}] {연}년 {월}월 급여명세서",
    "body": "{성명}님,\n\n{연}년 {월}월 급여명세서를 첨부해 보내 드립니다.\n확인 후 문의 사항은 담당자에게 알려 주세요.\n\n{회사}",
}
EMAIL = re.compile(r"^[A-Za-z0-9._%+\-]+@[A-Za-z0-9\-]+(\.[A-Za-z0-9\-]+)*\.[A-Za-z]{2,}$")


# ───────── 자료 처리 ─────────
def load_package(path):
    """zip을 읽어 (받는 사람 목록, {파일명: PDF 내용})을 돌려준다. 문제가 있는 줄은 error에 이유를 적는다."""
    with zipfile.ZipFile(path) as z:
        names = z.namelist()
        if LIST_NAME not in names:
            raise ValueError("발송목록이 없는 파일입니다. 급여앱에서 [메일 발송용 파일 저장]으로 만든 zip을 고르세요.")
        table = list(csv.DictReader(io.StringIO(z.read(LIST_NAME).decode("utf-8-sig"))))
        if not table or not {"연월", "회사", "성명", "이메일", "파일"} <= set(table[0]):
            raise ValueError("발송목록의 형식이 올바르지 않습니다.")
        rows, files = [], {}
        file_count, mail_count = {}, {}
        for r in table:
            file_count[r["파일"]] = file_count.get(r["파일"], 0) + 1
            key = (r["이메일"] or "").strip().lower()
            mail_count[key] = mail_count.get(key, 0) + 1
        for r in table:
            row = {k: (v or "").strip() for k, v in r.items() if k}
            error = ""
            data = z.read(row["파일"]) if row["파일"] in names else None
            if not re.fullmatch(r"\d{4}-\d{2}", row["연월"]):
                error = "연월 형식 오류"
            elif data is None or not data.startswith(b"%PDF"):
                error = "PDF 파일 없음"
            elif file_count[row["파일"]] > 1 or not row["파일"].startswith(f"{row['연월']}_급여명세서_{row['성명']}"):
                error = "파일과 직원 이름 불일치"
            elif not row["이메일"]:
                error = "이메일 없음"
            elif not EMAIL.match(row["이메일"]):
                error = "이메일 형식 오류"
            row["error"] = error
            row["duplicate"] = bool(row["이메일"]) and mail_count[row["이메일"].lower()] > 1
            rows.append(row)
            if data is not None:
                files[row["파일"]] = data
    return rows, files


def fill(template, row):
    year, month = row["연월"].split("-")
    for key, value in (("{회사}", row["회사"]), ("{성명}", row["성명"]), ("{연}", year), ("{월}", str(int(month)))):
        template = template.replace(key, value)
    return template


def build_message(cfg, row, pdf, to=None, test=False):
    """한 사람에게 본인의 PDF 하나만 붙인 메일을 만든다."""
    msg = EmailMessage()
    msg["From"] = formataddr((cfg["name"], cfg["sender"])) if cfg["name"] else cfg["sender"]
    msg["To"] = to or row["이메일"]
    msg["Subject"] = ("[시험] " if test else "") + fill(cfg["subject"], row).replace("\n", " ").strip()
    msg["Date"] = formatdate(localtime=True)
    msg["Message-ID"] = make_msgid(domain=cfg["sender"].rsplit("@", 1)[-1])
    msg.set_content(fill(cfg["body"], row))
    msg.add_attachment(pdf, maintype="application", subtype="pdf", filename=row["파일"])
    return msg


def connect(cfg, password):
    """메일 서버에 암호화 연결 후 로그인한다. 암호화 없는 연결은 이 PC 안의 시험 서버만 허용한다."""
    host, port, security = cfg["host"].strip(), int(cfg["port"]), cfg["security"]
    local = host in ("localhost", "127.0.0.1")
    context = ssl.create_default_context()
    if security == "ssl":
        server = smtplib.SMTP_SSL(host, port, local_hostname="localhost", timeout=30, context=context)
    elif security == "starttls" or (security == "none" and local):
        server = smtplib.SMTP(host, port, local_hostname="localhost", timeout=30)
        if security == "starttls":
            server.starttls(context=context)
    else:
        raise ValueError("암호화되지 않은 연결로는 보낼 수 없습니다.")
    try:
        server.login(cfg["sender"], password)
    except BaseException:
        server.close()
        raise
    return server


def read_sent(log_path):
    """이미 성공한 발송: {(연월, 성명, 이메일): 보낸 시각}"""
    if not log_path.exists():
        return {}
    with open(log_path, encoding="utf-8-sig", newline="") as f:
        return {(r["연월"], r["성명"], r["이메일"].lower()): r["보낸시각"] for r in csv.DictReader(f) if r.get("결과") == "성공"}


def append_log(log_path, row, to, result):
    new = not log_path.exists()
    with open(log_path, "a", encoding="utf-8-sig" if new else "utf-8", newline="") as f:
        writer = csv.writer(f)
        if new:
            writer.writerow(LOG_HEAD)
        writer.writerow([datetime.now().strftime("%Y-%m-%d %H:%M:%S"), row["연월"], row["성명"], to, row["파일"], result])


def explain(error):
    if isinstance(error, smtplib.SMTPAuthenticationError):
        return "로그인 실패: 보내는 주소, 앱 비밀번호, 메일의 SMTP(IMAP/SMTP) 사용 설정을 확인하세요."
    if isinstance(error, smtplib.SMTPRecipientsRefused):
        return "받는 주소가 거부되었습니다."
    if isinstance(error, ssl.SSLError):
        return "보안 연결에 실패했습니다. 서버 주소와 포트를 확인하세요."
    if isinstance(error, smtplib.SMTPResponseException):
        detail = error.smtp_error.decode("utf-8", "replace") if isinstance(error.smtp_error, bytes) else str(error.smtp_error)
        return f"메일 서버 오류 {error.smtp_code}: {detail[:120]}"
    if isinstance(error, (OSError, smtplib.SMTPException)):
        return f"연결에 실패했습니다: {error}"
    return str(error)


def send_all(cfg, password, jobs, files, log_path, report, test_to=None, delay=1.0):
    """jobs의 (번호, 줄)을 차례로 보낸다. report(번호, 상태)로 진행을 알리고 (성공, 실패) 수를 돌려준다."""
    test = test_to is not None
    if not test and log_path.exists():
        open(log_path, "a").close()          # 기록 파일이 잠겨 있으면 보내기 전에 멈춘다
    server = connect(cfg, password)
    done = failed = 0
    try:
        for n, (index, row) in enumerate(jobs):
            to = test_to or row["이메일"]
            report(index, "보내는 중…")
            lost = False
            try:
                server.send_message(build_message(cfg, row, files[row["파일"]], to, test))
                result, done = "성공", done + 1
            except Exception as error:      # 한 사람의 오류는 기록하고 다음 사람으로, 연결이 끊기면 중단
                lost = isinstance(error, smtplib.SMTPServerDisconnected) or (isinstance(error, OSError) and not isinstance(error, smtplib.SMTPException))
                result, failed = "실패: " + explain(error), failed + 1
            report(index, ("시험 " if test else "") + result)
            if not test:
                append_log(log_path, row, to, result)
            if lost:
                for rest, _ in jobs[n + 1:]:
                    report(rest, "보내지 않음 (연결 끊김)")
                break
            if n + 1 < len(jobs):
                time.sleep(delay)
    finally:
        try:
            server.quit()
        except (smtplib.SMTPException, OSError):
            server.close()
    return done, failed


# ───────── 화면 ─────────
class App:
    def __init__(self, root, base=BASE):
        self.root = root
        self.config_path = base / "메일발송_설정.json"
        self.log_path = base / "메일발송_기록.csv"
        self.rows, self.files, self.checked, self.busy = [], {}, set(), False
        self.events = queue.Queue()
        self.preview_dir = None
        self.cfg = dict(DEFAULTS)
        try:
            saved = json.loads(self.config_path.read_text(encoding="utf-8"))
            self.cfg.update({k: saved[k] for k in DEFAULTS if k in saved})
        except (OSError, ValueError, TypeError):
            pass
        self.build()
        root.protocol("WM_DELETE_WINDOW", self.close)
        root.after(100, self.poll)

    def build(self):
        root = self.root
        root.title("급여명세서 메일 발송")
        root.geometry("900x740")
        root.minsize(760, 620)
        ttk.Style().configure("Treeview", rowheight=24)
        pad = {"padx": 12, "pady": 5}

        box = ttk.LabelFrame(root, text="1. 발송 파일 (급여앱에서 저장한 …_메일발송.zip)", padding=8)
        box.pack(fill="x", **pad)
        self.path = tk.StringVar()
        ttk.Entry(box, textvariable=self.path, state="readonly").pack(side="left", fill="x", expand=True)
        self.open_button = ttk.Button(box, text="파일 고르기…", command=self.choose)
        self.open_button.pack(side="left", padx=(8, 0))

        box = ttk.LabelFrame(root, text="2. 보내는 메일 계정", padding=8)
        box.pack(fill="x", **pad)
        self.vars = {k: tk.StringVar(value=str(self.cfg[k])) for k in ("service", "host", "port", "security", "sender", "name")}
        self.password = tk.StringVar()

        def label(text, r, c):
            ttk.Label(box, text=text).grid(row=r, column=c, sticky="e", padx=(8, 4), pady=3)

        label("메일 서비스", 0, 0)
        service = ttk.Combobox(box, textvariable=self.vars["service"], values=list(SERVICES), state="readonly", width=16)
        service.grid(row=0, column=1, sticky="w")
        service.bind("<<ComboboxSelected>>", lambda _: self.pick_service())
        label("서버 주소", 0, 2)
        self.host_entry = ttk.Entry(box, textvariable=self.vars["host"], width=22)
        self.host_entry.grid(row=0, column=3, sticky="w")
        label("포트", 0, 4)
        self.port_entry = ttk.Entry(box, textvariable=self.vars["port"], width=6)
        self.port_entry.grid(row=0, column=5, sticky="w")
        self.security_box = ttk.Combobox(box, textvariable=self.vars["security"], values=["ssl", "starttls"], width=9)
        self.security_box.grid(row=0, column=6, sticky="w", padx=(8, 0))
        label("보내는 주소", 1, 0)
        ttk.Entry(box, textvariable=self.vars["sender"], width=28).grid(row=1, column=1, columnspan=2, sticky="we")
        label("앱 비밀번호", 1, 3)
        ttk.Entry(box, textvariable=self.password, show="●", width=24).grid(row=1, column=4, columnspan=3, sticky="w")
        label("보내는 사람 이름", 2, 0)
        ttk.Entry(box, textvariable=self.vars["name"], width=28).grid(row=2, column=1, columnspan=2, sticky="we")
        ttk.Label(box, text="앱 비밀번호는 저장하지 않습니다. 메일 설정에서 SMTP 사용을 켜야 합니다.", foreground="#666").grid(row=2, column=3, columnspan=4, sticky="w", padx=8)
        self.lock_server()

        box = ttk.LabelFrame(root, text="3. 제목과 본문 ({회사} {성명} {연} {월}은 자동으로 바뀝니다)", padding=8)
        box.pack(fill="x", **pad)
        self.subject = tk.StringVar(value=self.cfg["subject"])
        ttk.Entry(box, textvariable=self.subject).pack(fill="x")
        self.body = tk.Text(box, height=5, wrap="word", font=("맑은 고딕", 10), undo=True)
        self.body.insert("1.0", self.cfg["body"])
        self.body.pack(fill="x", pady=(6, 0))

        bottom = ttk.Frame(root, padding=(12, 4, 12, 10))
        bottom.pack(side="bottom", fill="x")
        self.status = tk.StringVar(value="발송 파일을 고르세요.")
        ttk.Label(bottom, textvariable=self.status, wraplength=480).pack(side="left")
        self.send_button = ttk.Button(bottom, text="선택한 사람에게 발송", command=self.send)
        self.send_button.pack(side="right")
        self.test_button = ttk.Button(bottom, text="나에게 시험 발송", command=self.send_test)
        self.test_button.pack(side="right", padx=8)

        box = ttk.LabelFrame(root, text="4. 받는 사람 (첫 칸을 눌러 선택·해제, 줄을 두 번 누르면 PDF 열기)", padding=8)
        box.pack(fill="both", expand=True, **pad)
        tools = ttk.Frame(box)
        tools.pack(fill="x", pady=(0, 6))
        ttk.Button(tools, text="보낼 수 있는 사람 모두 선택", command=self.select_all).pack(side="left")
        ttk.Button(tools, text="모두 해제", command=self.select_none).pack(side="left", padx=6)
        ttk.Button(tools, text="선택한 줄의 PDF 열기", command=self.preview).pack(side="left")
        columns = [("check", "발송", 50), ("name", "성명", 100), ("mail", "이메일", 200), ("file", "첨부 파일", 250), ("state", "상태", 240)]
        self.tree = ttk.Treeview(box, columns=[c[0] for c in columns], show="headings", selectmode="browse")
        for key, text, width in columns:
            self.tree.heading(key, text=text)
            self.tree.column(key, width=width, anchor="center" if key in ("check", "name") else "w", stretch=key != "check")
        bar = ttk.Scrollbar(box, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=bar.set)
        bar.pack(side="right", fill="y")
        self.tree.pack(side="left", fill="both", expand=True)
        self.tree.bind("<Button-1>", self.toggle)
        self.tree.bind("<Double-1>", lambda e: self.preview() if self.tree.identify_column(e.x) != "#1" else None)

    # 설정
    def pick_service(self):
        host, port, security = SERVICES[self.vars["service"].get()]
        self.vars["host"].set(host)
        self.vars["port"].set(str(port))
        self.vars["security"].set(security)
        self.lock_server()

    def lock_server(self):
        custom = self.vars["service"].get() == "직접 입력"
        for widget in (self.host_entry, self.port_entry, self.security_box):
            widget.configure(state="normal" if custom else "disabled")

    def settings(self):
        cfg = {k: v.get().strip() for k, v in self.vars.items()}
        cfg["subject"] = self.subject.get()
        cfg["body"] = self.body.get("1.0", "end-1c")
        return cfg

    def save_settings(self):
        try:
            self.config_path.write_text(json.dumps(self.settings(), ensure_ascii=False, indent=1), encoding="utf-8")
        except OSError:
            pass

    # 파일
    def choose(self):
        path = filedialog.askopenfilename(title="메일 발송용 zip 고르기", initialdir=str(Path.home() / "Downloads"),
                                          filetypes=[("압축 파일", "*.zip")])
        if path:
            self.load(path)

    def load_latest(self):
        folder = Path.home() / "Downloads"
        found = sorted(folder.glob("급여명세서_*_메일발송*.zip"), key=os.path.getmtime) if folder.is_dir() else []
        if found:
            self.load(str(found[-1]))

    def load(self, path):
        try:
            rows, files = load_package(path)
            sent = read_sent(self.log_path)
        except Exception as error:
            messagebox.showerror("파일을 열 수 없습니다", str(error))
            return
        self.rows, self.files, self.checked = rows, files, set()
        self.path.set(path)
        self.tree.delete(*self.tree.get_children())
        for i, row in enumerate(rows):
            when = sent.get((row["연월"], row["성명"], row["이메일"].lower()))
            if row["error"]:
                state = "보낼 수 없음: " + row["error"]
            elif when:
                state = f"이미 보냄 ({when})"
            elif row["duplicate"]:
                state = "같은 이메일이 또 있음 — 확인 필요"
            else:
                state = "대기"
                self.checked.add(i)
            self.tree.insert("", "end", iid=str(i), values=(self.mark(i), row["성명"], row["이메일"], row["파일"], state))
        months = ", ".join(sorted({r["연월"] for r in rows}))
        self.status.set(f"{months} · {len(rows)}명 중 {len(self.checked)}명 선택")

    def mark(self, i):
        return "－" if self.rows[i]["error"] else "☑" if i in self.checked else "☐"

    def refresh_marks(self):
        for i in range(len(self.rows)):
            self.tree.set(str(i), "check", self.mark(i))
        self.status.set(f"{len(self.rows)}명 중 {len(self.checked)}명 선택")

    def toggle(self, event):
        item = self.tree.identify_row(event.y)
        if self.busy or not item or self.tree.identify_column(event.x) != "#1" or self.rows[int(item)]["error"]:
            return
        self.checked ^= {int(item)}
        self.refresh_marks()

    def select_all(self):
        if not self.busy:
            sent = read_sent(self.log_path)
            self.checked = {i for i, r in enumerate(self.rows) if not r["error"] and not r["duplicate"]
                            and (r["연월"], r["성명"], r["이메일"].lower()) not in sent}
            self.refresh_marks()

    def select_none(self):
        if not self.busy:
            self.checked = set()
            self.refresh_marks()

    def preview(self):
        picked = self.tree.selection()
        if not picked:
            self.status.set("PDF를 볼 줄을 먼저 고르세요.")
            return
        name = self.rows[int(picked[0])]["파일"]
        if name not in self.files:
            return
        self.preview_dir = self.preview_dir or Path(tempfile.mkdtemp(prefix="payslip_preview_"))
        target = self.preview_dir / re.sub(r'[\\/:*?"<>|]', "_", name)
        try:
            target.write_bytes(self.files[name])
            os.startfile(target)
        except OSError as error:
            messagebox.showerror("PDF를 열 수 없습니다", str(error))

    # 발송
    def check_account(self):
        cfg = self.settings()
        if not self.rows:
            return "발송 파일을 먼저 고르세요."
        if not cfg["host"] or not cfg["port"].isdigit():
            return "서버 주소와 포트를 확인하세요."
        if not EMAIL.match(cfg["sender"]):
            return "보내는 주소를 확인하세요."
        if not self.password.get().strip():
            return "앱 비밀번호를 입력하세요."
        return ""

    def secret(self):
        value = self.password.get().strip()
        return value.replace(" ", "") if self.vars["service"].get() == "Gmail" else value

    def send_test(self):
        problem = self.check_account()
        if problem:
            messagebox.showwarning("확인", problem)
            return
        picked = [int(x) for x in self.tree.selection()] or sorted(self.checked)
        picked = [i for i in picked if self.rows[i]["파일"] in self.files]
        if not picked:
            messagebox.showwarning("확인", "시험에 쓸 직원 줄을 하나 고르세요.")
            return
        i, me = picked[0], self.settings()["sender"]
        if messagebox.askyesno("시험 발송", f"{self.rows[i]['성명']}님의 명세서를 내 주소({me})로 보냅니다.\n직원에게는 보내지 않습니다. 보낼까요?"):
            self.start([(i, self.rows[i])], test_to=me)

    def send(self):
        problem = self.check_account()
        if problem:
            messagebox.showwarning("확인", problem)
            return
        jobs = [(i, self.rows[i]) for i in sorted(self.checked)]
        if not jobs:
            messagebox.showwarning("확인", "받는 사람을 선택하세요.")
            return
        try:
            sent = read_sent(self.log_path)
        except OSError as error:
            messagebox.showerror("기록 파일 오류", f"메일발송_기록.csv를 읽을 수 없습니다. 파일을 닫고 다시 하세요.\n{error}")
            return
        lines = []
        for _, r in jobs:
            notes = (" ※이미 보낸 적 있음" if (r["연월"], r["성명"], r["이메일"].lower()) in sent else "") + (" ※같은 이메일 있음" if r["duplicate"] else "")
            lines.append(f"{r['연월']}  {r['성명']} → {r['이메일']}{notes}")
        shown = "\n".join(lines[:25]) + (f"\n… 외 {len(lines) - 25}명" if len(lines) > 25 else "")
        if messagebox.askyesno("발송 확인", f"{len(jobs)}명에게 각자의 급여명세서를 보냅니다.\n이메일 주소를 확인하세요.\n\n{shown}\n\n보낼까요?", default="no"):
            self.start(jobs)

    def start(self, jobs, test_to=None):
        self.busy = True
        for button in (self.send_button, self.test_button, self.open_button):
            button.configure(state="disabled")
        self.status.set("메일 서버에 연결하는 중…")
        self.save_settings()
        cfg, password = self.settings(), self.secret()

        def work():
            try:
                done, failed = send_all(cfg, password, jobs, self.files, self.log_path,
                                        lambda i, text: self.events.put(("row", i, text)), test_to)
                self.events.put(("end", f"{'시험 ' if test_to else ''}발송 완료: 성공 {done}명, 실패 {failed}명", failed, test_to))
            except PermissionError:
                self.events.put(("end", "메일발송_기록.csv에 쓸 수 없습니다. 파일을 닫고 다시 하세요.", 1, test_to))
            except Exception as error:
                self.events.put(("end", explain(error), 1, test_to))

        threading.Thread(target=work, daemon=True).start()

    def poll(self):
        while not self.events.empty():
            event = self.events.get()
            if event[0] == "row":
                _, i, text = event
                self.tree.set(str(i), "state", text)
                self.tree.see(str(i))
                if text == "성공":
                    self.checked.discard(i)
                    self.tree.set(str(i), "check", self.mark(i))
            else:
                _, text, failed, test_to = event
                self.busy = False
                for button in (self.send_button, self.test_button, self.open_button):
                    button.configure(state="normal")
                self.status.set(text)
                (messagebox.showwarning if failed else messagebox.showinfo)(
                    "결과", text + (f"\n받은편지함({test_to})에서 확인하세요." if test_to and not failed else ""))
        self.root.after(100, self.poll)

    def close(self):
        if self.busy:
            messagebox.showwarning("발송 중", "발송이 끝난 뒤 닫아 주세요.")
            return
        self.save_settings()
        if self.preview_dir:
            shutil.rmtree(self.preview_dir, ignore_errors=True)
        self.root.destroy()


def main():
    try:
        import ctypes
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except (ImportError, AttributeError, OSError):
        pass
    root = tk.Tk()
    app = App(root)
    app.load_latest()
    root.mainloop()


if __name__ == "__main__":
    main()
