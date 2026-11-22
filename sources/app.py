import base64
from contextlib import closing
from dataclasses import asdict
from datetime import datetime
import imaplib
import json
import msvcrt
import os
from pathlib import Path
import queue
import re
import sqlite3
import ssl
import threading
import tkinter as tk
from tkinter import messagebox, ttk

import win32crypt
from mailcheck import (
    Account, Cancelled, CheckError, PRESETS, connect, database, scan,
)

APP_DIR = Path(os.environ.get("LOCALAPPDATA", Path.home())) / "SteamMailChecker"
ARCHIVE = APP_DIR / "Letters"
INDEX = APP_DIR / "index.sqlite3"
SETTINGS = APP_DIR / "accounts.json"
INTERVAL = 30


def read_config():
    try:
        return json.loads(SETTINGS.read_text(encoding="utf-8"))
    except FileNotFoundError:
        return []


def save_config(accounts, protected):
    data = [asdict(account) | {"protected": protected.get(account.key, "")} for account in accounts.values()]
    temporary = SETTINGS.with_suffix(".tmp")
    temporary.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
    temporary.replace(SETTINGS)


def protect(secret):
    blob = win32crypt.CryptProtectData(secret.encode(), None, None, None, None, 1)
    return base64.b64encode(blob).decode("ascii")


def unprotect(encoded):
    return win32crypt.CryptUnprotectData(base64.b64decode(encoded), None, None, None, 1)[1].decode()


def monitor(account, secret, stop, wake, out):
    cursors = {}
    client = None
    delay = INTERVAL

    def notice(message, update=False):
        out.put((account.key, message, update))

    try:
        with closing(database(INDEX)) as db:
            while not stop.is_set():
                wake.clear()
                notice("Проверяю папки…")
                try:
                    if client is None:
                        client = connect(account, secret)
                    count, errors = scan(client, db, ARCHIVE, account, cursors, stop)
                    status = f"Онлайн · {datetime.now():%H:%M:%S} · новых: {count}"
                    if errors:
                        status += f" · недоступных папок: {len(errors)}"
                    notice(status, True)
                    delay = INTERVAL
                except Cancelled:
                    break
                except CheckError as error:
                    if client is None:
                        notice(str(error), True)
                        return
                    notice(f"{error} Повтор через {delay} с.", True)
                    try:
                        client.shutdown()
                    except OSError:
                        pass
                    client = None
                    delay = min(delay * 2, 300)
                except (ssl.SSLCertVerificationError, UnicodeError):
                    notice("Ошибка TLS или кодировки логина. Проверьте адрес IMAP и настройки входа.", True)
                    return
                except (OSError, imaplib.IMAP4.error):
                    if client is not None:
                        try:
                            client.shutdown()
                        except OSError:
                            pass
                        client = None
                    notice(f"Связь с IMAP прервана. Повтор через {delay} с.", True)
                    delay = min(delay * 2, 300)
                if stop.is_set():
                    break
                wake.wait(delay)
    except sqlite3.Error:
        notice("Ошибка локальной базы. Проверьте доступ и место на диске.", True)
    except Exception:
        notice("Не удалось обработать ответ IMAP или сохранить письмо.", True)
    finally:
        if client is not None:
            try:
                client.shutdown()
            except OSError:
                pass
        out.put((account.key, None, True))


class MailApp:
    def __init__(self, root):
        self.root = root
        self.accounts = {}
        self.secrets = {}
        self.protected = {}
        self.workers = {}
        self.messages = queue.Queue()
        self.paths = {}
        self.signature = None
        self.oldest_first = False
        self.only_selected = tk.BooleanVar(value=False)
        self.filter = tk.StringVar()
        root.title("Steam Mail Checker")
        root.geometry("1120x790")
        root.minsize(850, 620)

        style = ttk.Style(root)
        if "vista" in style.theme_names():
            style.theme_use("vista")
        style.configure("Title.TLabel", font=("Segoe UI", 21, "bold"))
        style.configure("Treeview", rowheight=24)
        main = ttk.Frame(root, padding=16)
        main.pack(expand=True, fill="both")
        ttk.Label(main, text="Steam Mail Checker", style="Title.TLabel").pack(anchor="w")
        ttk.Label(main, text="Письма Steam из ваших ящиков · проверка каждые 30 секунд").pack(anchor="w")
        controls = ttk.Frame(main)
        controls.pack(fill="x", pady=(10, 5))
        for label, action in (
            ("Добавить ящик", self.edit),
            ("Настроить", lambda: self.edit(self.selected_account())),
            ("Запустить все", self.start_all),
            ("Проверить сейчас", self.check_now),
            ("Остановить все", self.stop_all),
            ("Убрать ящик", self.remove),
        ):
            ttk.Button(controls, text=label, command=action).pack(side="left", padx=(0, 5))
        ttk.Button(controls, text="Папка TXT", command=lambda: os.startfile(ARCHIVE)).pack(side="right")
        self.account_list = self.table(main, 4, (
            ("address", "Ящик", 240), ("host", "Сервер IMAP", 215), ("status", "Состояние", 520),
        ))
        self.account_list.bind("<<TreeviewSelect>>", self.refresh)
        pane = ttk.Panedwindow(main, orient="vertical")
        pane.pack(fill="both", expand=True, pady=(10, 0))
        top = ttk.Frame(pane)
        pane.add(top, weight=2)
        bar = ttk.Frame(top)
        bar.pack(fill="x", pady=(0, 5))
        ttk.Label(bar, text="Фильтр по теме, ящику или отправителю").pack(side="left", padx=(0, 6))
        ttk.Entry(bar, textvariable=self.filter).pack(side="left", fill="x", expand=True)
        ttk.Checkbutton(bar, text="Только выбранный ящик", variable=self.only_selected,
                        command=self.refresh).pack(side="right", padx=(8, 0))
        self.letters = self.table(top, 9, (
            ("date", "Дата письма", 147), ("address", "Ящик", 210),
            ("subject", "Тема", 430), ("folder", "Папка IMAP", 160),
        ))
        self.letters.bind("<<TreeviewSelect>>", self.read_preview)
        self.letters.heading("date", text="Дата письма ↓", command=self.toggle_date)
        self.filter.trace_add("write", lambda *_: self.refresh())

        preview_box = ttk.LabelFrame(pane, text="Текст выбранного письма", padding=5)
        pane.add(preview_box, weight=3)
        sc = ttk.Scrollbar(preview_box)
        sc.pack(side="right", fill="y")
        self.preview = tk.Text(preview_box, wrap="word", undo=False, font=("Segoe UI", 10),
                               state="disabled", yscrollcommand=sc.set)
        self.preview.pack(expand=True, fill="both")
        sc.configure(command=self.preview.yview)
        self.summary = tk.StringVar()
        ttk.Label(main, textvariable=self.summary).pack(anchor="w", pady=(6, 0))
        self.root.protocol("WM_DELETE_WINDOW", self.close)

        try:
            for row in read_config():
                saved = row.pop("protected", "")
                account = Account(**row)
                self.accounts[account.key] = account
                if saved:
                    try:
                        self.secrets[account.key] = unprotect(saved)
                        self.protected[account.key] = saved
                    except Exception:
                        pass
                self.account_list.insert("", "end", iid=account.key, values=(
                    account.address, account.host, "Готов к запуску" if account.key in self.secrets else "Введите пароль",
                ))
        except (ValueError, OSError, TypeError):
            self.summary.set("Не удалось прочитать список ящиков.")
        self.refresh()
        root.after(500, self.tick)

    def table(self, parent, height, columns):
        frame = ttk.Frame(parent)
        frame.pack(fill="both", expand=height > 4)
        view = ttk.Treeview(frame, columns=[col[0] for col in columns], height=height,
                            show="headings", selectmode="browse")
        for key, label, width in columns:
            view.heading(key, text=label)
            view.column(key, width=width, minwidth=80, stretch=key in ("subject", "status", "folder"))
        scroll = ttk.Scrollbar(frame, orient="vertical", command=view.yview)
        view.configure(yscrollcommand=scroll.set)
        view.pack(side="left", expand=True, fill="both")
        scroll.pack(side="right", fill="y")
        return view

    def selected_account(self):
        selected = self.account_list.selection()
        return selected[0] if selected else None

    def edit(self, key=None):
        if key and key in self.workers:
            messagebox.showinfo("Steam Mail Checker", "Остановите проверку перед настройкой ящика.", parent=self.root)
            return
        original = self.accounts.get(key)
        dialog = tk.Toplevel(self.root)
        dialog.title("Настройка ящика")
        dialog.transient(self.root)
        panel = ttk.Frame(dialog, padding=15)
        panel.pack(fill="both", expand=True)
        fields = {}
        address_input = None
        for row, (name, label, initial) in enumerate((
            ("address", "Адрес / логин", original.address if original else ""),
            ("host", "Сервер IMAP", original.host if original else ""),
            ("port", "Порт TLS", str(original.port) if original else "993"),
            ("secret", "Пароль приложения / токен", self.secrets.get(key, "")),
        )):
            ttk.Label(panel, text=label).grid(row=row, column=0, sticky="w", padx=(0, 10), pady=4)
            variable = tk.StringVar(value=initial)
            fields[name] = variable
            entry = ttk.Entry(panel, textvariable=variable, width=39, show="•" if name == "secret" else "")
            entry.grid(row=row, column=1, sticky="ew", pady=4)
            if name == "address":
                entry.focus_set()
                address_input = entry
        oauth = tk.BooleanVar(value=bool(original and original.oauth))
        remember = tk.BooleanVar(value=key in self.protected)
        ttk.Checkbutton(panel, text="Использовать OAuth-токен (XOAUTH2)", variable=oauth).grid(
            row=4, columnspan=2, sticky="w", pady=(8, 1))
        ttk.Checkbutton(panel, text="Запомнить в моём профиле Windows", variable=remember).grid(
            row=5, columnspan=2, sticky="w", pady=(0, 8))
        details = ttk.Label(panel, wraplength=460)
        details.grid(row=6, columnspan=2, sticky="w", pady=(0, 8))

        def autofill(*_):
            domain = fields["address"].get().strip().lower().rpartition("@")[2]
            if not fields["host"].get().strip():
                fields["host"].set(PRESETS.get(domain, ""))
                if PRESETS.get(domain) == "outlook.office365.com":
                    oauth.set(True)

        def apply(start):
            address, host = (fields[name].get().strip() for name in ("address", "host"))
            secret = fields["secret"].get()
            host = host.lower()
            if not host:
                autofill()
                host = fields["host"].get()
            try:
                port = int(fields["port"].get())
                assert 1 <= port <= 65535
            except (ValueError, AssertionError):
                details.configure(text="Введите порт от 1 до 65535.")
                return
            if (not address or not re.fullmatch(r"[A-Za-z0-9][A-Za-z0-9.-]*", host)
                    or any(ord(c) < 32 for c in address + secret)):
                details.configure(text="Введите логин и имя сервера IMAP без пробелов и управляющих символов.")
                return
            if start and not secret:
                details.configure(text="Введите пароль приложения или токен.")
                return
            account = Account(address, host, port, oauth.get())
            if account.key in self.accounts and account.key != key:
                details.configure(text="Ящик уже добавлен. Откройте его через «Настроить».")
                return
            saved = ""
            try:
                if remember.get() and secret:
                    saved = protect(secret)
                new_accounts = {k: v for k, v in self.accounts.items() if k != key} | {account.key: account}
                new_protected = {k: v for k, v in self.protected.items() if k != key and k != account.key}
                if saved:
                    new_protected[account.key] = saved
                save_config(new_accounts, new_protected)
            except Exception:
                details.configure(text="Не удалось сохранить настройки в профиле Windows.")
                return
            self.accounts, self.protected = new_accounts, new_protected
            self.secrets.pop(key, None)
            if secret:
                self.secrets[account.key] = secret
            if key:
                self.account_list.delete(key)
            self.account_list.insert("", "end", iid=account.key,
                                     values=(address, host, "Готов к запуску" if secret else "Введите пароль"))
            self.account_list.selection_set(account.key)
            dialog.destroy()
            if start:
                self.start(account.key)

        dialog.bind("<Return>", lambda _: apply(True))
        dialog.bind("<Escape>", lambda _: dialog.destroy())
        address_input.bind("<FocusOut>", autofill)
        fields["address"].trace_add("write", lambda *_: details.configure(text=""))
        ttk.Button(panel, text="Сохранить", command=lambda: apply(False)).grid(row=7, column=0, sticky="w")
        ttk.Button(panel, text="Сохранить и запустить", command=lambda: apply(True)).grid(row=7, column=1, sticky="e")
        dialog.grab_set()

    def start(self, key):
        if key in self.workers:
            return
        if not self.secrets.get(key):
            self.account_list.set(key, "status", "Введите пароль или токен через «Настроить»")
            return
        stop = threading.Event()
        wake = threading.Event()
        thread = threading.Thread(target=monitor,
                                  args=(self.accounts[key], self.secrets[key], stop, wake, self.messages), daemon=True)
        self.workers[key] = (stop, wake, thread)
        thread.start()

    def start_all(self):
        for key in self.accounts:
            self.start(key)

    def check_now(self):
        for key in self.accounts:
            if key in self.workers:
                self.workers[key][1].set()
            else:
                self.start(key)

    def stop_all(self):
        for key, (stop, wake, _) in self.workers.items():
            stop.set()
            wake.set()
            self.account_list.set(key, "status", "Останавливаю…")

    def remove(self):
        key = self.selected_account()
        if not key:
            return
        if key in self.workers:
            messagebox.showinfo("Steam Mail Checker", "Сначала остановите проверку.", parent=self.root)
            return
        accounts = {k: v for k, v in self.accounts.items() if k != key}
        protected = {k: v for k, v in self.protected.items() if k != key}
        try:
            save_config(accounts, protected)
        except OSError:
            messagebox.showerror("Steam Mail Checker", "Не удалось сохранить список ящиков.", parent=self.root)
            return
        self.accounts, self.protected = accounts, protected
        self.secrets.pop(key, None)
        self.account_list.delete(key)

    def toggle_date(self):
        self.oldest_first = not self.oldest_first
        self.letters.heading("date", text="Дата письма ↑" if self.oldest_first else "Дата письма ↓")
        self.refresh()

    def refresh(self, _=None):
        try:
            with closing(sqlite3.connect(INDEX, timeout=0.2)) as db:
                rows = db.execute("""SELECT id,account,address,folder,stamp,sender,subject,path
                    FROM letters ORDER BY stamp DESC,id DESC""").fetchall()
        except sqlite3.Error:
            return
        selected = self.selected_account() if self.only_selected.get() else None
        needle = self.filter.get().casefold().strip()
        visible = [r for r in rows if (not selected or selected == r[1])
                   and (not needle or needle in " ".join((r[2], r[5], r[6])).casefold())]
        if self.oldest_first:
            visible.sort(key=lambda row: (not row[4], row[4], row[0]))
        current = tuple(row[0] for row in visible)
        if self.signature != current:
            self.signature = current
            old = self.letters.selection()
            self.letters.delete(*self.letters.get_children())
            self.paths.clear()
            for row_id, _, addr, folder, stamp, sender, subject, path in visible:
                iid = str(row_id)
                self.paths[iid] = ARCHIVE / path
                try:
                    date = datetime.fromtimestamp(stamp).strftime("%d.%m.%Y %H:%M") if stamp else "—"
                except (ValueError, OSError, OverflowError):
                    date = "—"
                self.letters.insert("", "end", iid=iid, values=(date, addr, subject, folder))
            if old and self.letters.exists(old[0]):
                self.letters.selection_set(old[0])
            self.read_preview()
        self.summary.set(f"Архив: {len(rows)} писем · показано: {len(visible)}")

    def read_preview(self, _=None):
        selection = self.letters.selection()
        row = selection[0] if selection else None
        try:
            content = self.paths[row].read_text(encoding="utf-8") if row else ""
        except (OSError, KeyError, UnicodeError):
            content = "Локальный TXT недоступен."
        self.preview.configure(state="normal")
        self.preview.delete("1.0", "end")
        self.preview.insert("1.0", content)
        self.preview.configure(state="disabled")

    def tick(self):
        update = False
        while True:
            try:
                key, status, changed = self.messages.get_nowait()
            except queue.Empty:
                break
            update |= changed
            if status is None:
                worker = self.workers.pop(key, None)
                if worker and worker[0].is_set() and self.account_list.exists(key):
                    self.account_list.set(key, "status", "Остановлено")
            elif self.account_list.exists(key):
                self.account_list.set(key, "status", status)
        if update:
            self.refresh()
        self.root.after(500, self.tick)

    def close(self):
        self.stop_all()
        self.root.destroy()


def main():
    APP_DIR.mkdir(parents=True, exist_ok=True)
    ARCHIVE.mkdir(exist_ok=True)
    root = tk.Tk()
    root.withdraw()
    with (APP_DIR / "session.lock").open("a+b") as lock:
        try:
            lock.seek(0)
            msvcrt.locking(lock.fileno(), msvcrt.LK_NBLCK, 1)
        except OSError:
            messagebox.showinfo("Steam Mail Checker", "Чекер уже запущен.", parent=root)
            root.destroy()
            return
        with closing(database(INDEX)):
            pass
        MailApp(root)
        root.deiconify()
        root.mainloop()


if __name__ == "__main__":
    main()
