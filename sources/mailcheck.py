from dataclasses import dataclass
from datetime import datetime, timezone
from email import policy
from email.parser import BytesParser
from email.utils import getaddresses, parsedate_to_datetime
from html.parser import HTMLParser
import hashlib
import imaplib
import re
import sqlite3
import ssl
from pathlib import Path

DOMAINS = ("steampowered.com", "steamcommunity.com")
PRESETS = {
    **dict.fromkeys(("gmail.com", "googlemail.com"), "imap.gmail.com"),
    **dict.fromkeys(("yandex.ru", "ya.ru", "yandex.com", "yandex.kz", "yandex.by"), "imap.yandex.ru"),
    **dict.fromkeys(("mail.ru", "inbox.ru", "list.ru", "bk.ru", "internet.ru"), "imap.mail.ru"),
    **dict.fromkeys(("outlook.com", "hotmail.com", "live.com", "msn.com"), "outlook.office365.com"),
    **dict.fromkeys(("icloud.com", "me.com", "mac.com"), "imap.mail.me.com"),
    "yahoo.com": "imap.mail.yahoo.com",
    "rambler.ru": "imap.rambler.ru",
}
LIST = re.compile(rb'^\(([^)]*)\)\s+(?:"(?:[^"\\]|\\.)*"|NIL)\s+(.+)$')
UID = re.compile(rb"\bUID (\d+)\b")
INTERNALDATE = re.compile(rb'INTERNALDATE "([^"]+)"')
INVALID_PATH = re.compile(r'[<>:"/\\|?*\x00-\x1f]')


@dataclass(frozen=True)
class Account:
    address: str
    host: str
    port: int = 993
    oauth: bool = False

    @property
    def key(self):
        return hashlib.sha256(f"{self.host.lower()}:{self.port}/{self.address.lower()}".encode()).hexdigest()[:16]


class CheckError(Exception):
    pass


class Cancelled(Exception):
    pass


class TextHTML(HTMLParser):
    blocks = {"p", "br", "div", "tr", "li", "h1", "h2", "h3", "section", "blockquote"}

    def __init__(self):
        super().__init__(convert_charrefs=True)
        self.hidden = 0
        self.chunks = []
        self.links = []

    def handle_starttag(self, tag, attrs):
        if tag in {"script", "style", "head"}:
            self.hidden += 1
        if not self.hidden and tag in self.blocks:
            self.chunks.append("\n")
        if not self.hidden and tag == "a":
            href = dict(attrs).get("href", "")
            if href and href.startswith(("https://", "http://")):
                self.links.append(href)

    def handle_endtag(self, tag):
        if tag in {"script", "style", "head"}:
            self.hidden = max(0, self.hidden - 1)
        if not self.hidden and tag in self.blocks:
            self.chunks.append("\n")

    def handle_data(self, data):
        if not self.hidden:
            self.chunks.append(re.sub(r"\s+", " ", data))

    def text(self):
        lines = [line.strip() for line in "".join(self.chunks).splitlines()]
        return "\n".join(filter(None, lines)) + (
            "\n\nСсылки из HTML:\n" + "\n".join(dict.fromkeys(self.links)) if self.links else ""
        )


def connect(account, secret):
    client = imaplib.IMAP4_SSL(account.host, account.port, ssl_context=ssl.create_default_context(), timeout=18)
    client.sock.settimeout(22)
    try:
        if account.oauth:
            client.authenticate("XOAUTH2", lambda challenge: (
                f"user={account.address}\x01auth=Bearer {secret}\x01\x01".encode() if not challenge else b""
            ))
        else:
            login = '"' + account.address.replace('\\', '\\\\').replace('"', '\\"') + '"'
            client.login(login, secret)
    except imaplib.IMAP4.error:
        client.shutdown()
        raise CheckError("Вход отклонён. Проверьте настройки IMAP и пароль приложения или OAuth-токен.") from None
    except BaseException:
        client.shutdown()
        raise
    return client


def checked(result):
    status, payload = result
    if status != "OK":
        raise CheckError("Сервер IMAP отклонил команду.")
    return payload


def folders(client):
    for entry in checked(client.list()):
        if not entry:
            continue
        if isinstance(entry, tuple):
            header, name = entry
            entry = header.rsplit(b" ", 1)[0] + b' "' + name.replace(b"\\", b"\\\\").replace(b'"', b'\\"') + b'"'
        match = LIST.fullmatch(entry)
        if not match or b"\\noselect" in match[1].lower().split() or b"\\nonexistent" in match[1].lower().split():
            continue
        raw = match[2]
        yield raw if raw.startswith(b'"') else b'"' + raw + b'"'


def readable_folder(raw):
    raw = re.sub(rb'\\(.)', rb'\1', raw[1:-1])
    import base64

    def utf7(match):
        chars = match[1].replace(b",", b"/")
        if not chars:
            return "&"
        try:
            return base64.b64decode(chars + b"=" * (-len(chars) % 4)).decode("utf-16-be")
        except ValueError:
            return match[0].decode("ascii")

    return "".join(
        utf7(match) if match[0].startswith(b"&") and match[1] is not None
        else match[0].decode("utf-8", "replace")
        for match in re.finditer(rb"&([A-Za-z0-9+,]*)-|[^&]+|&", raw)
    )


def server_number(client, item):
    return int(next((value for value in client.response(item)[1] if value and value.isdigit()), b"0"))


def steam_sender(message):
    senders = getaddresses([str(value) for value in message.get_all("From", [])], strict=False)
    for _, address in senders:
        _, at, host = address.lower().rpartition("@")
        if at and any(host == domain or host.endswith("." + domain) for domain in DOMAINS):
            return address
    return None


def clean(value):
    return str(value or "").replace("\r", " ").replace("\n", " ")


def get_body(message):
    part = message.get_body(preferencelist=("plain", "html"))
    if part is None:
        return "(Текстовая часть письма отсутствует.)"
    try:
        body = part.get_content(errors="replace")
    except (LookupError, UnicodeError, KeyError):
        data = part.get_payload(decode=True) or b""
        body = data.decode("utf-8", "replace")
    if part.get_content_subtype() == "html":
        parser = TextHTML()
        parser.feed(body)
        return parser.text().strip()
    return body.strip()


def mail_time(message, meta):
    try:
        date = parsedate_to_datetime(str(message.get("Date")))
        return int((date if date.tzinfo else date.replace(tzinfo=timezone.utc)).timestamp())
    except (ValueError, TypeError, OverflowError, OSError):
        pass
    recorded = INTERNALDATE.search(meta)
    if recorded:
        try:
            return int(datetime.strptime(recorded[1].decode("ascii").strip(), "%d-%b-%Y %H:%M:%S %z").timestamp())
        except ValueError:
            pass
    return 0


def database(path):
    db = sqlite3.connect(path, timeout=10)
    db.execute("PRAGMA journal_mode=WAL")
    db.execute("""CREATE TABLE IF NOT EXISTS letters (
        id INTEGER PRIMARY KEY, account TEXT NOT NULL, message_key TEXT NOT NULL,
        address TEXT NOT NULL, folder TEXT NOT NULL, stamp INTEGER NOT NULL,
        sender TEXT NOT NULL, subject TEXT NOT NULL, path TEXT NOT NULL,
        UNIQUE(account, message_key)
    )""")
    return db


def store(db, archive, account, folder, meta, raw):
    message = BytesParser(policy=policy.default).parsebytes(raw)
    sender = steam_sender(message)
    if sender is None:
        return False
    subject = clean(message.get("Subject")) or "(Без темы)"
    stamp = mail_time(message, meta)
    body = get_body(message)
    mail_id = clean(message.get("Message-ID")).strip()
    dedup = mail_id or f"{sender}\n{subject}\n{stamp}\n{body}"
    key = hashlib.sha256(dedup.encode("utf-8", "replace")).hexdigest()
    existing = db.execute("SELECT path FROM letters WHERE account=? AND message_key=?", (account.key, key)).fetchone()
    if existing and (archive / existing[0]).is_file():
        return False
    date = datetime.fromtimestamp(stamp, timezone.utc).strftime("%Y-%m-%d_%H-%M-%S") if stamp else "без-даты"
    account_name = INVALID_PATH.sub("_", account.address)[:55].rstrip(". ")
    folder_path = archive / f"mail_{account_name}_{account.key[:8]}"
    folder_path.mkdir(parents=True, exist_ok=True)
    path = archive / existing[0] if existing else folder_path / f"{date}_{key[:20]}.txt"
    path.parent.mkdir(parents=True, exist_ok=True)
    data = f"Ящик: {account.address}\nПапка IMAP: {folder}\nОт: {sender}\nКому: {clean(message.get('To'))}\nДата: {clean(message.get('Date')) or date}\nТема: {subject}\n\n{body}\n"
    temp = path.with_suffix(".tmp")
    temp.write_text(data, encoding="utf-8")
    temp.replace(path)
    with db:
        db.execute("""INSERT OR IGNORE INTO letters
            (account,message_key,address,folder,stamp,sender,subject,path) VALUES (?,?,?,?,?,?,?,?)""",
            (account.key, key, account.address, folder, stamp, sender, subject, str(path.relative_to(archive))))
    return True


def scan(client, db, archive, account, cursors, stop):
    new = 0
    errors = []
    for folder in dict.fromkeys(folders(client)):
        if stop.is_set():
            raise Cancelled()
        name = readable_folder(folder)
        result = client.select(folder, readonly=True)
        if result[0] != "OK":
            errors.append(name)
            continue
        validity = server_number(client, "UIDVALIDITY")
        upper = server_number(client, "UIDNEXT") - 1
        previous_validity, last = cursors.get(folder, (0, 0))
        if not validity or previous_validity != validity:
            last = 0
        if upper >= 0 and last >= upper:
            continue
        seqset = f"{last + 1}:{upper if upper >= 0 else '*'}"
        matches = checked(client.uid("SEARCH", None, "UID", seqset, "OR",
                                     "FROM", f'"{DOMAINS[0]}"', "FROM", f'"{DOMAINS[1]}"'))
        uids = sorted({int(uid) for row in matches if row for uid in row.split() if uid.isdigit() and int(uid) > last})
        for offset in range(0, len(uids), 10):
            if stop.is_set():
                raise Cancelled()
            chunk = uids[offset:offset + 10]
            data = checked(client.uid("FETCH", ",".join(map(str, chunk)), "(UID INTERNALDATE BODY.PEEK[])"))
            for entry in data:
                if isinstance(entry, tuple) and UID.search(entry[0]):
                    new += store(db, archive, account, name, *entry)
        if validity:
            cursors[folder] = (validity, upper if upper >= 0 else max(uids, default=last))
    return new, errors
