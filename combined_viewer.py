# Filename: combined_viewer.py
# Location: ./combined_viewer.py
# Summary: PyQt6 email viewer that pulls the newest matching Outlook email,
#          renders HTML (QWebEngine), extracts DOM {visible_text, links},
#          scores/selects a primary portal link, shows JSON in UI, and chats with an LLM via Ollama.
# Caller: Run directly: `python combined_viewer.py`
# References: PyQt6, win32com.client (Outlook), Ollama local API

import sys, json, traceback
from concurrent.futures import ThreadPoolExecutor
from typing import Dict, Any, List, Optional

import win32com.client
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QTextBrowser, QListWidget, QListWidgetItem, QGroupBox,
    QSplitter, QCheckBox, QLineEdit, QPushButton, QMessageBox
)
from PyQt6.QtWebEngineWidgets import QWebEngineView
from PyQt6.QtCore import QTimer, Qt, QMetaObject, Q_ARG

# --------- Config ---------

def load_config(path: str = "config.json") -> Dict[str, Any]:
    with open(path, "r", encoding="utf-8") as f:
        cfg = json.load(f)
    # sensible defaults
    cfg.setdefault("account", "Commercial Estimator")
    cfg.setdefault("sender_domain", "planhub.com")
    cfg.setdefault("senders", [])  # explicit allow-list overrides domain
    cfg.setdefault("ollama_host", "http://localhost:11434")
    cfg.setdefault("model", "gemma2:9b-instruct-q4")
    cfg.setdefault("dom_delay_ms", 800)
    return cfg

CONFIG = load_config()

# --------- Outlook helpers ---------

def find_store(namespace, display_name: str):
    for f in namespace.Folders:
        if f.Name.strip().lower() == display_name.strip().lower():
            return f
    return None

def sender_matches(smtp: str, cfg: Dict[str, Any]) -> bool:
    smtp_l = (smtp or "").lower()
    allow = [s.lower() for s in cfg.get("senders", [])]
    domain = (cfg.get("sender_domain") or "").lower()
    if allow and smtp_l in allow:
        return True
    return bool(domain and domain in smtp_l)

def get_newest_matching_email_html(cfg: Dict[str, Any]) -> tuple[Dict[str, Any], str]:
    """Return (pointer, html_or_text) for the newest email in cfg['account'] Inbox matching sender rules."""
    outlook = win32com.client.Dispatch("Outlook.Application")
    ns = outlook.GetNamespace("MAPI")

    store = find_store(ns, cfg["account"])
    if not store:
        raise RuntimeError(f"Outlook store not found: {cfg['account']}")

    inbox = store.Folders["Inbox"]
    items = inbox.Items
    items.Sort("[ReceivedTime]", True)

    for itm in items:
        smtp = ""
        try:
            # Most robust path in Exchange environments
            exu = itm.Sender.GetExchangeUser()
            smtp = (exu.PrimarySmtpAddress if exu else "") or ""
        except Exception:
            pass
        if not smtp:
            try:
                smtp = itm.SenderEmailAddress or ""
            except Exception:
                smtp = ""

        if sender_matches(smtp, cfg):
            ptr = {
                "account": cfg["account"],
                "mailbox": "Inbox",
                "message_id": getattr(itm, "EntryID", None),  # Outlook pointer
                "subject": getattr(itm, "Subject", ""),
                "from": smtp.lower(),
            }
            html = getattr(itm, "HTMLBody", "") or getattr(itm, "Body", "")
            return ptr, html

    raise RuntimeError(
        f"No matching email found in '{cfg['account']}' inbox "
        f"(senders={cfg.get('senders') or '[]'}, domain='{cfg.get('sender_domain')}')"
    )

# --------- Link scoring ---------

WHITELIST_DOMAINS = [
    "planhub.com", "buildingconnected.com", "constructconnect.com",
    "isqft.com", "procore.com"
]
INTENT_WORDS = ["view project", "submit", "bid", "open invite", "project", "plans", "portal", "itb"]
HREF_HINTS = ["project", "bid", "invite", "itb", "plan", "rfi"]
NEGATIVE_HINTS = ["unsubscribe", "preferences", "kb.", "knowledge", "support", "terms", "privacy"]

def score_link(href: str, text: str) -> float:
    href_l, text_l = (href or "").lower(), (text or "").lower()
    score = 0.0
    if not href_l:
        return -999.0
    for d in WHITELIST_DOMAINS:
        if d in href_l:
            score += 0.6
    for w in INTENT_WORDS:
        if w in text_l:
            score += 0.3
    if href_l.count("/") >= 4:
        score += 0.1
    for q in HREF_HINTS:
        if q in href_l:
            score += 0.1
    for bad in NEGATIVE_HINTS:
        if bad in href_l:
            score -= 1.0
    return score

def rank_links(links: List[Dict[str, str]]) -> List[Dict[str, Any]]:
    ranked = []
    for l in links:
        href = l.get("href", "")
        text = l.get("text", "")
        ranked.append({"text": text, "href": href, "score": score_link(href, text)})
    ranked.sort(key=lambda r: r["score"], reverse=True)
    return ranked

# --------- Ollama (LLM) ---------

def call_ollama(cfg: Dict[str, Any], message: str) -> str:
    # Lazy import to avoid import cost if unused
    import ollama
    client = ollama.Client(host=cfg["ollama_host"])
    resp = client.chat(model=cfg["model"], messages=[{"role": "user", "content": message}])
    return resp["message"]["content"].strip()

# --------- Main Window ---------

class CombinedViewer(QMainWindow):
    def __init__(self, html: str, ptr: Dict[str, Any], cfg: Dict[str, Any]):
        super().__init__()
        self.cfg = cfg
        self.ptr = ptr
        self.html = html
        self.email_data: Optional[Dict[str, Any]] = None
        self.pool = ThreadPoolExecutor(max_workers=2)

        self.setWindowTitle(ptr.get("subject", "Email Viewer"))
        self.resize(1600, 900)
        self._setup_ui()
        self._load_email()

    # ----- UI setup -----

    def _setup_ui(self):
        cw = QWidget()
        self.setCentralWidget(cw)
        main = QHBoxLayout(cw)

        splitter = QSplitter(Qt.Orientation.Horizontal)
        main.addWidget(splitter)

        # Left: Web view
        left = QWidget()
        left_layout = QVBoxLayout(left)
        self.view = QWebEngineView()
        left_layout.addWidget(self.view)
        splitter.addWidget(left)

        # Middle: Data panels
        middle = QWidget()
        mid_layout = QVBoxLayout(middle)

        self.details_group = QGroupBox("Email Details")
        self.details_group.setStyleSheet("background-color:#fff;")
        self._details_layout = QVBoxLayout(self.details_group)
        mid_layout.addWidget(self.details_group)

        self.body_group = QGroupBox("Email Body (Visible Text)")
        body_layout = QVBoxLayout(self.body_group)
        self.body_browser = QTextBrowser()
        self.body_browser.setStyleSheet(
            "background:#fff;color:#000;border:1px solid #ccc;font-family:Consolas,monospace;font-size:12px;"
        )
        body_layout.addWidget(self.body_browser)
        mid_layout.addWidget(self.body_group, 1)

        self.links_group = QGroupBox("Links")
        links_layout = QVBoxLayout(self.links_group)
        self.links_list = QListWidget()
        self.links_list.setStyleSheet(
            "background:#fff;color:#000;border:1px solid #ccc;font-family:Consolas,monospace;font-size:12px;"
        )
        links_layout.addWidget(self.links_list)
        mid_layout.addWidget(self.links_group, 1)

        splitter.addWidget(middle)

        # Right: LLM panel
        right = QWidget()
        right_layout = QVBoxLayout(right)
        right.setStyleSheet("background:#fff;")

        title = QLabel("🚀 AI Email Decoder (Ollama)")
        title.setStyleSheet("font-size:18px;font-weight:bold;color:#000;")
        right_layout.addWidget(title)

        opts = QGroupBox("Include in Prompt")
        opts.setStyleSheet("QGroupBox{background:#fff;border:1px solid #ccc;}")
        opts_layout = QVBoxLayout(opts)
        self.header_check = QCheckBox("Message Header")
        self.body_check = QCheckBox("Message Body")
        self.links_check = QCheckBox("Links")
        for w in (self.header_check, self.body_check, self.links_check):
            w.setStyleSheet("color:#000;background:#fff;")
            opts_layout.addWidget(w)
        right_layout.addWidget(opts)

        self.chat_display = QTextBrowser()
        self.chat_display.setStyleSheet("background:#f9f9f9;color:#000;border:1px solid #ccc;")
        right_layout.addWidget(self.chat_display, 1)

        input_row = QHBoxLayout()
        self.chat_input = QLineEdit()
        self.chat_input.setPlaceholderText("Ask the model about this email…")
        self.chat_input.setStyleSheet("background:#fff;color:#000;border:1px solid #ccc;padding:4px;")
        self.send_btn = QPushButton("Send")
        self.copy_btn = QPushButton("Copy JSON")
        self.send_btn.setStyleSheet("background:#e0e0e0;color:#000;border:1px solid #ccc;padding:4px;")
        self.copy_btn.setStyleSheet("background:#e0e0e0;color:#000;border:1px solid #ccc;padding:4px;")
        self.send_btn.clicked.connect(self._send_to_llm)
        self.copy_btn.clicked.connect(self._copy_json)
        input_row.addWidget(self.chat_input)
        input_row.addWidget(self.send_btn)
        input_row.addWidget(self.copy_btn)
        right_layout.addLayout(input_row)

        splitter.addWidget(right)
        splitter.setSizes([533, 534, 533])  # roughly equal thirds

    # ----- Loading & DOM extraction -----

    def _load_email(self):
        self.view.setHtml(self.html)
        QTimer.singleShot(self.cfg.get("dom_delay_ms", 800), self._extract_dom)

    def _extract_dom(self):
        js = """
        (() => {
          const isVisible = (el) => {
            const s = window.getComputedStyle(el);
            if (s.display === 'none' || s.visibility === 'hidden' || parseFloat(s.opacity||'1') === 0) return false;
            const r = el.getBoundingClientRect();
            return r.width > 0 && r.height > 0;
          };
          const links = Array.from(document.querySelectorAll('a'))
            .filter(a => a.href && !a.href.startsWith('mailto:') && !a.href.startsWith('tel:') && isVisible(a))
            .map(a => ({ text: (a.innerText||'').trim(), href: a.href }));
          const text = (document.body && document.body.innerText) ? document.body.innerText : '';
          return { text, links };
        })();
        """
        self.view.page().runJavaScript(js, self._on_dom)

    def _on_dom(self, dom):
        try:
            visible_text = (dom or {}).get("text", "") or ""
            links = (dom or {}).get("links", []) or []

            # Rank & choose primary portal
            ranked = rank_links(links)
            primary = ranked[0]["href"] if ranked and ranked[0]["score"] > 0.3 else None

            self.email_data = {
                "email_ptr": self.ptr,
                "dom": {
                    "visible_text": visible_text,
                    "links": links  # full hrefs preserved
                },
                "links": {
                    "primary_portal": primary,
                    "aux": [{"text": r["text"], "href": r["href"]} for r in ranked[1:6]]
                }
            }

            # Persist a working file for downstream tools
            with open("email_output.json", "w", encoding="utf-8") as f:
                json.dump(self.email_data, f, indent=2)

            # Update UI
            self._display_details(self.ptr, self.email_data["links"].get("primary_portal"))
            self._display_body(visible_text)
            self._display_links(links)

        except Exception as e:
            QMessageBox.critical(self, "DOM Error", f"{e}\n\n{traceback.format_exc()}")

    # ----- UI helpers -----

    def _display_details(self, ptr: Dict[str, Any], primary: Optional[str]):
        # Clear old
        while self._details_layout.count():
            item = self._details_layout.takeAt(0)
            w = item.widget()
            if w:
                w.deleteLater()

        # Pointer fields
        for k, v in ptr.items():
            lbl = QLabel(f"{k.replace('_',' ').title()}: {v}")
            lbl.setTextInteractionFlags(Qt.TextInteractionFlag.TextSelectableByMouse)
            lbl.setWordWrap(True)
            lbl.setMaximumWidth(500)
            self._details_layout.addWidget(lbl)

        # Primary portal
        if primary:
            p = QLabel(f"Primary Portal: {primary}")
            p.setTextInteractionFlags(Qt.TextInteractionFlag.TextSelectableByMouse)
            p.setWordWrap(True)
            p.setMaximumWidth(500)
            self._details_layout.addWidget(p)

    def _display_body(self, text: str):
        self.body_browser.setPlainText(text)

    def _display_links(self, links: List[Dict[str, str]]):
        self.links_list.clear()
        for l in links:
            text = l.get("text", "") or ""
            href = l.get("href", "") or ""
            ui_href = href if len(href) <= 100 else href[:100] + "…"
            item = QListWidgetItem(f"{text} -> {ui_href}")
            item.setToolTip(href)  # full URL on hover
            self.links_list.addItem(item)

    # ----- LLM chat -----

    def _compose_context(self) -> str:
        parts = []
        if self.header_check.isChecked():
            parts.append("Header:\n" + json.dumps(self.email_data.get("email_ptr", {}), indent=2))
        if self.body_check.isChecked():
            parts.append("Body:\n" + (self.email_data.get("dom", {}).get("visible_text") or ""))
        if self.links_check.isChecked():
            parts.append("Links:\n" + json.dumps(self.email_data.get("dom", {}).get("links", []), indent=2))
        return "\n\n".join(parts)

    def _send_to_llm(self):
        if not self.email_data:
            self.chat_display.append("No email data yet.")
            return
        user_msg = self.chat_input.text().strip()
        if not user_msg:
            return

        context = self._compose_context()
        full_prompt = (
            f"{context}\n\n"
            "You are a bid-invite extractor. Answer the user precisely.\n"
            "If asked, extract fields as JSON using keys: project_name, address, zip, due_date, gc_name, contacts[], links.primary.\n"
        ) + f"\nUser: {user_msg}"

        self.chat_display.append(f"You: {user_msg}")
        self.chat_input.clear()
        self.chat_display.append("Model: …")

        def task():
            try:
                return call_ollama(self.cfg, full_prompt)
            except Exception as e:
                return f"[ERROR] {e}"

        fut = self.pool.submit(task)

        def on_done():
            try:
                txt = fut.result()
            except Exception as e:
                txt = f"[ERROR] {e}"
            # marshal result back to UI thread
            QMetaObject.invokeMethod(
                self.chat_display,
                "append",
                Qt.ConnectionType.QueuedConnection,
                Q_ARG(str, f"Model: {txt}")
            )

        # finish callback (in UI thread)
        QTimer.singleShot(0, on_done)

    def _copy_json(self):
        if not self.email_data:
            return
        QApplication.clipboard().setText(json.dumps(self.email_data, indent=2))
        self.chat_display.append("✓ JSON copied to clipboard.")

# --------- Entry point ---------

def main():
    try:
        ptr, html = get_newest_matching_email_html(CONFIG)
    except Exception as e:
        print(f"Outlook error: {e}")
        print(traceback.format_exc())
        sys.exit(1)

    app = QApplication(sys.argv)
    win = CombinedViewer(html, ptr, CONFIG)
    win.show()
    sys.exit(app.exec())

if __name__ == "__main__":
    main()
