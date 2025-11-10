# Filename: combined_viewer.py
# Location: ./combined_viewer.py
# Summary: Combined PyQt6 application showing email viewer on left and JSON data on right
# Dependencies: PyQt6, json, win32com.client

import sys, json, win32com.client
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QTextBrowser, QListWidget, QListWidgetItem, QGroupBox,
    QSplitter, QCheckBox, QLineEdit, QPushButton
)
from PyQt6.QtWebEngineWidgets import QWebEngineView
from PyQt6.QtCore import QTimer, Qt
import ollama

# Load configuration
with open('config.json', 'r') as f:
    config = json.load(f)

COMMERCIAL_ESTIMATOR_DISPLAY = config["account"]

def get_latest_planhub_email_html():
    outlook = win32com.client.Dispatch("Outlook.Application")
    ns = outlook.GetNamespace("MAPI")

    # Find the "Commercial Estimator" store
    store = None
    for f in ns.Folders:
        if f.Name.strip().lower() == COMMERCIAL_ESTIMATOR_DISPLAY.lower():
            store = f
            break
    if not store:
        raise RuntimeError("Commercial Estimator account not found in Outlook.Folders")

    inbox = store.Folders["Inbox"]

    # Walk newest→oldest and pick first from planhub.com
    items = inbox.Items
    items.Sort("[ReceivedTime]", True)

    for itm in items:
        # Try to get SMTP address robustly
        smtp = ""
        try:
            smtp = (itm.Sender.GetExchangeUser().PrimarySmtpAddress or "").lower()
        except:
            pass
        if not smtp:
            try:
                smtp = (itm.SenderEmailAddress or "").lower()
            except:
                smtp = ""

        if smtp == config["sender"]:
            # Pointer fields you care about
            ptr = {
                "account": COMMERCIAL_ESTIMATOR_DISPLAY,
                "mailbox": "Inbox",
                "message_id": getattr(itm, "EntryID", None),  # Outlook EntryID (useful pointer)
                "subject": getattr(itm, "Subject", ""),
                "from": smtp
            }
            html = getattr(itm, "HTMLBody", "") or getattr(itm, "Body", "")
            return ptr, html

    raise RuntimeError(f"No matching {config['sender']} email found in {config['account']} inbox")

class CombinedViewer(QMainWindow):
    def __init__(self, html, ptr):
        super().__init__()
        self.ptr = ptr
        self.html = html
        self.email_data = None
        self.setWindowTitle(ptr.get("subject", "Email Viewer"))
        self.setGeometry(100, 100, 1600, 900)
        self.setup_ui()
        self.load_email()

    def setup_ui(self):
        """Setup the main UI with splitter."""
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QHBoxLayout(central_widget)

        # Splitter for left, middle, right
        splitter = QSplitter(Qt.Orientation.Horizontal)
        main_layout.addWidget(splitter)

        # Left side: Email viewer
        left_widget = QWidget()
        left_layout = QVBoxLayout(left_widget)
        self.view = QWebEngineView()
        left_layout.addWidget(self.view)
        splitter.addWidget(left_widget)

        # Middle: JSON data display
        middle_widget = QWidget()
        middle_layout = QVBoxLayout(middle_widget)

        # Email details section
        self.details_group = QGroupBox("Email Details")
        self.details_group.setStyleSheet("background-color: #ffffff;")
        details_layout = QVBoxLayout()
        self.details_group.setLayout(details_layout)
        middle_layout.addWidget(self.details_group)

        # Email body section
        self.body_group = QGroupBox("Email Body")
        body_layout = QVBoxLayout()
        self.body_browser = QTextBrowser()
        self.body_browser.setStyleSheet("""
            background-color: #ffffff;
            color: #000000;
            border: 1px solid #ccc;
            font-family: Arial, sans-serif;
            font-size: 12px;
        """)
        body_layout.addWidget(self.body_browser)
        self.body_group.setLayout(body_layout)
        middle_layout.addWidget(self.body_group, 1)  # Expand

        # Links section
        self.links_group = QGroupBox("Links")
        links_layout = QVBoxLayout()
        self.links_list = QListWidget()
        self.links_list.setStyleSheet("""
            background-color: #ffffff;
            color: #000000;
            border: 1px solid #ccc;
            font-family: Arial, sans-serif;
            font-size: 12px;
        """)
        links_layout.addWidget(self.links_list)
        self.links_group.setLayout(links_layout)
        middle_layout.addWidget(self.links_group, 1)  # Expand

        splitter.addWidget(middle_widget)

        # Right side: AI Email Decoder (Chat with Gemini)
        right_widget = QWidget()
        right_layout = QVBoxLayout(right_widget)
        right_widget.setStyleSheet("background-color: #ffffff;")
        
        # Title
        title_label = QLabel("🚀 AI Email Decoder (Gemma)")
        title_label.setStyleSheet("""
            font-size: 18px;
            font-weight: bold;
            color: #000000;
            font-family: Arial, sans-serif;
        """)
        right_layout.addWidget(title_label)
        
        # Options group
        options_group = QGroupBox("Include in Prompt")
        options_group.setStyleSheet("""
            QGroupBox {
                color: #000000;
                font-weight: bold;
                background-color: #ffffff;
                border: 1px solid #ccc;
                margin-top: 5px;
            }
            QGroupBox::title {
                color: #000000;
                font-weight: bold;
            }
        """)
        options_layout = QVBoxLayout()
        self.header_check = QCheckBox("Message Header")
        self.header_check.setStyleSheet("color: #000000; background-color: #ffffff;")
        self.body_check = QCheckBox("Message Body")
        self.body_check.setStyleSheet("color: #000000; background-color: #ffffff;")
        self.links_check = QCheckBox("Links")
        self.links_check.setStyleSheet("color: #000000; background-color: #ffffff;")
        options_layout.addWidget(self.header_check)
        options_layout.addWidget(self.body_check)
        options_layout.addWidget(self.links_check)
        options_group.setLayout(options_layout)
        right_layout.addWidget(options_group)
        
        # Chat display
        self.chat_display = QTextBrowser()
        self.chat_display.setStyleSheet("""
            background-color: #f9f9f9;
            color: #000000;
            border: 1px solid #ccc;
            font-family: Arial, sans-serif;
            font-size: 12px;
        """)
        right_layout.addWidget(self.chat_display, 1)
        
        # Input area
        input_layout = QHBoxLayout()
        self.chat_input = QLineEdit()
        self.chat_input.setPlaceholderText("Ask Gemma about the email...")
        self.chat_input.setStyleSheet("""
            background-color: #ffffff;
            color: #000000;
            border: 1px solid #ccc;
            padding: 4px;
        """)
        input_layout.addWidget(self.chat_input)
        self.send_button = QPushButton("Send")
        self.send_button.setStyleSheet("""
            background-color: #e0e0e0;
            color: #000000;
            border: 1px solid #ccc;
            padding: 4px;
        """)
        self.send_button.clicked.connect(self.send_to_gemma)
        input_layout.addWidget(self.send_button)
        right_layout.addLayout(input_layout)
        
        splitter.addWidget(right_widget)
        splitter.setSizes([533, 534, 533])  # Exact equal thirds

    def load_email(self):
        """Load the email into the web view."""
        self.view.setHtml(self.html)

        # Wait a tick for JS/DOM to settle, then extract
        QTimer.singleShot(800, self.extract_dom)

    def extract_dom(self):
        """Extract DOM data and update right side."""
        js = """
        (() => {
          const links = Array.from(document.querySelectorAll('a'))
            .map(a => ({text: (a.innerText||'').trim(), href: a.href.length > 50 ? a.href.substring(0, 50) + '...' : a.href}))
            .filter(x => x.href && !x.href.startsWith('mailto:') && !x.href.startsWith('tel:'));
          const text = document.body ? document.body.innerText : '';
          return { text, links };
        })();
        """
        self.view.page().runJavaScript(js, self.on_dom)

    def on_dom(self, dom):
        """Handle extracted DOM data."""
        out = {
            "email_ptr": self.ptr,
            "dom": {
                "visible_text": dom.get("text", "") if dom else "",
                "links": dom.get("links", []) if dom else []
            }
        }

        # Save to JSON file
        with open('email_output.json', 'w') as f:
            json.dump(out, f, indent=2)

        self.email_data = out

        # Update UI
        self.display_email_details(self.ptr)
        self.display_body(out["dom"]["visible_text"])
        self.display_links(out["dom"]["links"])

    def display_email_details(self, ptr):
        """Display email pointer details."""
        details_layout = self.details_group.layout()

        # Clear previous
        for i in reversed(range(details_layout.count())):
            widget = details_layout.itemAt(i).widget()
            if widget:
                widget.setParent(None)

        # Add new
        for key, value in ptr.items():
            label = QLabel(f"{key.replace('_', ' ').title()}: {value}")
            label.setStyleSheet("font-weight: bold; color: #000000;")
            details_layout.addWidget(label)

    def display_body(self, text):
        """Display the email body text."""
        self.body_browser.setPlainText(text)

    def display_links(self, links):
        """Display the list of links."""
        self.links_list.clear()
        for link in links:
            text = link.get('text', '')
            href = link.get('href', '')
            item_text = f"{text} -> {href}"
            item = QListWidgetItem(item_text)
            item.setToolTip(href)  # Show full href on hover
            self.links_list.addItem(item)

    def send_to_gemma(self):
        """Send selected data and user message to Gemma."""
        if not self.email_data:
            self.chat_display.append("No email data available.")
            return
        
        user_message = self.chat_input.text().strip()
        if not user_message:
            return
        
        # Collect selected data
        context = ""
        if self.header_check.isChecked():
            context += f"Header: {json.dumps(self.email_data['email_ptr'], indent=2)}\n\n"
        if self.body_check.isChecked():
            context += f"Body: {self.email_data['dom']['visible_text']}\n\n"
        if self.links_check.isChecked():
            context += f"Links: {json.dumps(self.email_data['dom']['links'], indent=2)}\n\n"
        
        full_message = f"{context}User: {user_message}"
        
        # Add to chat
        self.chat_display.append(f"You: {user_message}")
        self.chat_input.clear()
        self.chat_display.append("Gemma: Thinking...")
        
        # Send to Ollama
        try:
            client = ollama.Client(host=config["ollama_host"])
            response = client.chat(
                model=config["model"],
                messages=[{"role": "user", "content": full_message}],
            )
            gemma_response = response["message"]["content"].strip()
            self.chat_display.append(f"Gemma: {gemma_response}")
        except Exception as e:
            self.chat_display.append(f"Gemma: Error - {e}")
        
        # Scroll to bottom
        scrollbar = self.chat_display.verticalScrollBar()
        scrollbar.setValue(scrollbar.maximum())

if __name__ == "__main__":
    ptr, html = get_latest_planhub_email_html()
    app = QApplication(sys.argv)
    viewer = CombinedViewer(html, ptr)
    viewer.show()
    sys.exit(app.exec())