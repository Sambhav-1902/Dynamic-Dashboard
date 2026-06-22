"""
Outlook (Classic Desktop App) Mail Listener — COM Automation Version
-----------------------------------------------------------------------
This is the OFFICE LAPTOP version of the mail tracker. It automates the
already-running, already-signed-in Outlook desktop application via COM
(pywin32) — NO Microsoft Graph API, NO admin consent, NO app registration
required. It only works on Windows with Outlook (Classic) installed.

All AI logic, Excel writing, and dashboard generation are IDENTICAL to
the Gmail version (track_outlook_mails.py) — only the email source
changed.

Each ticket is automatically assigned:
- Status  (Pending, Ongoing, Resolved)
- Category (multi-label, detected by AI)
- Due Date (extracted by AI)
- Followups counter
- Ongoing Since (timestamp of when the ticket first became Ongoing,
  used by the dashboard's "days ongoing" list)

Status logic:
- New ticket arrives                          -> Pending, Followups = 0, run category detection (once)
- Reply from original sender + Pending        -> increment Followups, keep Pending
- Reply from original sender + Ongoing        -> keep Ongoing
- Reply from original sender + Resolved       -> NEW row (new issue), run category detection (once)
- Reply from anyone else                      -> conversational resolution check (a small local
                                                  instruction model reads the FULL conversation and
                                                  decides if every question/issue from the original
                                                  sender has been answered)
                                                  -> Ongoing or Resolved. Defaults to Ongoing if
                                                  the resolution model errors out.

Category is detected once when a ticket is first created and does not
change on replies — short reply text does not carry reliable category info.
Category detection uses margin-based matching: the top-scoring category is
always included, plus any category within 0.05 of the top score (provided
the top score is at least 0.5). This can occasionally over-include or
mis-categorize on certain short/ambiguous phrasings — see detect_categories()
docstring for known edge cases.

Dashboard: a "Dashboard" sheet is regenerated as a static snapshot each
time historical load completes (and once more when the script is
stopped) — overview counts, a status pie chart, grouped Pending/Ongoing
bar charts by category and by assignee, and a list of Ongoing tickets
sorted by days-in-Ongoing. See update_dashboard() for details.

Requirements:
    pip install pywin32 openpyxl sentence-transformers torch gliner parsedatetime requests

Also requires Ollama running locally with gemma3:1b imported from a local GGUF file.
The model was imported using a Modelfile pointing at the locally cached GGUF:
    1. Create a Modelfile containing: FROM <path to google_gemma-3-1b-it-Q4_K_M.gguf>
    2. Run: ollama create gemma3:1b -f Modelfile
    3. Verify: ollama run gemma3:1b "Reply with only YES"
Ollama must be running when this script starts (check the system tray or run
`ollama serve`). If unreachable, resolution checks default to Ongoing.

Run:
    python track_outlook_mails_com.py
"""

import win32com.client
import pythoncom
import email
import requests
from email.header import decode_header
from email.utils import parseaddr, parsedate_to_datetime
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.chart import PieChart, BarChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.drawing.text import CharacterProperties, ParagraphProperties
import os
import re
import time
import parsedatetime
from datetime import datetime, timedelta
from collections import Counter, defaultdict
from gliner import GLiNER
from sentence_transformers import util as st_util

OUTPUT_FILE          = "mail_tracker.xlsx"
ALLOWED_SENDERS_FILE = "allowed_senders.txt"
POLL_INTERVAL        = 15  # seconds between each inbox check

# MAPI property tag for the full raw transport headers (Message-ID,
# In-Reply-To, Date) — same info the Gmail/IMAP version gets natively.
PR_TRANSPORT_MESSAGE_HEADERS = "http://schemas.microsoft.com/mapi/proptag/0x007D001E"
# MAPI property tag for resolving "EX" (Exchange) sender addresses to SMTP
PR_SMTP_ADDRESS = "http://schemas.microsoft.com/mapi/proptag/0x39FE001E"

# olMail = 43 (MailItem.Class) — used to skip meeting requests, receipts, etc.
OL_MAIL_ITEM    = 43
OL_FOLDER_INBOX = 6

# AI models
print("Loading AI models...")
deadline_model = GLiNER.from_pretrained("urchade/gliner_small-v2.1")
# Uses sentence embeddings + semantic similarity — fast (0.04s) and accurate
from sentence_transformers import SentenceTransformer
category_model     = SentenceTransformer('BAAI/bge-base-en-v1.5')
print("AI models ready.\n")

# Only emails from these senders will be saved.
ALLOWED_EMAILS = [
    "sambhavsingwi@gmail.com",
    "cs1230722@iitd.ac.in",
]

# Allow everyone from these domains
ALLOWED_DOMAINS = [
    # "exlservice.com",
]


# ---------------------------------------------------------------------------
# Dynamic allowed senders management
# ---------------------------------------------------------------------------

def load_allowed_senders():
    senders = set(e.lower() for e in ALLOWED_EMAILS)
    if os.path.exists(ALLOWED_SENDERS_FILE):
        with open(ALLOWED_SENDERS_FILE, "r", encoding="utf-8") as f:
            for line in f:
                addr = line.strip().lower()
                if addr:
                    senders.add(addr)
    return senders


def save_allowed_senders(senders):
    with open(ALLOWED_SENDERS_FILE, "w", encoding="utf-8") as f:
        for addr in sorted(senders):
            f.write(addr + "\n")


def extract_and_update_senders(msg, current_allowed):
    new_addresses = []
    for header in ["To", "Cc", "Bcc"]:
        raw = msg.get(header, "")
        if not raw:
            continue
        for part in raw.split(","):
            _, addr = parseaddr(part.strip())
            addr = addr.lower().strip()
            if addr and "@" in addr and addr not in current_allowed:
                new_addresses.append(addr)
                current_allowed.add(addr)
    if new_addresses:
        save_allowed_senders(current_allowed)
        print(f"Added {len(new_addresses)} new sender(s) to allowed list:")
        for addr in new_addresses:
            print(f"   + {addr}")
    return current_allowed


# ---------------------------------------------------------------------------
# Sender filter
# ---------------------------------------------------------------------------

def is_allowed(sender_email, allowed_senders):
    sender_email = sender_email.lower().strip()
    if sender_email in allowed_senders:
        return True
    return False

    # Inactive: filter by domain
    # sender_domain = sender_email.split("@")[-1]
    # if sender_domain in [d.lower() for d in ALLOWED_DOMAINS]:
    #     return True
    # return False


# ---------------------------------------------------------------------------
# Email parsing helpers (COM equivalents of the Gmail/IMAP version)
# ---------------------------------------------------------------------------

def decode_str(value):
    if not value:
        return ""
    parts = decode_header(value)
    decoded = []
    for part, charset in parts:
        if isinstance(part, bytes):
            decoded.append(part.decode(charset or "utf-8", errors="replace"))
        else:
            decoded.append(part)
    return " ".join(decoded).strip()


def strip_html(text):
    text = re.sub(r'<[^>]+>', ' ', text)
    text = re.sub(r'&nbsp;', ' ', text)
    text = re.sub(r'&amp;', '&', text)
    text = re.sub(r'&lt;', '<', text)
    text = re.sub(r'&gt;', '>', text)
    text = re.sub(r'&quot;', '"', text)
    text = re.sub(r'\s+', ' ', text)
    return text.strip()


def get_sender_smtp_address(item):
    """
    Returns a clean SMTP email address for the sender.
    For internal Exchange senders, SenderEmailAddress returns an EX
    format string instead of a normal SMTP email address.
    PR_SMTP_ADDRESS resolves this to a clean address.
    Confirmed working for both EX and SMTP sender types.
    """
    try:
        sender = item.Sender  # AddressEntry
        if sender is not None:
            smtp = sender.PropertyAccessor.GetProperty(PR_SMTP_ADDRESS)
            if smtp:
                return smtp
    except Exception:
        pass
    try:
        if getattr(item, "SenderEmailType", "") == "SMTP":
            return item.SenderEmailAddress
    except Exception:
        pass
    return getattr(item, "SenderEmailAddress", "") or ""


def parse_email_com(item):
    """
    Builds the same dict structure as parse_email() in the Gmail version,
    but reads from an Outlook COM MailItem instead of raw IMAP bytes.

    Gets the full raw transport headers via PropertyAccessor and parses
    them with Python's email module — confirmed this gives real
    Message-ID, In-Reply-To, and Date headers (with In-Reply-To correctly
    pointing to the parent's Message-ID for replies), so
    extract_and_update_senders() and find_matching_row() work unchanged.
    """
    try:
        raw_headers = item.PropertyAccessor.GetProperty(PR_TRANSPORT_MESSAGE_HEADERS)
    except Exception:
        raw_headers = ""

    msg = email.message_from_string(raw_headers) if raw_headers else email.message.Message()

    sender_email = get_sender_smtp_address(item)
    sender_name  = getattr(item, "SenderName", "") or sender_email

    subject = getattr(item, "Subject", "") or "(No Subject)"

    # Date — prefer the parsed header (matches IMAP version's timezone
    # normalization, confirmed needed since headers can carry non-local
    # offsets like -0700 PDT); fall back to Outlook's ReceivedTime if the
    # header is missing or unparseable.
    date_str = None
    header_date = msg.get("Date")
    if header_date:
        try:
            parsed_dt = parsedate_to_datetime(header_date)
            if parsed_dt.tzinfo is not None:
                parsed_dt = parsed_dt.astimezone()
            date_str = parsed_dt.strftime("%Y-%m-%d %H:%M:%S")
        except Exception:
            date_str = None
    if date_str is None:
        try:
            rt = item.ReceivedTime  # pywintypes datetime, already local
            date_str = datetime(rt.year, rt.month, rt.day, rt.hour, rt.minute, rt.second).strftime("%Y-%m-%d %H:%M:%S")
        except Exception:
            date_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    # Body — prefer plain text Body; fall back to stripped HTMLBody
    body = ""
    try:
        body = item.Body or ""
    except Exception:
        body = ""
    if not body.strip():
        try:
            body = strip_html(item.HTMLBody or "")
        except Exception:
            body = ""
    body = body.strip()[:2000]

    message_id  = msg.get("Message-ID", "").strip()
    in_reply_to = msg.get("In-Reply-To", "").strip()

    # If headers were unavailable (e.g. item not yet fully synced), fall
    # back to EntryID as a stable identifier so thread matching still has
    # *something* to compare on subject at least.
    if not message_id:
        try:
            message_id = f"<entryid-{item.EntryID}@local>"
        except Exception:
            message_id = ""

    return {
        "sender_name":  sender_name or sender_email,
        "sender_email": sender_email,
        "subject":      decode_str(subject),
        "received":     date_str,
        "body":         body,
        "message_id":   message_id,
        "in_reply_to":  in_reply_to,
        "raw_msg":      msg,
        "_entry_id":    getattr(item, "EntryID", None),  # used for dedup tracking
    }


def clean_subject(subject):
    cleaned = subject.strip()
    prefix_pattern = re.compile(r'^(re|fwd|fw)\s*:\s*', re.IGNORECASE)
    while prefix_pattern.match(cleaned):
        cleaned = prefix_pattern.sub('', cleaned).strip()
    return cleaned.lower()


# ---------------------------------------------------------------------------
# AI — Status detection
# ---------------------------------------------------------------------------

def strip_quoted_content(body):
    """
    Removes quoted/forwarded content from an email body, returning only
    the new text the sender actually wrote. Handles common separator
    patterns used by Outlook, Gmail, and other email clients.
    """
    text = body.strip()

    # Separators that mark the start of quoted/forwarded content
    separators = [
        "________________________________",
        "-----Original Message-----",
        "-----Forwarded Message-----",
        "-------- Original Message --------",
        "-------------------------",
    ]

    for sep in separators:
        idx = text.find(sep)
        if idx > 5:
            text = text[:idx].strip()
            break

    # Handle "On <date>, <name> wrote:" pattern (Gmail style)
    on_wrote_pattern = re.compile(
        r'\n?On\s+(Mon|Tue|Wed|Thu|Fri|Sat|Sun)[a-z]*,.*?wrote:',
        re.IGNORECASE | re.DOTALL
    )
    match = on_wrote_pattern.search(text)
    if match and match.start() > 5:
        text = text[:match.start()].strip()

    # Remove lines that start with '>' (quoted reply lines)
    lines = text.split("\n")
    clean_lines = [line for line in lines if not line.strip().startswith(">")]
    text = "\n".join(clean_lines).strip()

    # Remove common email disclaimer/caution footers
    disclaimer_patterns = [
        r'CAUTION:.*?safe\.',
    ]
    for pattern in disclaimer_patterns:
        text = re.sub(pattern, '', text, flags=re.IGNORECASE | re.DOTALL).strip()

    return text


def get_first_300_words(body):
    """
    Returns the first 300 words of the latest reply only, with quoted
    thread content stripped out.
    """
    text  = strip_quoted_content(body)
    words = text.split()
    return " ".join(words[:300])


# ---------------------------------------------------------------------------
# AI — Conversational resolution check (Gemma 3 1B via local Ollama)
# ---------------------------------------------------------------------------
#
# Uses the same Ollama setup as the personal laptop version. The model
# could not be downloaded via `ollama pull` (corporate network blocks
# Ollama's registry), but was imported from a local GGUF file instead:
#
#   1. The GGUF was downloaded from Hugging Face (bartowski/google_gemma-3-1b-it-GGUF)
#   2. A Modelfile was created pointing at the local file path
#   3. `ollama create gemma3:1b -f Modelfile` registered it with Ollama
#
# Ollama runs as a background service on Windows after installation.
# No extra Python packages needed — just the `requests` library (already
# used elsewhere in this script).

OLLAMA_URL   = "http://localhost:11434/api/generate"
OLLAMA_MODEL = "gemma3:1b"


def is_conversation_resolved(conversation_log, original_sender):
    """
    Asks gemma3:1b (via local Ollama) a single yes/no question: has EVERY
    question/issue raised by original_sender been answered across the
    WHOLE conversation (not just the latest reply)?

    Identical to the personal laptop version — same model, same prompt,
    same decision logic, same ~6-7s per call speed.

    Returns True (Resolved) or False (Ongoing). Defaults to False if
    Ollama is unreachable or gives an unclear response.
    """
    prompt = f"""Below is a support ticket conversation, shown NEWEST message first. Each message shows the sender's email and their message.

Conversation:
{conversation_log}

The person who raised this ticket is: {original_sender}

Question: Has EVERY question or issue raised by {original_sender} (across all of their messages) been clearly and explicitly answered or fixed by the OTHER person's replies in this conversation?

If even ONE question or issue from {original_sender} is still unanswered or unresolved, answer NO.
Only answer YES if you can find an explicit answer/fix for EACH thing {original_sender} asked about.

Answer with ONLY one word: YES or NO."""

    try:
        response = requests.post(OLLAMA_URL, json={
            "model": OLLAMA_MODEL,
            "prompt": prompt,
            "stream": False,
            "options": {"temperature": 0.1}
        }, timeout=30)
        text = response.json()["response"].strip().upper()

        if "YES" in text and "NO" not in text:
            return True
        elif "NO" in text:
            return False
        else:
            print(f"   Unclear response from Gemma ({text!r}) — defaulting to Ongoing")
            return False

    except Exception as e:
        print(f"   Could not reach Ollama ({e}) — defaulting to Ongoing.")
        return False


# ---------------------------------------------------------------------------
# AI — Category detection
# ---------------------------------------------------------------------------

# Category labels — descriptive sentences work better for semantic similarity
CATEGORY_LABELS = [
    'unable to login or access account, password not working, account locked',
    'payment transaction failed, duplicate charge, refund error, amount wrongly deducted, EMI payment not going through',
    'loan application status, loan approval delay, loan disbursement not credited, loan tenure or interest rate change, loan account closure',
    'wrong data in report, data mismatch, incorrect figures, reconciliation error',
    'application crash, software bug, system not working, page not loading',
    'API error, integration failure, sync timeout, third party system not responding',
    'system very slow, page timeout, high latency, unresponsive system',
    'KYC verification failed, eSign not working, document upload compliance issue',
    'how to do something, asking for information, no technical problem just a question',
]

# Full display names mapped from labels
LABEL_TO_FULL = {
    'unable to login or access account, password not working, account locked'              : 'Access and Login Issues',
    'payment transaction failed, duplicate charge, refund error, amount wrongly deducted, EMI payment not going through' : 'Payment Processing Issues',
    'loan application status, loan approval delay, loan disbursement not credited, loan tenure or interest rate change, loan account closure' : 'Loan Processing Issues',
    'wrong data in report, data mismatch, incorrect figures, reconciliation error'        : 'Data and Reporting Issues',
    'application crash, software bug, system not working, page not loading'               : 'System and Application Errors',
    'API error, integration failure, sync timeout, third party system not responding'     : 'Integration and API Issues',
    'system very slow, page timeout, high latency, unresponsive system'                   : 'Performance Issues',
    'KYC verification failed, eSign not working, document upload compliance issue'        : 'Compliance and KYC Issues',
    'how to do something, asking for information, no technical problem just a question'   : 'General Query',
}

CATEGORY_MARGIN        = 0.05  # categories within this margin of the top score are also included
CATEGORY_MIN_TOP_SCORE = 0.5   # if even the top score is below this, text is too vague to categorize confidently

# Pre-encode category label embeddings once at startup
CATEGORY_EMBS = category_model.encode(CATEGORY_LABELS, convert_to_tensor=True)

# Question words that suggest a general query
QUESTION_WORDS = ["how to", "how do", "what is", "what are", "what's",
                  "can you tell", "can you please tell", "could you tell",
                  "please tell", "please let me know", "please explain",
                  "i want to know", "i wanted to know", "could you",
                  "would you", "is there", "are there", "do you know",
                  "when is", "when will", "when can"]

# Problem words that confirm a technical issue
PROBLEM_WORDS  = ["failed", "error", "not working", "stuck", "issue", "problem",
                  "crash", "unable", "cannot", "can't", "broken", "wrong",
                  "incorrect", "missing", "slow", "timeout", "deducted", "not loading"]

def detect_categories(body):
    """
    Detects categories using sentence embedding margin-based matching
    (Approach B from final comparison testing — 88% on 25 realistic
    short/long/tricky test cases, with focused 1-2 category output
    in the vast majority of cases).

    1. Pre-check for general queries (question words, no problem words)
    2. Truncate to first 150 words
    3. Encode and compare against category label embeddings
    4. Take the top score; include any category within MARGIN of it
       (only if top score >= MIN_TOP_SCORE, otherwise just take top 1)
    Returns comma-separated full category name string.

    Known limitations (documented edge cases from testing):
    - Pure informational questions containing "loan" (e.g. "what documents
      are needed for a personal loan?") may be detected as Loan Processing
      Issues instead of General Query.
    - "Login to make a payment" style emails may be detected as Payment
      Processing Issues instead of Access and Login Issues.
    - App names containing "pay" (e.g. "Paymentor") may bias toward
      Payment Processing Issues even when the issue is a system/app error.
    """
    words = body.strip().split()
    text  = " ".join(words[:150])
    text_lower = text.lower()

    # Pre-check: if question words present and no problem words -> General Query
    has_question = any(q in text_lower for q in QUESTION_WORDS)
    has_problem  = any(p in text_lower for p in PROBLEM_WORDS)
    if has_question and not has_problem:
        return "General Query"

    try:
        emb    = category_model.encode(text, convert_to_tensor=True)
        scores = st_util.cos_sim(emb, CATEGORY_EMBS)[0].tolist()

        score_map = {LABEL_TO_FULL[label]: score for label, score in zip(CATEGORY_LABELS, scores)}
        top_score = max(score_map.values())

        if top_score < CATEGORY_MIN_TOP_SCORE:
            detected = [max(score_map.items(), key=lambda x: x[1])[0]]
        else:
            cutoff   = top_score - CATEGORY_MARGIN
            detected = [cat for cat, score in score_map.items() if score >= cutoff]

    except Exception as e:
        print(f"Category detection error: {e} — defaulting to General Query")
        detected = ["General Query"]

    return ", ".join(detected)


# ---------------------------------------------------------------------------
# AI — Due date extraction
# ---------------------------------------------------------------------------

_cal = parsedatetime.Calendar()

_SPECIAL_CASES = {
    "weekend"    : lambda: (datetime.now() + timedelta(days=max(1, (5 - datetime.now().weekday()) % 7))).replace(hour=17, minute=0, second=0, microsecond=0),
    "eod"        : lambda: datetime.now().replace(hour=17, minute=0, second=0, microsecond=0),
    "end of day" : lambda: datetime.now().replace(hour=17, minute=0, second=0, microsecond=0),
    "eow"        : lambda: (datetime.now() + timedelta(days=max(1, (4 - datetime.now().weekday()) % 7))).replace(hour=17, minute=0, second=0, microsecond=0),
    "end of week": lambda: (datetime.now() + timedelta(days=max(1, (4 - datetime.now().weekday()) % 7))).replace(hour=17, minute=0, second=0, microsecond=0),
}

DEADLINE_KEYWORDS = [
    "by", "before", "due", "deadline", "complete", "finish",
    "done", "fix", "resolve", "submit", "deliver", "send",
    "no later than", "needs to be", "should be", "must be",
    "have to", "need to", "expected by", "required by",
    "end of day", "end of week", "eod", "eow",
]


def find_deadline_sentence(text, entity_text):
    sentences = re.split(r'[.!?\n]', text)
    for sentence in sentences:
        if entity_text.lower() in sentence.lower():
            if any(kw in sentence.lower() for kw in DEADLINE_KEYWORDS):
                return sentence
    return None


def extract_deadline(body):
    text       = get_first_300_words(body)
    text_lower = text.lower()

    if "asap" in text_lower:
        return "ASAP"
    if "urgent" in text_lower:
        return "URGENT"

    try:
        entities = deadline_model.predict_entities(text, ["date", "time", "duration"], threshold=0.3)
    except Exception as e:
        print(f"Due date extraction error: {e}")
        return ""

    if not entities:
        return ""

    for entity in entities:
        date_text = entity["text"].strip().lower()
        if not find_deadline_sentence(text, date_text):
            continue
        for key, fn in _SPECIAL_CASES.items():
            if key in date_text:
                return fn().strftime("%Y-%m-%d %H:%M")
        result, status = _cal.parseDT(date_text, datetime.now())
        if status != 0:
            return result.strftime("%Y-%m-%d %H:%M")

    return ""



# ---------------------------------------------------------------------------
# Assignee name extraction
# ---------------------------------------------------------------------------

def extract_assignee_name(sender_name, sender_email, body):
    """
    Extracts the assignee name using three methods in order of preference:
    1. Sender name from email header (most reliable)
    2. Clean up email address prefix
    3. GLiNER scans the first 300 words of body for person names
    Returns a clean name string or empty string if nothing found.
    """

    # Method 1 — Sender name from header
    if sender_name and sender_name.strip():
        name = sender_name.strip()
        # Make sure it looks like a real name not a generic sender
        # Skip names that look like system senders e.g. "No Reply", "Support Team"
        skip_keywords = ["noreply", "no-reply", "support", "team", "system",
                         "notification", "alert", "admin", "info", "do-not-reply"]
        if not any(kw in name.lower() for kw in skip_keywords):
            return name

    # Method 2 — Clean up email address
    if sender_email and "@" in sender_email:
        local_part = sender_email.split("@")[0]
        # Remove common prefixes like ex_, usr_, emp_
        local_part = re.sub(r'^(ex_|usr_|emp_|user_|staff_)', '', local_part, flags=re.IGNORECASE)
        # Replace dots, underscores, hyphens with spaces
        local_part = re.sub(r'[._-]', ' ', local_part)
        # Remove numbers
        local_part = re.sub(r'\d+', '', local_part).strip()
        if local_part:
            # Capitalize each word
            return local_part.title()

    # Method 3 — GLiNER scans body for person names
    if body and body.strip():
        text = get_first_300_words(body)
        try:
            entities = deadline_model.predict_entities(text, ["person name"], threshold=0.4)
            if entities:
                # Take the first detected person name
                return entities[0]["text"].strip().title()
        except Exception as e:
            print(f"Name extraction error: {e}")

    return ""

# ---------------------------------------------------------------------------
# Excel setup and operations
# ---------------------------------------------------------------------------

HEADERS = [
    "#",
    "Sender Name",
    "Sender Email",
    "Subject",
    "Category",
    "Received Date & Time",
    "Latest Reply Date & Time",
    "Status",
    "Ongoing Since",
    "Assigned To",
    "Followups",
    "Due Date",
    "Body",
    "Message-ID",
]

COL_NUM          = 1
COL_SENDER_NAME  = 2
COL_SENDER_EMAIL = 3
COL_SUBJECT      = 4
COL_CATEGORY     = 5
COL_RECEIVED     = 6
COL_LATEST_DATE  = 7
COL_STATUS       = 8
COL_ONGOING_SINCE = 9
COL_ASSIGNED_TO  = 10
COL_FOLLOWUPS    = 11
COL_DUE_DATE     = 12
COL_BODY         = 13
COL_MESSAGE_ID   = 14

STATUS_COLORS = {
    "Pending":  "FFF2CC",
    "Ongoing":  "DDEEFF",
    "Resolved": "E2EFDA",
}


def init_excel(filepath):
    if os.path.exists(filepath):
        print(f"Found existing file — new emails will be appended to '{filepath}'")
        return

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Inbox Tracker"

    header_fill = PatternFill("solid", start_color="1F4E79")
    header_font = Font(bold=True, color="FFFFFF", name="Arial")

    for col, header in enumerate(HEADERS, start=1):
        cell = ws.cell(row=1, column=col, value=header)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    ws.column_dimensions["A"].width = 5
    ws.column_dimensions["B"].width = 22
    ws.column_dimensions["C"].width = 32
    ws.column_dimensions["D"].width = 40
    ws.column_dimensions["E"].width = 35   # Category
    ws.column_dimensions["F"].width = 22
    ws.column_dimensions["G"].width = 22
    ws.column_dimensions["H"].width = 12   # Status
    ws.column_dimensions["I"].width = 20   # Ongoing Since
    ws.column_dimensions["J"].width = 30   # Assigned To
    ws.column_dimensions["K"].width = 10   # Followups
    ws.column_dimensions["L"].width = 20   # Due Date
    ws.column_dimensions["M"].width = 60   # Body
    ws.column_dimensions["N"].width = 40   # Message-ID
    ws.row_dimensions[1].height = 30

    wb.save(filepath)
    print(f"Created new tracker file: '{filepath}'")


def apply_status_color(cell, status):
    color = STATUS_COLORS.get(status, "FFFFFF")
    cell.fill      = PatternFill("solid", start_color=color)
    cell.font      = Font(name="Arial", size=10, bold=True)
    cell.alignment = Alignment(horizontal="center", vertical="top")


def get_all_message_ids(ws):
    id_to_row = {}
    for row in range(2, ws.max_row + 1):
        mid = ws.cell(row=row, column=COL_MESSAGE_ID).value
        if mid:
            id_to_row[mid.strip()] = row
    return id_to_row


def get_all_subjects(ws):
    subject_to_row = {}
    for row in range(2, ws.max_row + 1):
        subject = ws.cell(row=row, column=COL_SUBJECT).value
        if subject:
            subject_to_row[clean_subject(subject)] = row
    return subject_to_row


def find_matching_row(em, ws):
    id_to_row  = get_all_message_ids(ws)
    target_row = id_to_row.get(em["in_reply_to"])
    if target_row:
        return target_row, "message-id"
    subject_to_row   = get_all_subjects(ws)
    cleaned_incoming = clean_subject(em["subject"])
    target_row       = subject_to_row.get(cleaned_incoming)
    if target_row:
        return target_row, "subject"
    return None, None


def append_new_email(em, filepath):
    """Adds a new row for a fresh ticket with Status=Pending, Followups=0, and runs AI detection."""
    try:
        wb       = openpyxl.load_workbook(filepath)
        ws       = wb["Inbox Tracker"]
        next_row = ws.max_row + 1
        index    = next_row - 1

        alt_fill = PatternFill("solid", start_color="DCE6F1")
        row_fill = alt_fill if index % 2 == 0 else PatternFill()

        print(f"   Running category detection...")
        category = detect_categories(em["body"])
        due_date = extract_deadline(em["body"])

        # Use the same timestamped format as replies, so the conversation
        # log is consistent from the very first entry
        clean_body   = strip_quoted_content(em["body"])
        initial_body = f"[{em['received']}] {em['sender_email']}:\n{clean_body}"

        values = [
            index,
            em["sender_name"],
            em["sender_email"],
            em["subject"],
            category,
            em["received"],
            "",           # Latest Reply Date
            "Pending",
            "",           # Ongoing Since — blank until status first becomes Ongoing
            "",           # Assigned To — empty until someone other than sender replies
            0,            # Followups
            due_date,
            initial_body,
            em["message_id"],
        ]

        for col, value in enumerate(values, start=1):
            cell = ws.cell(row=next_row, column=col, value=value)
            cell.font      = Font(name="Arial", size=10)
            cell.fill      = row_fill
            cell.alignment = Alignment(vertical="top", wrap_text=(col == COL_BODY))

        apply_status_color(ws.cell(row=next_row, column=COL_STATUS), "Pending")
        ws.cell(row=next_row, column=COL_FOLLOWUPS).alignment = Alignment(horizontal="center", vertical="top")

        wb.save(filepath)
        return True
    except PermissionError:
        return False


def update_existing_row(em, target_row, filepath):
    """
    Updates an existing ticket row based on who replied and current status.

    Logic:
    - Reply from original sender + Pending  -> increment Followups, keep Pending
    - Reply from original sender + Ongoing  -> keep Ongoing
    - Reply from original sender + Resolved -> create NEW row (new issue)
    - Reply from anyone else                -> AI status check

    Category is NOT re-evaluated on replies — it is set once when the
    ticket is first created and remains unchanged. Short reply text
    (acknowledgements, one-liners) does not carry reliable category
    information and was causing mis-categorization.
    """
    try:
        wb = openpyxl.load_workbook(filepath)
        ws = wb["Inbox Tracker"]

        original_sender = (ws.cell(row=target_row, column=COL_SENDER_EMAIL).value or "").lower().strip()
        reply_sender    = em["sender_email"].lower().strip()
        current_status  = ws.cell(row=target_row, column=COL_STATUS).value or "Pending"
        is_same_sender  = (reply_sender == original_sender)

        if is_same_sender and current_status == "Resolved":
            # New issue raised after resolution — handled by process_email, not here
            return "new_issue"

        # Category is set once at ticket creation and not re-evaluated here
        due_date = extract_deadline(em["body"])

        # Strip quoted/forwarded content from the new reply so we only
        # store the new text the sender actually wrote, then build the
        # updated conversation log (newest entry on top). This is needed
        # up front since the Gemma resolution check below requires the
        # FULL conversation, not just the latest message.
        existing_body = ws.cell(row=target_row, column=COL_BODY).value or ""
        clean_reply   = strip_quoted_content(em["body"])
        new_entry     = f"[{em['received']}] {em['sender_email']}:\n{clean_reply}"
        combined_body = f"{new_entry}\n\n---\n\n{existing_body}" if existing_body else new_entry

        if is_same_sender and current_status == "Pending":
            # Sender following up — increment Followups, keep Pending
            current_followups = ws.cell(row=target_row, column=COL_FOLLOWUPS).value or 0
            new_followups     = int(current_followups) + 1
            ws.cell(row=target_row, column=COL_FOLLOWUPS).value     = new_followups
            ws.cell(row=target_row, column=COL_FOLLOWUPS).alignment = Alignment(horizontal="center", vertical="top")
            status = "Pending"
            print(f"   Ticket raiser following up — status stays Pending, followups now {new_followups}")

        elif is_same_sender and current_status == "Ongoing":
            # Sender adding info while TechOps is working — keep Ongoing
            status = "Ongoing"
            print(f"   Sender adding info — status stays Ongoing")

        else:
            # Someone else replied — ask Gemma whether ALL questions/issues
            # raised by the original sender have been answered across the
            # FULL conversation (handles multi-turn Q&A tickets where each
            # question is answered in a separate reply).
            print(f"   Reply from someone else — running conversational resolution check...")
            resolved = is_conversation_resolved(combined_body, original_sender)
            status   = "Resolved" if resolved else "Ongoing"
        combined_body = combined_body[:10000]  # Cap total length to keep file manageable

        ws.cell(row=target_row, column=COL_BODY).value          = combined_body
        ws.cell(row=target_row, column=COL_BODY).alignment      = Alignment(vertical="top", wrap_text=True)
        ws.cell(row=target_row, column=COL_LATEST_DATE).value   = em["received"]
        ws.cell(row=target_row, column=COL_LATEST_DATE).alignment = Alignment(vertical="top")
        status_cell = ws.cell(row=target_row, column=COL_STATUS, value=status)
        apply_status_color(status_cell, status)

        # Record when the ticket first became Ongoing — used for the
        # "days in Ongoing" dashboard list. Only set the very first time
        # status transitions into Ongoing; later Ongoing->Ongoing updates
        # (e.g. sender adding more info) must not reset this timestamp.
        if status == "Ongoing" and current_status != "Ongoing":
            existing_ongoing_since = ws.cell(row=target_row, column=COL_ONGOING_SINCE).value
            if not existing_ongoing_since:
                ws.cell(row=target_row, column=COL_ONGOING_SINCE).value     = em["received"]
                ws.cell(row=target_row, column=COL_ONGOING_SINCE).alignment = Alignment(vertical="top")

        if due_date:
            ws.cell(row=target_row, column=COL_DUE_DATE).value     = due_date
            ws.cell(row=target_row, column=COL_DUE_DATE).alignment = Alignment(vertical="top")

        # Update Assigned To — only when reply is NOT from original sender
        if not is_same_sender:
            assignee = extract_assignee_name(em["sender_name"], em["sender_email"], em["body"])
            if assignee:
                current_assigned = ws.cell(row=target_row, column=COL_ASSIGNED_TO).value or ""
                existing = [a.strip() for a in current_assigned.split(",") if a.strip()]
                if assignee not in existing:
                    existing.append(assignee)
                    ws.cell(row=target_row, column=COL_ASSIGNED_TO).value     = ", ".join(existing)
                    ws.cell(row=target_row, column=COL_ASSIGNED_TO).alignment = Alignment(vertical="top", wrap_text=True)
                    print(f"   Assigned to: {assignee}")

        wb.save(filepath)
        return True
    except PermissionError:
        return False


# ---------------------------------------------------------------------------
# Deciding what to do with each incoming email
# ---------------------------------------------------------------------------

def process_email(em, retry_queue, filepath):
    if em["in_reply_to"]:
        wb         = openpyxl.load_workbook(filepath)
        ws         = wb["Inbox Tracker"]
        target_row, match_type = find_matching_row(em, ws)

        if target_row:
            result = update_existing_row(em, target_row, filepath)

            if result == "new_issue":
                # Original sender raised a new issue after resolution — new row
                print(f"New issue raised after resolution — creating new ticket.")
                saved = append_new_email(em, filepath)
                if saved:
                    print(f"New ticket saved")
                else:
                    retry_queue.append(em)
                    print(f"Excel is open — queued for retry.")
                print(f"   From    : {em['sender_name']} <{em['sender_email']}>")
                print(f"   Subject : {em['subject']}")
                print(f"   Time    : {em['received']}\n")

            elif result is True:
                print(f"Reply received — ticket updated (matched by {match_type})")
                print(f"   Subject    : {em['subject']}")
                print(f"   Replied at : {em['received']}\n")

            else:
                retry_queue.append(em)
                print(f"Excel is open — reply queued: '{em['subject']}'\n")

        else:
            print(f"Got a reply but couldn't find the original ticket, saving as new.")
            saved = append_new_email(em, filepath)
            if not saved:
                retry_queue.append(em)
                print(f"   Excel is open — will retry when it's closed.\n")

    else:
        saved = append_new_email(em, filepath)
        if saved:
            print(f"New ticket saved")
        else:
            retry_queue.append(em)
            print(f"New ticket received but Excel is open — added to queue.")
        print(f"   From    : {em['sender_name']} <{em['sender_email']}>")
        print(f"   Subject : {em['subject']}")
        print(f"   Time    : {em['received']}\n")



# ---------------------------------------------------------------------------
# Outlook connection helpers (replaces IMAP connection helpers)
# ---------------------------------------------------------------------------

def connect_outlook():
    """
    Connects to the running (or starts a new) Outlook session via COM and
    returns the Inbox folder. No credentials needed — uses whichever
    account is already signed into the Outlook desktop app.
    """
    outlook   = win32com.client.Dispatch("Outlook.Application")
    namespace = outlook.GetNamespace("MAPI")
    inbox     = namespace.GetDefaultFolder(OL_FOLDER_INBOX)
    return namespace, inbox


def get_inbox_entry_ids(inbox):
    """
    Returns the set of EntryIDs for all mail items currently in the inbox.
    Equivalent to safe_get_ids() in the IMAP version.
    """
    try:
        ids = set()
        items = inbox.Items
        for item in items:
            try:
                if item.Class == OL_MAIL_ITEM:
                    ids.add(item.EntryID)
            except Exception:
                continue
        return ids
    except Exception as e:
        print(f"Could not reach Outlook inbox: {e}")
        return None


def fetch_item_by_entry_id(namespace, entry_id):
    """
    Fetches a single mail item by EntryID and parses it.
    Equivalent to safe_fetch_email() in the IMAP version.
    """
    try:
        item = namespace.GetItemFromID(entry_id)
        if item.Class == OL_MAIL_ITEM:
            return parse_email_com(item)
    except Exception as e:
        print(f"Could not fetch email: {e}")
    return None


def fetch_items_since(inbox, since_dt):
    """
    Fetches all mail items received after since_dt from the inbox.
    Returns list of (entry_id, parsed_email) tuples sorted oldest first,
    with non-reply items before reply items (same ordering logic as the
    IMAP version's safe_fetch_emails_since(), to ensure original tickets
    are processed before their replies during historical load).
    """
    results = []
    try:
        items = inbox.Items
        items.Sort("[ReceivedTime]", False)  # ascending — oldest first
        # Restrict to items received after since_dt for efficiency
        restrict_str = since_dt.strftime("[ReceivedTime] > '%m/%d/%Y %H:%M %p'")
        try:
            filtered = items.Restrict(restrict_str)
        except Exception:
            filtered = items  # fall back to scanning everything

        for item in filtered:
            try:
                if item.Class != OL_MAIL_ITEM:
                    continue
                em = parse_email_com(item)
                try:
                    em_dt = datetime.strptime(em["received"], "%Y-%m-%d %H:%M:%S")
                except Exception:
                    em_dt = since_dt
                if em_dt > since_dt:
                    results.append((em["_entry_id"], em, em_dt))
            except Exception:
                continue
    except Exception as e:
        print(f"Could not fetch historical emails: {e}")
        return []

    # Sort strictly oldest first by received datetime
    results.sort(key=lambda x: x[2])

    # Originals before replies (same fix as the IMAP version — handles
    # same-minute timestamp ties between an original and its reply)
    no_reply  = [(eid, em, dt) for eid, em, dt in results if not em["in_reply_to"]]
    has_reply = [(eid, em, dt) for eid, em, dt in results if em["in_reply_to"]]
    no_reply.sort(key=lambda x: x[2])
    has_reply.sort(key=lambda x: x[2])
    results = no_reply + has_reply

    return [(eid, em) for eid, em, _ in results]


# ---------------------------------------------------------------------------
# Historical load (Outlook COM version)
# ---------------------------------------------------------------------------

def get_last_record_time(filepath):
    """
    Scans all rows in Excel and returns the latest timestamp across
    both Received Date and Latest Reply Date columns.
    Returns None if file does not exist or has no data rows.
    """
    if not os.path.exists(filepath):
        return None

    try:
        wb = openpyxl.load_workbook(filepath)
        ws = wb["Inbox Tracker"]

        if ws.max_row <= 1:
            return None  # Only header row — treat as empty

        latest_dt = None

        for row in range(2, ws.max_row + 1):
            for col in [COL_RECEIVED, COL_LATEST_DATE]:
                val = ws.cell(row=row, column=col).value
                if not val:
                    continue
                try:
                    dt = datetime.strptime(str(val), "%Y-%m-%d %H:%M:%S")
                    if latest_dt is None or dt > latest_dt:
                        latest_dt = dt
                except Exception:
                    continue

        return latest_dt

    except Exception as e:
        print(f"Could not read last record time: {e}")
        return None


def load_historical_emails_com(inbox, allowed_senders, retry_queue):
    """
    Loads historical emails based on Excel state:
    - No file or empty  -> fetch last 24 hours
    - Has data          -> fetch from last record timestamp to now
    Processes them oldest first so replies link to their original tickets.
    """
    last_time = get_last_record_time(OUTPUT_FILE)

    if last_time is None:
        since_dt = datetime.now() - timedelta(hours=24)
        print(f"No existing data — loading emails from last 24 hours ({since_dt.strftime('%Y-%m-%d %H:%M:%S')} to now)...")
    else:
        since_dt = last_time
        print(f"Existing data found — loading emails since last record ({since_dt.strftime('%Y-%m-%d %H:%M:%S')})...")

    emails = fetch_items_since(inbox, since_dt)

    if not emails:
        print("No new historical emails to load.")
        return allowed_senders

    print(f"Found {len(emails)} historical email(s) to process. Processing oldest first...")
    print()

    for i, (eid, em) in enumerate(emails, 1):
        print(f"[{i}/{len(emails)}] Processing: '{em['subject']}' from {em['sender_email']}")
        if not is_allowed(em["sender_email"], allowed_senders):
            print(f"   Ignored (not in allowed list)")
            continue
        allowed_senders = extract_and_update_senders(em["raw_msg"], allowed_senders)
        process_email(em, retry_queue, OUTPUT_FILE)

    print(f"Historical load complete — {len(emails)} email(s) processed.")
    print()
    return allowed_senders


# ---------------------------------------------------------------------------
# Due date check — runs every poll cycle
# ---------------------------------------------------------------------------

def check_due_date_and_resolve(filepath):
    """
    Due dates are purely informational (displayed in the Due Date column
    for reference) and do NOT drive status changes. Per direction from
    the supervisor: a ticket whose due date has passed should simply
    remain at whatever status it already has (e.g. stay Ongoing) — no
    automatic Overdue status and no auto-resolving. This function is
    kept as a no-op placeholder in case due-date-driven behavior is
    wanted again in the future, and so the call site in the polling
    loop doesn't need to change.
    """
    pass


# ---------------------------------------------------------------------------
# Dashboard generation — Python-computed snapshot (run once/twice a day,
# after the latest batch of emails has been processed). All values and
# charts are static at generation time; re-run the script to refresh.
# ---------------------------------------------------------------------------

DASH_SHEET          = "Dashboard"
DASH_STATUSES       = ["Pending", "Ongoing", "Resolved"]
DASH_STATUS_COLORS  = {"Pending": "FFC000", "Ongoing": "5B9BD5", "Resolved": "70AD47"}

DASH_TITLE_FONT        = Font(bold=True, size=16, name="Arial", color="1F4E79")
DASH_SUBTITLE_FONT     = Font(italic=True, size=9, name="Arial", color="888888")
DASH_SECTION_FONT      = Font(bold=True, size=12, name="Arial", color="FFFFFF")
DASH_SECTION_FILL      = PatternFill("solid", start_color="2E75B6")
DASH_LABEL_FONT        = Font(bold=True, size=10, name="Arial")
DASH_VALUE_FONT        = Font(size=10, name="Arial")
DASH_TABLE_HEADER_FONT = Font(bold=True, size=10, name="Arial", color="FFFFFF")
DASH_TABLE_HEADER_FILL = PatternFill("solid", start_color="44546A")
DASH_URGENT_FILL       = PatternFill("solid", start_color="FCE4D6")  # long-running Ongoing tickets


def _dash_set_title(title_holder, size_pt=14, bold=True, color="1F1F1F"):
    """
    Applies an explicit font size and overlay=False to a chart's or
    axis's EXISTING title (set beforehand via chart.title = "..." or
    axis.title = "..."). Modifies the existing paragraph/run in place
    rather than replacing it, so the original title text is preserved.

    Newer openpyxl versions (3.1.4+) leave title font size at a tiny
    default and don't explicitly set overlay=False, which makes real
    Excel render titles too small and stacked on top of the chart /
    axis tick labels instead of placed above/beside them (other
    renderers like LibreOffice are more forgiving and hide this bug).
    """
    if title_holder.title is None or title_holder.title.tx is None:
        return
    para = title_holder.title.tx.rich.p[0]
    cp = CharacterProperties(sz=size_pt * 100, b=bold, solidFill=color)
    para.pPr = ParagraphProperties(defRPr=cp)
    for run in para.r:
        run.rPr = cp
    title_holder.title.overlay = False


def _dash_section_header(ws, row, text, span=6):
    cell = ws.cell(row=row, column=2, value=text)
    cell.font = DASH_SECTION_FONT
    cell.fill = DASH_SECTION_FILL
    for c in range(2, 2 + span):
        ws.cell(row=row, column=c).fill = DASH_SECTION_FILL
    return row + 1


def _dash_parse_dt(value, fmt="%Y-%m-%d %H:%M:%S"):
    if not value:
        return None
    if isinstance(value, datetime):
        return value
    try:
        return datetime.strptime(str(value), fmt)
    except Exception:
        return None


def _dash_read_tickets(data_ws):
    """Reads all data rows into a list of dicts for easy aggregation."""
    tickets = []
    for row in range(2, data_ws.max_row + 1):
        num = data_ws.cell(row=row, column=COL_NUM).value
        if num is None:
            continue
        tickets.append({
            "num":           num,
            "sender_name":   data_ws.cell(row=row, column=COL_SENDER_NAME).value or "",
            "sender_email":  data_ws.cell(row=row, column=COL_SENDER_EMAIL).value or "",
            "subject":       data_ws.cell(row=row, column=COL_SUBJECT).value or "",
            "category":      data_ws.cell(row=row, column=COL_CATEGORY).value or "",
            "received":      data_ws.cell(row=row, column=COL_RECEIVED).value or "",
            "latest":        data_ws.cell(row=row, column=COL_LATEST_DATE).value or "",
            "status":        data_ws.cell(row=row, column=COL_STATUS).value or "",
            "ongoing_since": data_ws.cell(row=row, column=COL_ONGOING_SINCE).value or "",
            "assigned_to":   data_ws.cell(row=row, column=COL_ASSIGNED_TO).value or "",
            "followups":     data_ws.cell(row=row, column=COL_FOLLOWUPS).value or 0,
            "due_date":      data_ws.cell(row=row, column=COL_DUE_DATE).value or "",
        })
    return tickets


def _dash_status_counts(tickets):
    counts = Counter(t["status"] for t in tickets)
    return {s: counts.get(s, 0) for s in DASH_STATUSES}


def _dash_category_breakdown(tickets, categories):
    """
    For each category, counts how many tickets are Pending and how many
    are Ongoing. A ticket with multiple comma-separated categories counts
    toward EACH of its categories (matches how detect_categories() stores
    multi-label results). Returns {category: {"Pending": n, "Ongoing": n}}.
    """
    breakdown = {cat: {"Pending": 0, "Ongoing": 0} for cat in categories}
    for t in tickets:
        if t["status"] not in ("Pending", "Ongoing"):
            continue
        raw_cats = [c.strip() for c in (t["category"] or "").split(",") if c.strip()]
        for cat in raw_cats:
            if cat in breakdown:
                breakdown[cat][t["status"]] += 1
    return breakdown


def _dash_assignee_breakdown(tickets):
    """
    For each assignee, counts Pending and Ongoing tickets. A ticket can
    have multiple comma-separated assignees (credited to each). Tickets
    with no assignee are grouped under "Unassigned". Returns a dict
    sorted by total (Pending+Ongoing) descending, capped at the top 12
    to keep the chart readable.
    """
    breakdown = defaultdict(lambda: {"Pending": 0, "Ongoing": 0})
    for t in tickets:
        if t["status"] not in ("Pending", "Ongoing"):
            continue
        raw_names = [a.strip() for a in (t["assigned_to"] or "").split(",") if a.strip()]
        if not raw_names:
            raw_names = ["Unassigned"]
        for name in raw_names:
            breakdown[name][t["status"]] += 1

    sorted_items = sorted(breakdown.items(), key=lambda kv: -(kv[1]["Pending"] + kv[1]["Ongoing"]))
    return dict(sorted_items[:12])


def _dash_ongoing_list(tickets, now):
    """
    Returns Ongoing tickets sorted by days-in-Ongoing descending (whole
    days, computed from the Ongoing Since timestamp to now). Tickets
    missing Ongoing Since (shouldn't normally happen for an Ongoing
    ticket, but handled defensively) are treated as 0 days.
    """
    ongoing = [t for t in tickets if t["status"] == "Ongoing"]
    enriched = []
    for t in ongoing:
        since_dt = _dash_parse_dt(t["ongoing_since"])
        days = (now - since_dt).days if since_dt else 0
        enriched.append({**t, "days_ongoing": days})
    enriched.sort(key=lambda t: -t["days_ongoing"])
    return enriched


def update_dashboard(filepath):
    """
    Rebuilds the 'Dashboard' sheet as a static snapshot of the current
    'Inbox Tracker' data: overview counts, a status pie chart, grouped
    Pending/Ongoing bar charts by category and by assignee, and a list
    of all Ongoing tickets sorted by how many days they've been Ongoing
    (rows ongoing 14+ days are highlighted). Meant to be run once or
    twice a day, not on every poll cycle — everything here is computed
    fresh from scratch each call, so it's always consistent with
    whatever is in the data sheet at the time it's run.

    Returns True on success, False if the file is locked (e.g. open in
    Excel) or the data sheet is missing.
    """
    try:
        wb = openpyxl.load_workbook(filepath)
    except PermissionError:
        print("Could not update dashboard — Excel is open.")
        return False
    except Exception as e:
        print(f"Could not update dashboard: {e}")
        return False

    if "Inbox Tracker" not in wb.sheetnames:
        print("Could not update dashboard — 'Inbox Tracker' sheet not found.")
        return False

    data_ws    = wb["Inbox Tracker"]
    tickets    = _dash_read_tickets(data_ws)
    now        = datetime.now()
    categories = list(LABEL_TO_FULL.values())

    if DASH_SHEET in wb.sheetnames:
        del wb[DASH_SHEET]
    ws = wb.create_sheet(DASH_SHEET, 0)
    ws.sheet_view.showGridLines = False

    # Landscape + fit-to-width so printing/exporting to PDF doesn't cut
    # off columns — purely a print/export concern, doesn't affect the
    # normal on-screen Excel view.
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True

    status_counts      = _dash_status_counts(tickets)
    cat_breakdown       = _dash_category_breakdown(tickets, categories)
    assignee_breakdown  = _dash_assignee_breakdown(tickets)
    ongoing_list        = _dash_ongoing_list(tickets, now)

    total = len(tickets)
    resolution_rate = (status_counts["Resolved"] / total) if total else 0
    avg_followups   = (sum(t["followups"] for t in tickets) / total) if total else 0

    # -------------------------------------------------------------
    # Title
    # -------------------------------------------------------------
    ws["B2"] = "TechOps Ticket Dashboard"
    ws["B2"].font = DASH_TITLE_FONT
    ws["B3"] = f"Generated: {now.strftime('%Y-%m-%d %H:%M')}"
    ws["B3"].font = DASH_SUBTITLE_FONT

    # -------------------------------------------------------------
    # Section 1: Overview counts
    # -------------------------------------------------------------
    row = 5
    row = _dash_section_header(ws, row, "Overview")
    items = [
        ("Total Tickets", total, None),
        ("Pending", status_counts["Pending"], None),
        ("Ongoing", status_counts["Ongoing"], None),
        ("Resolved", status_counts["Resolved"], None),
        ("Resolution Rate", resolution_rate, "0.0%"),
        ("Avg Followups / Ticket", avg_followups, "0.00"),
    ]
    for label, value, numfmt in items:
        ws.cell(row=row, column=2, value=label).font = DASH_LABEL_FONT
        c = ws.cell(row=row, column=4, value=value)
        c.font = DASH_VALUE_FONT
        if numfmt:
            c.number_format = numfmt
        row += 1
    row += 1

    # -------------------------------------------------------------
    # Section 2: Status pie chart — chart on the left (the main thing
    # to glance at), small data table to its right for exact numbers.
    # -------------------------------------------------------------
    row = _dash_section_header(ws, row, "Status Breakdown")
    status_table_row = row
    ws.cell(row=row, column=6, value="Status").font = DASH_LABEL_FONT
    ws.cell(row=row, column=7, value="Count").font = DASH_LABEL_FONT
    row += 1
    for status in DASH_STATUSES:
        ws.cell(row=row, column=6, value=status).font = DASH_VALUE_FONT
        ws.cell(row=row, column=7, value=status_counts[status]).font = DASH_VALUE_FONT
        row += 1
    status_table_end = row - 1

    pie = PieChart()
    pie.title = "Tickets by Status"
    data_ref = Reference(ws, min_col=7, min_row=status_table_row, max_row=status_table_end)
    cats_ref = Reference(ws, min_col=6, min_row=status_table_row + 1, max_row=status_table_end)
    pie.add_data(data_ref, titles_from_data=True)
    pie.set_categories(cats_ref)
    pie.height = 7.5
    pie.width = 11
    pie.dataLabels = DataLabelList()
    pie.dataLabels.showVal = False
    pie.dataLabels.showPercent = True
    pie.dataLabels.showCatName = False
    pie.dataLabels.showLegendKey = False
    pie.dataLabels.showSerName = False
    _dash_set_title(pie, size_pt=14)
    ws.add_chart(pie, f"B{status_table_row}")

    row = max(row, status_table_row + 16) + 1

    # -------------------------------------------------------------
        # -------------------------------------------------------------
    # Section 3: Category breakdown — grouped bar chart (Pending/Ongoing).
    # Chart on the left, data table further right (this chart is wider
    # than the pie chart, so the table needs to start further over).
    # -------------------------------------------------------------
    row = _dash_section_header(ws, row, "Tickets by Category (Pending vs Ongoing)")
    cat_table_row = row
    ws.cell(row=row, column=10, value="Category").font = DASH_LABEL_FONT
    ws.cell(row=row, column=11, value="Pending").font = DASH_LABEL_FONT
    ws.cell(row=row, column=12, value="Ongoing").font = DASH_LABEL_FONT
    row += 1
    for cat in categories:
        ws.cell(row=row, column=10, value=cat).font = DASH_VALUE_FONT
        ws.cell(row=row, column=11, value=cat_breakdown[cat]["Pending"]).font = DASH_VALUE_FONT
        ws.cell(row=row, column=12, value=cat_breakdown[cat]["Ongoing"]).font = DASH_VALUE_FONT
        row += 1
    cat_table_end = row - 1

    cat_bar = BarChart()
    cat_bar.type = "col"
    cat_bar.grouping = "clustered"
    cat_bar.title = "Tickets by Category"
    cat_bar.y_axis.title = "Number of Tickets"
    data_ref = Reference(ws, min_col=11, max_col=12, min_row=cat_table_row, max_row=cat_table_end)
    cats_ref = Reference(ws, min_col=10, min_row=cat_table_row + 1, max_row=cat_table_end)
    cat_bar.add_data(data_ref, titles_from_data=True)
    cat_bar.set_categories(cats_ref)
    cat_bar.series[0].graphicalProperties.solidFill = DASH_STATUS_COLORS["Pending"]
    cat_bar.series[1].graphicalProperties.solidFill = DASH_STATUS_COLORS["Ongoing"]
    cat_bar.height = 9
    cat_bar.width = 20
    cat_bar.x_axis.textRotation = -30
    _dash_set_title(cat_bar, size_pt=14)
    _dash_set_title(cat_bar.y_axis, size_pt=10)
    # openpyxl 3.1.4+ leaves axes implicitly "deleted" in the chart XML
    # unless explicitly told otherwise, which makes real Excel hide both
    # axis lines AND their labels entirely (other renderers like
    # LibreOffice/VS Code's preview show them anyway, masking the bug).
    cat_bar.x_axis.delete = False
    cat_bar.y_axis.delete = False
    ws.add_chart(cat_bar, f"B{cat_table_row}")

    row = max(row, cat_table_row + 19) + 1

    # -------------------------------------------------------------
        # -------------------------------------------------------------
    # Section 4: Assignee breakdown — grouped bar chart (Pending/Ongoing).
    # Chart on the left, data table further right.
    # -------------------------------------------------------------
    row = _dash_section_header(ws, row, "Tickets by Assignee (Pending vs Ongoing)")
    assignee_table_row = row
    ws.cell(row=row, column=10, value="Assignee").font = DASH_LABEL_FONT
    ws.cell(row=row, column=11, value="Pending").font = DASH_LABEL_FONT
    ws.cell(row=row, column=12, value="Ongoing").font = DASH_LABEL_FONT
    row += 1
    for name, counts in assignee_breakdown.items():
        ws.cell(row=row, column=10, value=name).font = DASH_VALUE_FONT
        ws.cell(row=row, column=11, value=counts["Pending"]).font = DASH_VALUE_FONT
        ws.cell(row=row, column=12, value=counts["Ongoing"]).font = DASH_VALUE_FONT
        row += 1
    assignee_table_end = row - 1

    assignee_bar = BarChart()
    assignee_bar.type = "col"
    assignee_bar.grouping = "clustered"
    assignee_bar.title = "Tickets by Assignee"
    assignee_bar.y_axis.title = "Number of Tickets"
    data_ref = Reference(ws, min_col=11, max_col=12, min_row=assignee_table_row, max_row=assignee_table_end)
    cats_ref = Reference(ws, min_col=10, min_row=assignee_table_row + 1, max_row=assignee_table_end)
    assignee_bar.add_data(data_ref, titles_from_data=True)
    assignee_bar.set_categories(cats_ref)
    assignee_bar.series[0].graphicalProperties.solidFill = DASH_STATUS_COLORS["Pending"]
    assignee_bar.series[1].graphicalProperties.solidFill = DASH_STATUS_COLORS["Ongoing"]
    assignee_bar.height = 9
    assignee_bar.width = 18
    assignee_bar.x_axis.textRotation = -30
    assignee_bar.x_axis.delete = False
    assignee_bar.y_axis.delete = False
    _dash_set_title(assignee_bar, size_pt=14)
    _dash_set_title(assignee_bar.y_axis, size_pt=10)
    ws.add_chart(assignee_bar, f"B{assignee_table_row}")

    row = max(row, assignee_table_row + 19) + 1

    # -------------------------------------------------------------
    # Section 5: Ongoing tickets list, sorted by days-in-Ongoing desc
    # -------------------------------------------------------------
    row = _dash_section_header(ws, row, f"Ongoing Tickets ({len(ongoing_list)}) — Sorted by Days Ongoing", span=7)
    list_header_row = row
    list_headers = ["Ticket #", "Subject", "Assignee", "Category", "Days Ongoing", "Sender"]
    for i, h in enumerate(list_headers):
        c = ws.cell(row=row, column=2 + i, value=h)
        c.font = DASH_TABLE_HEADER_FONT
        c.fill = DASH_TABLE_HEADER_FILL
        c.alignment = Alignment(horizontal="center", vertical="center")
    row += 1

    thin = Side(style="thin", color="D9D9D9")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    if not ongoing_list:
        ws.cell(row=row, column=2, value="No tickets are currently Ongoing.").font = DASH_VALUE_FONT
        row += 1
    else:
        for t in ongoing_list:
            values = [t["num"], t["subject"], t["assigned_to"] or "Unassigned",
                      t["category"], t["days_ongoing"], t["sender_email"]]
            long_ongoing = t["days_ongoing"] >= 14
            row_fill = DASH_URGENT_FILL if long_ongoing else PatternFill()
            for col_offset, val in enumerate(values):
                c = ws.cell(row=row, column=2 + col_offset, value=val)
                c.font = Font(name="Arial", size=10, bold=long_ongoing)
                c.fill = row_fill
                c.border = border
                if col_offset in (0, 4):  # Ticket # and Days Ongoing read better centered
                    c.alignment = Alignment(horizontal="center", vertical="top")
                else:
                    c.alignment = Alignment(vertical="top", wrap_text=(col_offset == 1))
            row += 1

    # Column widths — shared across all sections on this sheet. B-G serve
    # the Overview section, the Status table (F/G), and the Ongoing
    # Tickets list (which needs the most width, for subjects/emails).
    # H is left at default width as a visual buffer so the Category/
    # Assignee tables (J/K) don't crowd the wider bar charts next to them.
    ws.column_dimensions["A"].width = 2
    ws.column_dimensions["B"].width = 30
    ws.column_dimensions["C"].width = 38
    ws.column_dimensions["D"].width = 28
    ws.column_dimensions["E"].width = 35
    ws.column_dimensions["F"].width = 13
    ws.column_dimensions["G"].width = 30
    ws.column_dimensions["J"].width = 32
    ws.column_dimensions["K"].width = 10
    ws.column_dimensions["L"].width = 10

    # Inserting the Dashboard sheet at index 0 makes it the active sheet
    # by default — explicitly restore "Inbox Tracker" as active so the
    # file always opens on the data sheet, and so nothing downstream
    # that might rely on wb.active gets the wrong sheet.
    wb.active = wb.sheetnames.index("Inbox Tracker")

    try:
        wb.save(filepath)
    except PermissionError:
        print("Could not save dashboard — Excel is open.")
        return False

    return True


# ---------------------------------------------------------------------------
# Main listener loop
# ---------------------------------------------------------------------------

def listen():
    print("\nConnecting to Outlook (Classic) desktop app...")

    # pythoncom.CoInitialize() is needed when COM is used from a thread
    # other than the one that created it. We're single-threaded here, but
    # calling it is harmless and protects against future changes.
    pythoncom.CoInitialize()

    try:
        namespace, inbox = connect_outlook()
    except Exception as e:
        print(f"Could not connect to Outlook: {e}")
        print("Make sure Outlook (Classic) is installed and you are signed in.")
        return

    known_ids       = get_inbox_entry_ids(inbox)
    poll_count      = 0
    retry_queue     = []
    allowed_senders = load_allowed_senders()

    if known_ids is None:
        print("Could not connect. Please check that Outlook is running and signed in.")
        return

    print(f"Connected. {len(known_ids)} existing email(s) in inbox will be ignored.")
    print(f"Loaded {len(allowed_senders)} allowed sender(s).")

    # Snapshot known_ids BEFORE historical load so any emails that arrive
    # during historical processing are caught in the first live poll cycle
    print("Checking for historical emails to load...")
    allowed_senders = load_historical_emails_com(inbox, allowed_senders, retry_queue)

    print("Updating dashboard...")
    if update_dashboard(OUTPUT_FILE):
        print("Dashboard updated.\n")
    else:
        print("Could not update dashboard right now — you can re-run the script later to refresh it.\n")

    print(f"Starting live listener — checking every {POLL_INTERVAL} seconds. Press Ctrl+C to stop.\n")

    while True:
        try:
            time.sleep(POLL_INTERVAL)
            poll_count += 1

            if retry_queue:
                still_failed = []
                for em in retry_queue:
                    if em["in_reply_to"]:
                        wb            = openpyxl.load_workbook(OUTPUT_FILE)
                        ws            = wb["Inbox Tracker"]
                        target_row, _ = find_matching_row(em, ws)
                        if target_row:
                            result = update_existing_row(em, target_row, OUTPUT_FILE)
                            if result == "new_issue":
                                saved = append_new_email(em, OUTPUT_FILE)
                                if not saved:
                                    still_failed.append(em)
                                else:
                                    print(f"Saved new issue from queue: '{em['subject']}'")
                            elif result is True:
                                print(f"Updated from queue: '{em['subject']}'")
                            else:
                                still_failed.append(em)
                        else:
                            saved = append_new_email(em, OUTPUT_FILE)
                            if saved:
                                print(f"Saved from queue: '{em['subject']}'")
                            else:
                                still_failed.append(em)
                    else:
                        saved = append_new_email(em, OUTPUT_FILE)
                        if saved:
                            print(f"Saved from queue: '{em['subject']}'")
                        else:
                            still_failed.append(em)
                retry_queue = still_failed
                if retry_queue:
                    print(f"Excel is still open — {len(retry_queue)} email(s) still waiting.")
                else:
                    # All queued emails saved — refresh dashboard
                    update_dashboard(OUTPUT_FILE)

            check_due_date_and_resolve(OUTPUT_FILE)

            current_ids = get_inbox_entry_ids(inbox)

            if current_ids is None:
                print(f"Could not reach Outlook this cycle, will try again in {POLL_INTERVAL}s...")
                continue

            new_ids = current_ids - known_ids

            if new_ids:
                for eid in new_ids:
                    em = fetch_item_by_entry_id(namespace, eid)
                    if em:
                        if not is_allowed(em["sender_email"], allowed_senders):
                            print(f"Ignored email from {em['sender_email']} (not in allowed list)")
                            continue
                        allowed_senders = extract_and_update_senders(em["raw_msg"], allowed_senders)
                        process_email(em, retry_queue, OUTPUT_FILE)
                known_ids = current_ids
                # Refresh dashboard now that new data has been processed —
                # only called once per poll cycle even if multiple emails
                # arrived, to avoid redundant rebuilds.
                update_dashboard(OUTPUT_FILE)

            else:
                if poll_count % 4 == 0:
                    queued_note = f"  ({len(retry_queue)} emails queued)" if retry_queue else ""
                    print(f"Listening...  {datetime.now().strftime('%H:%M:%S')}{queued_note}")

        except KeyboardInterrupt:
            if retry_queue:
                print(f"\nNote: {len(retry_queue)} email(s) were never saved because Excel was open:")
                for em in retry_queue:
                    print(f"   - '{em['subject']}' from {em['sender_email']}")
            print("\nUpdating dashboard one last time before stopping...")
            if update_dashboard(OUTPUT_FILE):
                print("Dashboard updated.")
            else:
                print("Could not update dashboard — you can re-run the script later to refresh it.")
            print("Stopped.")
            break


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

def main():
    print("Outlook (Classic) Mail Listener\n")
    init_excel(OUTPUT_FILE)
    listen()


if __name__ == "__main__":
    main()
