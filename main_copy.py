import os
import time
import uuid
from datetime import datetime
from html import escape

import openpyxl
import streamlit as st
from dotenv import load_dotenv
from google import genai


# ============================================================
# CONFIGURATION
# ============================================================

load_dotenv()

st.set_page_config(
    page_title="PY BANK • AI Banking",
    page_icon="🏦",
    layout="wide",
    initial_sidebar_state="expanded",
)

FILE_NAME = "bank_records_1.xlsx"
GEMINI_MODEL = "gemini-3.8-flash"


# ============================================================
# HELPERS FOR HTML
# ============================================================

def render_html(content):
    """
    Render actual HTML with Streamlit's dedicated HTML renderer.
    This prevents HTML from appearing as literal text.
    """
    st.html(content)


def safe(value):
    """Safely insert dynamic values into HTML."""
    return escape(str(value or ""))


# ============================================================
# CUSTOM CSS
# ============================================================

CSS = r"""
<style>

@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&family=Space+Grotesk:wght@500;600;700&display=swap');

:root {
    --bg: #06101c;
    --bg2: #09172a;
    --card: rgba(13, 28, 49, 0.76);
    --card2: rgba(17, 36, 61, 0.62);
    --line: rgba(255,255,255,.09);
    --text: #eff7ff;
    --muted: #91a6c1;
    --cyan: #5de4ff;
    --violet: #9f7cff;
    --green: #56edb1;
    --red: #ff7298;
}

html,
body,
[class*="css"] {
    font-family: 'Inter', sans-serif;
}

.stApp {
    background:
        radial-gradient(
            circle at 7% 8%,
            rgba(93,228,255,.13),
            transparent 26%
        ),
        radial-gradient(
            circle at 93% 10%,
            rgba(159,124,255,.15),
            transparent 28%
        ),
        radial-gradient(
            circle at 55% 100%,
            rgba(86,237,177,.07),
            transparent 32%
        ),
        linear-gradient(
            135deg,
            var(--bg),
            var(--bg2)
        );
    color: var(--text);
}


/* =========================================================
   IMPORTANT FIX:
   Keep the application below Streamlit's top header.
   ========================================================= */

[data-testid="stHeader"] {
    background: rgba(6, 16, 28, 0.96) !important;
    border-bottom: 1px solid rgba(255,255,255,.06);
    backdrop-filter: blur(14px);
    z-index: 999999 !important;
}

[data-testid="stAppViewContainer"] {
    background: transparent;
}

.main .block-container,
[data-testid="stAppViewContainer"] .main .block-container,
section.main > div.block-container {
    max-width: 1450px !important;
    padding-top: 5.8rem !important;
    padding-bottom: 2.5rem !important;
}


/* =========================================================
   SIDEBAR
   ========================================================= */

[data-testid="stSidebar"] {
    background:
        linear-gradient(
            180deg,
            rgba(5,14,25,.98),
            rgba(8,21,37,.96)
        );
    border-right: 1px solid var(--line);
}

[data-testid="stSidebar"] * {
    color: var(--text);
}


/* =========================================================
   BRAND
   ========================================================= */

.brand {
    display: flex;
    align-items: center;
    gap: 12px;
    margin: 4px 0 20px;
}

.brand-icon {
    width: 48px;
    height: 48px;
    border-radius: 15px;

    display: grid;
    place-items: center;

    font-size: 25px;

    background:
        linear-gradient(
            135deg,
            rgba(93,228,255,.18),
            rgba(159,124,255,.28)
        );

    border: 1px solid rgba(255,255,255,.12);

    box-shadow:
        0 12px 30px rgba(93,228,255,.12);
}

.brand-name {
    font-family: 'Space Grotesk', sans-serif;
    font-size: 1.32rem;
    font-weight: 800;
    letter-spacing: .5px;
}

.brand-sub {
    color: var(--muted);
    font-size: .72rem;
}


/* =========================================================
   HERO
   ========================================================= */

.hero {
    position: relative;
    overflow: hidden;

    padding: 28px 30px;

    border-radius: 28px;

    background:
        linear-gradient(
            135deg,
            rgba(93,228,255,.09),
            rgba(159,124,255,.13)
        ),
        rgba(9,22,39,.77);

    border: 1px solid var(--line);

    box-shadow:
        0 28px 80px rgba(0,0,0,.34);

    backdrop-filter: blur(18px);

    margin-bottom: 20px;
}

.hero::after {
    content: "";

    position: absolute;

    width: 240px;
    height: 240px;

    right: -70px;
    top: -85px;

    border-radius: 50%;

    background:
        radial-gradient(
            circle,
            rgba(93,228,255,.17),
            transparent 70%
        );

    pointer-events: none;
}

.hero-kicker {
    color: var(--cyan);

    text-transform: uppercase;

    letter-spacing: 2px;

    font-size: .70rem;

    font-weight: 800;
}

.hero-title {
    font-family: 'Space Grotesk', sans-serif;

    font-size: clamp(
        2.1rem,
        4vw,
        3.25rem
    );

    font-weight: 700;

    line-height: 1.02;

    margin: 7px 0 9px;
}

.hero-copy {
    color: var(--muted);

    max-width: 760px;

    line-height: 1.65;

    font-size: .94rem;
}

.hero-row {
    display: flex;

    justify-content: space-between;

    align-items: center;

    gap: 18px;

    flex-wrap: wrap;
}

.live-chip {
    display: inline-flex;

    align-items: center;

    gap: 7px;

    border-radius: 999px;

    padding: 9px 13px;

    color: var(--green);

    background:
        rgba(86,237,177,.07);

    border:
        1px solid rgba(86,237,177,.23);

    font-size: .76rem;

    font-weight: 800;

    white-space: nowrap;
}

.live-dot {
    width: 8px;
    height: 8px;

    border-radius: 50%;

    background: var(--green);

    box-shadow:
        0 0 15px rgba(86,237,177,.9);
}


/* =========================================================
   GLASS
   ========================================================= */

.glass {
    background: var(--card);

    border:
        1px solid var(--line);

    border-radius: 21px;

    padding: 19px;

    box-shadow:
        0 18px 55px rgba(0,0,0,.20);

    backdrop-filter: blur(16px);
}


/* =========================================================
   SECTIONS
   ========================================================= */

.section-title {
    font-family: 'Space Grotesk', sans-serif;

    font-size: 1.2rem;

    font-weight: 700;

    margin: 17px 0 11px;
}


/* =========================================================
   FEATURE CARDS
   ========================================================= */

.feature-card {
    min-height: 145px;

    padding: 18px;

    border-radius: 19px;

    background:
        linear-gradient(
            145deg,
            rgba(93,228,255,.065),
            rgba(159,124,255,.075)
        );

    border:
        1px solid var(--line);

    transition:
        transform .2s ease,
        border-color .2s ease;
}

.feature-card:hover {
    transform: translateY(-2px);

    border-color:
        rgba(93,228,255,.25);
}

.feature-icon {
    font-size: 1.55rem;

    margin-bottom: 7px;
}

.feature-title {
    font-weight: 800;

    font-size: .97rem;
}

.feature-copy {
    color: var(--muted);

    font-size: .77rem;

    line-height: 1.55;

    margin-top: 5px;
}


/* =========================================================
   BALANCE
   ========================================================= */

.balance-card {
    padding: 22px;

    border-radius: 23px;

    background:
        linear-gradient(
            135deg,
            rgba(93,228,255,.12),
            rgba(159,124,255,.14)
        ),
        rgba(12,27,48,.82);

    border:
        1px solid rgba(255,255,255,.11);

    box-shadow:
        0 20px 60px rgba(0,0,0,.27);
}

.balance-label {
    color: var(--muted);

    font-size: .74rem;

    text-transform: uppercase;

    letter-spacing: 1.4px;

    font-weight: 800;
}

.balance-value {
    font-family: 'Space Grotesk', sans-serif;

    font-size: clamp(
        2.25rem,
        5vw,
        3.4rem
    );

    line-height: 1;

    font-weight: 700;

    margin: 7px 0 12px;
}

.account-chip {
    display: inline-block;

    border-radius: 999px;

    padding: 6px 10px;

    font-size: .69rem;

    border:
        1px solid rgba(255,255,255,.10);

    background:
        rgba(255,255,255,.04);

    color: #b8c7d9;
}


/* =========================================================
   METRICS
   ========================================================= */

.metric-card {
    padding: 18px;

    border-radius: 20px;

    background: var(--card2);

    border:
        1px solid var(--line);
}

.metric-label {
    color: var(--muted);

    font-size: .69rem;

    text-transform: uppercase;

    letter-spacing: 1.2px;

    font-weight: 800;
}

.metric-number {
    font-family: 'Space Grotesk', sans-serif;

    font-size: 1.78rem;

    font-weight: 700;

    margin-top: 5px;
}

.pill {
    display: inline-block;

    margin-top: 8px;

    padding: 4px 8px;

    border-radius: 999px;

    font-size: .65rem;

    font-weight: 800;
}

.pill-green {
    color: var(--green);

    background:
        rgba(86,237,177,.07);

    border:
        1px solid rgba(86,237,177,.16);
}

.pill-cyan {
    color: var(--cyan);

    background:
        rgba(93,228,255,.07);

    border:
        1px solid rgba(93,228,255,.16);
}

.pill-red {
    color: var(--red);

    background:
        rgba(255,114,152,.07);

    border:
        1px solid rgba(255,114,152,.16);
}


/* =========================================================
   TRANSACTIONS
   ========================================================= */

.transaction-card {
    padding: 14px 16px;

    margin-bottom: 9px;

    border-radius: 16px;

    background:
        rgba(255,255,255,.028);

    border:
        1px solid var(--line);
}

.transaction-row {
    display: flex;

    justify-content: space-between;

    align-items: center;

    gap: 12px;
}

.transaction-name {
    font-weight: 800;

    font-size: .87rem;
}

.transaction-meta {
    color: var(--muted);

    font-size: .68rem;

    margin-top: 4px;
}

.transaction-money {
    font-family: 'Space Grotesk', sans-serif;

    font-size: .93rem;

    font-weight: 700;

    white-space: nowrap;
}

.money-in {
    color: var(--green);
}

.money-out {
    color: var(--red);
}


/* =========================================================
   AI
   ========================================================= */

.ai-card {
    padding: 20px;

    border-radius: 22px;

    background:
        linear-gradient(
            145deg,
            rgba(159,124,255,.11),
            rgba(93,228,255,.08)
        ),
        rgba(14,26,47,.78);

    border:
        1px solid rgba(159,124,255,.18);

    box-shadow:
        0 20px 60px rgba(0,0,0,.22);
}

.ai-badge {
    display: inline-flex;

    align-items: center;

    gap: 7px;

    border-radius: 999px;

    padding: 6px 10px;

    background:
        rgba(159,124,255,.09);

    color: #c7b9ff;

    border:
        1px solid rgba(159,124,255,.18);

    font-size: .67rem;

    font-weight: 800;
}


/* =========================================================
   FOOTER
   ========================================================= */

.footer {
    color: #6e829b;

    font-size: .70rem;

    text-align: center;

    padding: 28px 0 8px;
}


/* =========================================================
   STREAMLIT BUTTONS
   ========================================================= */

div[data-testid="stButton"] > button,
div[data-testid="stFormSubmitButton"] > button {
    border-radius: 13px;

    min-height: 44px;

    border:
        1px solid rgba(255,255,255,.10);

    background:
        linear-gradient(
            135deg,
            rgba(93,228,255,.11),
            rgba(159,124,255,.12)
        );

    color: #f3f8ff;

    font-weight: 800;

    transition: .18s ease;
}

div[data-testid="stButton"] > button:hover,
div[data-testid="stFormSubmitButton"] > button:hover {
    transform: translateY(-1px);

    border-color:
        rgba(93,228,255,.30);

    box-shadow:
        0 9px 25px rgba(93,228,255,.10);
}


/* =========================================================
   INPUTS
   ========================================================= */

div[data-testid="stTextInput"] input,
div[data-testid="stNumberInput"] input {
    background:
        rgba(5,15,28,.72);

    color: #ffffff;

    border:
        1px solid rgba(255,255,255,.10);

    border-radius: 12px;
}


/* =========================================================
   TABS
   ========================================================= */

.stTabs [data-baseweb="tab"] {
    border-radius: 11px;

    padding: 7px 13px;

    background:
        rgba(255,255,255,.025);

    border:
        1px solid var(--line);
}


/* =========================================================
   MOBILE
   ========================================================= */

@media (max-width: 900px) {

    .main .block-container,
    [data-testid="stAppViewContainer"] .main .block-container,
    section.main > div.block-container {
        padding-left: 1rem !important;
        padding-right: 1rem !important;
        padding-top: 5.2rem !important;
    }

    .hero {
        padding: 22px 20px;
    }

    .hero-title {
        font-size: 2.15rem;
    }

    .transaction-row {
        align-items: flex-start;
    }

}

</style>
"""

render_html(CSS)


# ============================================================
# GEMINI SETUP
# ============================================================

api_key = os.getenv("GEMINI_API_KEY")

try:
    if not api_key:
        api_key = st.secrets["GEMINI_API_KEY"]
except Exception:
    pass

if api_key:
    client = genai.Client(api_key=api_key)
    gemini_ready = True
else:
    client = None
    gemini_ready = False


# ============================================================
# EXCEL DATABASE
# ============================================================

def create_excel_file():
    if not os.path.exists(FILE_NAME):

        wb = openpyxl.Workbook()

        sheet = wb.active
        sheet.title = "Bank Records"

        headers = [
            "Account No",
            "Name",
            "PIN",
            "Transaction ID",
            "Transaction Type",
            "Amount",
            "Previous Balance",
            "Current Balance",
            "Date-Time",
        ]

        sheet.append(headers)

        for cell in sheet[1]:
            cell.font = openpyxl.styles.Font(
                bold=True
            )

        sheet.freeze_panes = "A2"

        wb.save(FILE_NAME)
        wb.close()


create_excel_file()


def now_text():
    return datetime.now().strftime(
        "%d-%m-%Y %H:%M:%S"
    )


def transaction_id():
    return str(uuid.uuid4())[:8].upper()


def account_exists(account_no):

    wb = openpyxl.load_workbook(
        FILE_NAME,
        read_only=True,
        data_only=True,
    )

    sheet = wb["Bank Records"]

    for row in sheet.iter_rows(
        min_row=2,
        values_only=True,
    ):

        if (
            row[0] is not None
            and str(row[0]) == str(account_no)
        ):

            wb.close()
            return True

    wb.close()

    return False


def get_account(account_no):

    wb = openpyxl.load_workbook(
        FILE_NAME,
        read_only=True,
        data_only=True,
    )

    sheet = wb["Bank Records"]

    for row in sheet.iter_rows(
        min_row=2,
        values_only=True,
    ):

        if (
            row[0] is not None
            and str(row[0]) == str(account_no)
            and row[1] is not None
        ):

            result = {
                "account_no": str(row[0]),
                "name": str(row[1]),
                "pin": (
                    ""
                    if row[2] is None
                    else str(row[2])
                ),
            }

            wb.close()

            return result

    wb.close()

    return None


def authenticate(account_no, pin):

    account = get_account(
        account_no
    )

    if (
        account
        and account["pin"] == str(pin)
    ):
        return account

    return None


def get_balance(account_no):

    wb = openpyxl.load_workbook(
        FILE_NAME,
        read_only=True,
        data_only=True,
    )

    sheet = wb["Bank Records"]

    balance = None

    for row in sheet.iter_rows(
        min_row=2,
        values_only=True,
    ):

        if (
            row[0] is not None
            and str(row[0]) == str(account_no)
        ):

            try:
                balance = float(row[7])
            except (
                TypeError,
                ValueError,
            ):
                balance = 0.0

    wb.close()

    return balance


def add_transaction(
    account_no,
    name,
    tx_type,
    amount,
    previous_balance,
    current_balance,
):

    tx_id = transaction_id()
    dt = now_text()

    wb = openpyxl.load_workbook(
        FILE_NAME
    )

    sheet = wb["Bank Records"]

    sheet.append(
        [
            str(account_no),
            None,
            None,
            tx_id,
            tx_type,
            round(float(amount), 2),
            round(
                float(previous_balance),
                2,
            ),
            round(
                float(current_balance),
                2,
            ),
            dt,
        ]
    )

    wb.save(FILE_NAME)
    wb.close()

    return tx_id


def get_history(account_no):

    wb = openpyxl.load_workbook(
        FILE_NAME,
        read_only=True,
        data_only=True,
    )

    sheet = wb["Bank Records"]

    history = []

    for row in sheet.iter_rows(
        min_row=2,
        values_only=True,
    ):

        if (
            row[0] is not None
            and str(row[0]) == str(account_no)
        ):

            history.append(
                {
                    "Transaction ID":
                        str(row[3] or ""),

                    "Type":
                        str(row[4] or ""),

                    "Amount":
                        float(row[5] or 0),

                    "Previous Balance":
                        float(row[6] or 0),

                    "Current Balance":
                        float(row[7] or 0),

                    "Date-Time":
                        row[8],
                }
            )

    wb.close()

    return history


def initials(name):

    cleaned = (
        name or "?"
    ).strip()

    pieces = cleaned.split()

    if len(pieces) >= 2:

        return (
            pieces[0][0]
            + pieces[-1][0]
        ).upper()

    return (
        cleaned[:2].upper()
        if cleaned
        else "PY"
    )


# ============================================================
# SESSION STATE
# ============================================================

session_defaults = {
    "logged_in": False,
    "account_no": None,
    "name": None,
    "page": "Home",
    "ai_report": None,
    "chat_answer": None,
    "last_tx": None,
    "money_mode": "Deposit",
}

for key, value in session_defaults.items():

    if key not in st.session_state:
        st.session_state[key] = value


def logout():

    st.session_state.logged_in = False
    st.session_state.account_no = None
    st.session_state.name = None
    st.session_state.page = "Home"
    st.session_state.ai_report = None
    st.session_state.chat_answer = None
    st.session_state.last_tx = None


# ============================================================
# LIVE CLOCK
# ============================================================

@st.fragment(run_every="1s")
def live_clock():

    current = datetime.now()

    render_html(
        f"""
        <div style="
            text-align:right;
            color:#8fa4bf;
            font-size:.72rem;
            margin-bottom:10px;
        ">
            <span style="
                color:#56edb1;
                font-weight:800;
            ">
                ● LIVE
            </span>

            &nbsp;

            {safe(current.strftime("%A, %d %B %Y"))}

            &nbsp;•&nbsp;

            {safe(current.strftime("%H:%M:%S"))}
        </div>
        """
    )


# ============================================================
# AI HELPERS
# ============================================================

def build_ai_context(history, balance):

    lines = []

    for tx in history[-20:]:

        lines.append(
            f"- {tx['Date-Time']} | "
            f"{tx['Type']} | "
            f"₹{tx['Amount']:.2f} | "
            f"balance after "
            f"₹{tx['Current Balance']:.2f}"
        )

    recent_text = (
        "\n".join(lines)
        if lines
        else "No transactions yet."
    )

    return (
        f"Current balance: "
        f"₹{balance or 0:.2f}\n"
        f"Recent transactions:\n"
        f"{recent_text}"
    )


def generate_ai_report():

    if not gemini_ready:
        return None

    balance = (
        get_balance(
            st.session_state.account_no
        )
        or 0
    )

    history = get_history(
        st.session_state.account_no
    )

    context = build_ai_context(
        history,
        balance,
    )

    prompt = f"""
You are PY BANK AI, an educational banking dashboard assistant.

Analyze the following account activity and create a concise, attractive
financial activity report for the account owner.

ACCOUNT DATA
{context}

Use exactly this Markdown structure:

# 🤖 PY BANK AI Report

## 📌 Account Snapshot
Summarize the current balance and transaction count.

## 💸 Money Flow
Summarize deposits, withdrawals, and net movement.

## 🔎 Activity Patterns
Identify simple patterns that can be supported by the supplied transaction data.

## ✨ Positive Signals
Mention useful observations from the data.

## ⚠️ Things to Watch
Mention unusual or noteworthy activity only when supported by the supplied data.

## 🧭 Practical Next Steps
Give general, educational organization suggestions. Do not recommend
specific investments, loans, securities, or financial products.

## 📝 Important Note
State that the report is an AI-generated summary of the provided transaction
data and is not professional financial advice.

Rules:
- Use only supplied transaction data.
- Do not invent merchants, dates, income sources, or expenses.
- Do not claim real-time banking information.
- Use INR formatting.
- Keep it readable and visually friendly.
"""

    response = client.models.generate_content(
        model=GEMINI_MODEL,
        contents=prompt,
    )

    return response.text


def generate_ai_answer(question):

    if not gemini_ready:
        return None

    balance = (
        get_balance(
            st.session_state.account_no
        )
        or 0
    )

    history = get_history(
        st.session_state.account_no
    )

    context = build_ai_context(
        history,
        balance,
    )

    prompt = f"""
You are PY BANK AI, an educational assistant inside a personal banking
prototype.

Account context:
{context}

User question:
{question}

Answer helpfully using only information that can be supported by the
account context above.

You may explain transaction history, balances, totals, and basic budgeting
concepts in general terms.

Do not:
- invent transactions,
- expose or discuss the user's PIN,
- pretend to access external bank systems,
- provide investment/loan/security recommendations,
- make claims about live balances beyond the supplied data.

Keep the answer concise and use Markdown.
"""

    response = client.models.generate_content(
        model=GEMINI_MODEL,
        contents=prompt,
    )

    return response.text


# ============================================================
# TOP BRAND
# ============================================================

render_html(
    """
    <div class="brand">

        <div class="brand-icon">
            🏦
        </div>

        <div>

            <div class="brand-name">
                PY BANK
            </div>

            <div class="brand-sub">
                AI-powered personal banking prototype
            </div>

        </div>

    </div>
    """
)

live_clock()


# ============================================================
# LOGIN / LANDING PAGE
# ============================================================

if not st.session_state.logged_in:

    render_html(
        """
        <div class="hero">

            <div class="hero-row">

                <div>

                    <div class="hero-kicker">
                        Smart banking • Gemini assisted
                    </div>

                    <div class="hero-title">
                        Bank smarter with PY BANK.
                    </div>

                    <div class="hero-copy">
                        A polished Streamlit banking experience with your
                        original Excel-backed account system, live interface,
                        transaction analytics and optional Gemini assistance.
                    </div>

                </div>

                <div class="live-chip">

                    <span class="live-dot"></span>

                    Python • Streamlit • Gemini

                </div>

            </div>

        </div>
        """
    )

    left, right = st.columns(
        [1.15, .85],
        gap="large",
    )

    with left:

        login_tab, create_tab = st.tabs(
            [
                "🔐 Login",
                "✨ Create Account",
            ]
        )

        with login_tab:

            st.markdown(
                "### Welcome back 👋"
            )

            with st.form("login_form"):

                account_no = st.text_input(
                    "Account Number",
                    placeholder=(
                        "Enter your account number"
                    ),
                )

                pin = st.text_input(
                    "4-Digit PIN",
                    type="password",
                    max_chars=4,
                    placeholder="••••",
                )

                login = st.form_submit_button(
                    "🚀 Enter PY BANK",
                    use_container_width=True,
                    type="primary",
                )

            if login:

                account_no = account_no.strip()

                if not account_no or not pin:

                    st.error(
                        "Please enter both account "
                        "number and PIN."
                    )

                elif (
                    len(pin) != 4
                    or not pin.isdigit()
                ):

                    st.error(
                        "PIN must contain exactly 4 digits."
                    )

                else:

                    account = authenticate(
                        account_no,
                        pin,
                    )

                    if account:

                        st.session_state.logged_in = True

                        st.session_state.account_no = (
                            account["account_no"]
                        )

                        st.session_state.name = (
                            account["name"]
                        )

                        st.session_state.page = "Home"

                        st.toast(
                            "Login successful 👋",
                            icon="✅",
                        )

                        st.rerun()

                    else:

                        st.error(
                            "Login failed. "
                            "Check your credentials."
                        )

        with create_tab:

            st.markdown(
                "### Start your PY BANK journey ✨"
            )

            with st.form("create_form"):

                new_name = st.text_input(
                    "Full Name",
                    placeholder="e.g. Parth Sharma",
                )

                new_account = st.text_input(
                    "Choose Account Number",
                    placeholder="Numbers only",
                )

                new_pin = st.text_input(
                    "Create 4-Digit PIN",
                    type="password",
                    max_chars=4,
                    placeholder="••••",
                )

                opening = st.number_input(
                    "Opening Balance (₹)",
                    min_value=0.0,
                    step=100.0,
                    format="%.2f",
                )

                create = st.form_submit_button(
                    "✨ Create My Account",
                    use_container_width=True,
                    type="primary",
                )

            if create:

                new_name = new_name.strip()
                new_account = new_account.strip()

                if not new_name:

                    st.error(
                        "Please enter your name."
                    )

                elif not new_account.isdigit():

                    st.error(
                        "Account number must contain digits only."
                    )

                elif account_exists(new_account):

                    st.error(
                        "That account number already exists."
                    )

                elif (
                    len(new_pin) != 4
                    or not new_pin.isdigit()
                ):

                    st.error(
                        "PIN must contain exactly 4 digits."
                    )

                elif opening < 0:

                    st.error(
                        "Opening balance cannot be negative."
                    )

                else:

                    amount = round(
                        opening,
                        2,
                    )

                    txid = transaction_id()

                    wb = openpyxl.load_workbook(
                        FILE_NAME
                    )

                    sheet = wb[
                        "Bank Records"
                    ]

                    sheet.append(
                        [
                            new_account,
                            new_name,
                            new_pin,
                            txid,
                            "Opening",
                            amount,
                            0.00,
                            amount,
                            now_text(),
                        ]
                    )

                    wb.save(
                        FILE_NAME
                    )

                    wb.close()

                    st.success(
                        "🎉 Account created successfully!"
                    )

                    st.info(
                        f"Account Number: "
                        f"{new_account}"
                    )

                    st.info(
                        f"Opening Balance: "
                        f"₹{amount:,.2f}"
                    )

    with right:

        render_html(
            """
            <div class="section-title">
                ✨ PY BANK features
            </div>

            <div class="feature-card">

                <div class="feature-icon">
                    💎
                </div>

                <div class="feature-title">
                    Premium dashboard
                </div>

                <div class="feature-copy">
                    Decorative cards, live status indicators and a polished
                    fintech-style interface.
                </div>

            </div>

            <br>

            <div class="feature-card">

                <div class="feature-icon">
                    🤖
                </div>

                <div class="feature-title">
                    Gemini AI assistant
                </div>

                <div class="feature-copy">
                    Generate easy-to-read summaries of the transaction data
                    stored in your prototype account.
                </div>

            </div>

            <br>

            <div class="feature-card">

                <div class="feature-icon">
                    📊
                </div>

                <div class="feature-title">
                    Money insights
                </div>

                <div class="feature-copy">
                    View money-in, money-out, transaction counts and a balance
                    timeline from your workbook.
                </div>

            </div>

            <br>

            <div class="feature-card">

                <div class="feature-icon">
                    ⚡
                </div>

                <div class="feature-title">
                    Fast actions
                </div>

                <div class="feature-copy">
                    Deposit, withdraw, view history and download your activity
                    in a few clicks.
                </div>

            </div>
            """
        )

        if not gemini_ready:

            st.warning(
                "🔑 Gemini is optional, but the AI features "
                "are disabled until GEMINI_API_KEY is configured."
            )

        else:

            st.success(
                "🤖 Gemini AI is connected."
            )

    render_html(
        """
        <div class="footer">
            PY BANK • Built with Python + Streamlit + Gemini
        </div>
        """
    )

    st.stop()


# ============================================================
# SIDEBAR
# ============================================================

with st.sidebar:

    render_html(
        """
        <div class="brand">

            <div class="brand-icon">
                🏦
            </div>

            <div>

                <div class="brand-name">
                    PY BANK
                </div>

                <div class="brand-sub">
                    Personal banking console
                </div>

            </div>

        </div>
        """
    )

    sidebar_name = safe(
        st.session_state.name
    )

    sidebar_account = safe(
        st.session_state.account_no
    )

    sidebar_initials = safe(
        initials(st.session_state.name)
    )

    render_html(
        f"""
        <div class="glass" style="text-align:center;">

            <div style="
                width:64px;
                height:64px;
                border-radius:20px;
                margin:0 auto 12px;
                display:grid;
                place-items:center;
                font-size:24px;
                font-weight:800;
                background:
                    linear-gradient(
                        135deg,
                        rgba(93,228,255,.20),
                        rgba(159,124,255,.26)
                    );
                border:1px solid rgba(255,255,255,.11);
            ">
                {sidebar_initials}
            </div>

            <div style="font-weight:800;">
                {sidebar_name}
            </div>

            <div style="
                color:#91a6c1;
                font-size:.72rem;
                margin-top:4px;
            ">
                Account {sidebar_account}
            </div>

        </div>
        """
    )

    st.divider()

    st.subheader(
        "🧭 Navigation"
    )

    pages = {
        "🏠 Home": "Home",
        "💸 Deposit / Withdraw": "Money",
        "📜 Transactions": "History",
        "📊 Insights": "Insights",
        "🤖 PY BANK AI": "AI",
        "👤 Profile": "Profile",
    }

    current_index = list(
        pages.values()
    ).index(
        st.session_state.page
    )

    selected = st.radio(
        "Choose",
        list(pages.keys()),
        index=current_index,
        label_visibility="collapsed",
    )

    st.session_state.page = (
        pages[selected]
    )

    st.divider()

    side_balance = (
        get_balance(
            st.session_state.account_no
        )
        or 0
    )

    render_html(
        f"""
        <div class="metric-label">
            Available Balance
        </div>

        <div style="
            font-family:'Space Grotesk';
            font-size:1.55rem;
            font-weight:700;
        ">
            ₹{side_balance:,.2f}
        </div>

        <div style="
            color:#56edb1;
            font-size:.68rem;
            margin-top:4px;
        ">
            ● Live account
        </div>
        """
    )

    st.divider()

    if gemini_ready:
        st.caption(
            "🤖 Gemini AI: Connected"
        )
    else:
        st.caption(
            "🤖 Gemini AI: Not configured"
        )

    st.caption(
        "Prototype database: "
        "bank_records_1.xlsx"
    )

    if st.button(
        "🚪 Logout",
        use_container_width=True,
    ):

        logout()

        st.rerun()


# ============================================================
# LOGGED-IN HERO
# ============================================================

name_html = safe(
    st.session_state.name
)

render_html(
    f"""
    <div class="hero">

        <div class="hero-row">

            <div>

                <div class="hero-kicker">
                    Personal banking dashboard
                </div>

                <div class="hero-title">
                    Good to see you,
                    {name_html}.
                </div>

                <div class="hero-copy">
                    Monitor your account, move money, explore activity
                    and ask PY BANK AI to explain your transaction data.
                </div>

            </div>

            <div class="live-chip">

                <span class="live-dot"></span>

                Live session

            </div>

        </div>

    </div>
    """
)


# ============================================================
# HOME
# ============================================================

if st.session_state.page == "Home":

    balance = (
        get_balance(
            st.session_state.account_no
        )
        or 0
    )

    history = get_history(
        st.session_state.account_no
    )

    left, right = st.columns(
        [1.35, .65],
        gap="large",
    )

    with left:

        account_html = safe(
            st.session_state.account_no
        )

        render_html(
            f"""
            <div class="balance-card">

                <div class="balance-label">
                    Total available balance
                </div>

                <div class="balance-value">
                    ₹{balance:,.2f}
                </div>

                <span class="account-chip">
                    A/C {account_html}
                </span>

                <span class="account-chip"
                      style="margin-left:6px;">
                    ● Active
                </span>

            </div>
            """
        )

    with right:

        render_html(
            """
            <div class="ai-card">

                <div class="ai-badge">
                    🤖 GEMINI ENABLED
                </div>

                <h3 style="
                    margin:10px 0 6px;
                ">
                    Your AI banking companion
                </h3>

                <div style="
                    color:#91a6c1;
                    font-size:.78rem;
                    line-height:1.55;
                ">
                    Turn your transaction history into a clear
                    AI-generated activity summary.
                </div>

            </div>
            """
        )

    st.markdown(
        '<div class="section-title">📊 Account Overview</div>',
        unsafe_allow_html=True,
    )

    deposits = sum(
        x["Amount"]
        for x in history
        if x["Type"] == "Deposit"
    )

    withdrawals = sum(
        x["Amount"]
        for x in history
        if x["Type"] == "Withdrawal"
    )

    net_flow = (
        deposits - withdrawals
    )

    overview = st.columns(4)

    with overview[0]:

        render_html(
            f"""
            <div class="metric-card">

                <div class="metric-label">
                    Transactions
                </div>

                <div class="metric-number">
                    {len(history)}
                </div>

                <div class="pill pill-cyan">
                    All activity
                </div>

            </div>
            """
        )

    with overview[1]:

        render_html(
            f"""
            <div class="metric-card">

                <div class="metric-label">
                    Money In
                </div>

                <div class="metric-number">
                    ₹{deposits:,.0f}
                </div>

                <div class="pill pill-green">
                    ↑ Deposits
                </div>

            </div>
            """
        )

    with overview[2]:

        render_html(
            f"""
            <div class="metric-card">

                <div class="metric-label">
                    Money Out
                </div>

                <div class="metric-number">
                    ₹{withdrawals:,.0f}
                </div>

                <div class="pill pill-red">
                    ↓ Withdrawals
                </div>

            </div>
            """
        )

    with overview[3]:

        flow_class = (
            "pill-green"
            if net_flow >= 0
            else "pill-red"
        )

        flow_icon = (
            "↑"
            if net_flow >= 0
            else "↓"
        )

        render_html(
            f"""
            <div class="metric-card">

                <div class="metric-label">
                    Net Flow
                </div>

                <div class="metric-number">
                    ₹{net_flow:,.0f}
                </div>

                <div class="pill {flow_class}">
                    {flow_icon} Since opening
                </div>

            </div>
            """
        )

    st.markdown(
        '<div class="section-title">⚡ Quick Services</div>',
        unsafe_allow_html=True,
    )

    q1, q2, q3, q4 = st.columns(4)

    with q1:

        if st.button(
            "💰 Deposit",
            use_container_width=True,
        ):

            st.session_state.page = "Money"
            st.session_state.money_mode = "Deposit"

            st.rerun()

    with q2:

        if st.button(
            "💸 Withdraw",
            use_container_width=True,
        ):

            st.session_state.page = "Money"
            st.session_state.money_mode = "Withdraw"

            st.rerun()

    with q3:

        if st.button(
            "📜 History",
            use_container_width=True,
        ):

            st.session_state.page = "History"

            st.rerun()

    with q4:

        if st.button(
            "🤖 AI Report",
            use_container_width=True,
        ):

            st.session_state.page = "AI"

            st.rerun()

    st.markdown(
        '<div class="section-title">🕘 Recent Activity</div>',
        unsafe_allow_html=True,
    )

    recent = list(
        reversed(
            history[-5:]
        )
    )

    if not recent:

        st.info(
            "Your recent activity will appear "
            "here after your first transaction."
        )

    else:

        for tx in recent:

            incoming = tx["Type"] in {
                "Opening",
                "Deposit",
            }

            prefix = (
                "+"
                if incoming
                else "-"
            )

            cls = (
                "money-in"
                if incoming
                else "money-out"
            )

            icon = (
                "↗"
                if incoming
                else "↘"
            )

            render_html(
                f"""
                <div class="transaction-card">

                    <div class="transaction-row">

                        <div>

                            <div class="transaction-name">
                                {icon}
                                {safe(tx["Type"])}
                            </div>

                            <div class="transaction-meta">
                                {safe(tx["Date-Time"])}
                                • ID
                                {safe(tx["Transaction ID"])}
                            </div>

                        </div>

                        <div class="transaction-money {cls}">
                            {prefix}
                            ₹{tx["Amount"]:,.2f}
                        </div>

                    </div>

                </div>
                """
            )


# ============================================================
# MONEY
# ============================================================

elif st.session_state.page == "Money":

    st.markdown(
        '<div class="section-title">💸 Move Your Money</div>',
        unsafe_allow_html=True,
    )

    mode = st.session_state.get(
        "money_mode",
        "Deposit",
    )

    action = st.radio(
        "Choose operation",
        ["Deposit", "Withdraw"],
        index=(
            0
            if mode == "Deposit"
            else 1
        ),
        horizontal=True,
    )

    st.session_state.money_mode = action

    current_balance = (
        get_balance(
            st.session_state.account_no
        )
        or 0
    )

    left, right = st.columns(
        [1.05, .95],
        gap="large",
    )

    with left:

        with st.form("money_form"):

            amount = st.number_input(
                "Amount (₹)",
                min_value=0.01,
                step=100.0,
                format="%.2f",
            )

            note = st.text_input(
                "Reference / Note",
                placeholder="Optional",
            )

            submit = st.form_submit_button(
                (
                    "💰 Deposit Funds"
                    if action == "Deposit"
                    else "💸 Withdraw Funds"
                ),
                use_container_width=True,
                type="primary",
            )

        if submit:

            if amount <= 0:

                st.error(
                    "Amount must be greater than ₹0."
                )

            elif (
                action == "Withdraw"
                and amount > current_balance
            ):

                st.error(
                    "Insufficient balance."
                )

                st.warning(
                    f"Available balance: "
                    f"₹{current_balance:,.2f}"
                )

            else:

                new_balance = (
                    round(
                        current_balance + amount,
                        2,
                    )
                    if action == "Deposit"
                    else round(
                        current_balance - amount,
                        2,
                    )
                )

                txid = add_transaction(
                    st.session_state.account_no,
                    st.session_state.name,
                    action,
                    amount,
                    current_balance,
                    new_balance,
                )

                st.session_state.last_tx = {
                    "type": action,
                    "amount": amount,
                    "new_balance": new_balance,
                    "txid": txid,
                    "note": note,
                    "time": now_text(),
                }

                st.toast(
                    f"{action} successful • "
                    f"₹{amount:,.2f}",
                    icon="✅",
                )

                st.rerun()

    with right:

        account_html = safe(
            st.session_state.account_no
        )

        render_html(
            f"""
            <div class="balance-card">

                <div class="balance-label">
                    Current balance
                </div>

                <div class="balance-value">
                    ₹{current_balance:,.2f}
                </div>

                <span class="account-chip">
                    A/C {account_html}
                </span>

            </div>
            """
        )

        latest = st.session_state.get(
            "last_tx"
        )

        if latest:

            render_html(
                f"""
                <div class="ai-card"
                     style="margin-top:14px;">

                    <div class="ai-badge">
                        ✅ TRANSACTION COMPLETE
                    </div>

                    <h4 style="
                        margin:10px 0 3px;
                    ">
                        {safe(latest["type"])}
                        •
                        ₹{latest["amount"]:,.2f}
                    </h4>

                    <div style="
                        color:#91a6c1;
                        font-size:.72rem;
                    ">
                        {safe(latest["time"])}
                        • ID
                        {safe(latest["txid"])}
                    </div>

                    <div style="
                        margin-top:11px;
                        color:#56edb1;
                        font-weight:800;
                    ">
                        New balance:
                        ₹{latest["new_balance"]:,.2f}
                    </div>

                </div>
                """
            )


# ============================================================
# HISTORY
# ============================================================

elif st.session_state.page == "History":

    st.markdown(
        '<div class="section-title">📜 Transaction History</div>',
        unsafe_allow_html=True,
    )

    history = list(
        reversed(
            get_history(
                st.session_state.account_no
            )
        )
    )

    if not history:

        st.info(
            "No transactions found."
        )

    else:

        top1, top2 = st.columns(
            [.55, .45]
        )

        with top1:

            filter_type = st.selectbox(
                "Transaction type",
                [
                    "All",
                    "Opening",
                    "Deposit",
                    "Withdrawal",
                ],
            )

        with top2:

            search_text = st.text_input(
                "Search",
                placeholder=(
                    "Transaction ID or date..."
                ),
            )

        if filter_type != "All":

            history = [
                x
                for x in history
                if x["Type"] == filter_type
            ]

        if search_text.strip():

            term = (
                search_text
                .strip()
                .lower()
            )

            history = [
                x
                for x in history
                if (
                    term
                    in x["Transaction ID"]
                    .lower()
                )
                or (
                    term
                    in str(
                        x["Date-Time"]
                    ).lower()
                )
            ]

        st.caption(
            f"{len(history)} transaction(s) shown"
        )

        for tx in history:

            incoming = tx["Type"] in {
                "Opening",
                "Deposit",
            }

            prefix = (
                "+"
                if incoming
                else "-"
            )

            cls = (
                "money-in"
                if incoming
                else "money-out"
            )

            render_html(
                f"""
                <div class="transaction-card">

                    <div class="transaction-row">

                        <div>

                            <div class="transaction-name">
                                {safe(tx["Type"])}
                            </div>

                            <div class="transaction-meta">
                                {safe(tx["Date-Time"])}
                                •
                                {safe(tx["Transaction ID"])}
                            </div>

                            <div class="transaction-meta">
                                Balance after:
                                ₹{tx["Current Balance"]:,.2f}
                            </div>

                        </div>

                        <div class="transaction-money {cls}">
                            {prefix}
                            ₹{tx["Amount"]:,.2f}
                        </div>

                    </div>

                </div>
                """
            )

        export_lines = [
            "# PY BANK Transaction Statement",
            "",
            f"Account: {st.session_state.account_no}",
            f"Name: {st.session_state.name}",
            f"Generated: {now_text()}",
            "",
        ]

        for tx in history:

            export_lines.append(
                f"- {tx['Date-Time']} | "
                f"{tx['Type']} | "
                f"₹{tx['Amount']:.2f} | "
                f"Balance "
                f"₹{tx['Current Balance']:.2f} | "
                f"ID {tx['Transaction ID']}"
            )

        st.download_button(
            "📥 Download Statement",
            data="\n".join(
                export_lines
            ),
            file_name=(
                f"PY_BANK_"
                f"{st.session_state.account_no}"
                f"_statement.md"
            ),
            mime="text/markdown",
            width="stretch",
        )


# ============================================================
# INSIGHTS
# ============================================================

elif st.session_state.page == "Insights":

    st.markdown(
        '<div class="section-title">📊 Banking Insights</div>',
        unsafe_allow_html=True,
    )

    history = get_history(
        st.session_state.account_no
    )

    deposits = [
        x["Amount"]
        for x in history
        if x["Type"] == "Deposit"
    ]

    withdrawals = [
        x["Amount"]
        for x in history
        if x["Type"] == "Withdrawal"
    ]

    c1, c2, c3, c4 = st.columns(4)

    with c1:

        st.metric(
            "Total Deposits",
            f"₹{sum(deposits):,.2f}",
        )

    with c2:

        st.metric(
            "Total Withdrawals",
            f"₹{sum(withdrawals):,.2f}",
        )

    with c3:

        st.metric(
            "Deposit Count",
            len(deposits),
        )

    with c4:

        st.metric(
            "Withdrawal Count",
            len(withdrawals),
        )

    if history:

        chart_col, flow_col = st.columns(
            [1.25, .75],
            gap="large",
        )

        with chart_col:

            st.markdown(
                "#### 📈 Balance timeline"
            )

            chart_data = {}

            for item in history:

                try:

                    dt = datetime.strptime(
                        str(
                            item["Date-Time"]
                        ),
                        "%d-%m-%Y %H:%M:%S",
                    )

                    chart_data[dt] = (
                        item["Current Balance"]
                    )

                except (
                    TypeError,
                    ValueError,
                ):

                    pass

            if chart_data:

                st.line_chart(
                    chart_data
                )

        with flow_col:

            st.markdown(
                "#### 💰 Money flow"
            )

            st.bar_chart(
                {
                    "Money In":
                        sum(deposits),

                    "Money Out":
                        sum(withdrawals),
                }
            )

            render_html(
                f"""
                <div class="glass">

                    <div class="metric-label">
                        Transaction mix
                    </div>

                    <div style="
                        margin-top:8px;
                        line-height:1.8;
                    ">

                        <b>{len(deposits)}</b>
                        deposits

                        <br>

                        <b>{len(withdrawals)}</b>
                        withdrawals

                        <br>

                        <b>{len(history)}</b>
                        total records

                    </div>

                </div>
                """
            )

    else:

        st.info(
            "Create or use an account "
            "to unlock insights."
        )


# ============================================================
# AI
# ============================================================

elif st.session_state.page == "AI":

    st.markdown(
        '<div class="section-title">🤖 PY BANK AI</div>',
        unsafe_allow_html=True,
    )

    if not gemini_ready:

        st.warning(
            "Gemini is not configured. Add GEMINI_API_KEY "
            "to `.env` locally or Streamlit Secrets when deploying."
        )

        st.stop()

    render_html(
        """
        <div class="ai-card">

            <div class="ai-badge">
                ✦ POWERED BY GOOGLE GEMINI
            </div>

            <h2 style="
                margin:10px 0 6px;
            ">
                Meet your banking assistant.
            </h2>

            <div style="
                color:#91a6c1;
                line-height:1.65;
                font-size:.87rem;
            ">
                Ask questions about your transaction history
                or generate a visual, easy-to-read activity
                report from your PY BANK account.
            </div>

        </div>
        """
    )

    st.markdown(
        "#### ✨ AI Report"
    )

    if st.button(
        "🪄 Analyze My Banking Activity",
        use_container_width=True,
        type="primary",
    ):

        with st.status(
            "🤖 PY BANK AI is analyzing your account...",
            expanded=True,
        ) as status:

            st.write(
                "🔎 Reading recent transaction activity..."
            )

            time.sleep(.35)

            st.write(
                "📊 Calculating money flow..."
            )

            time.sleep(.35)

            st.write(
                "🧠 Sending the relevant account context to Gemini..."
            )

            time.sleep(.35)

            try:

                report = (
                    generate_ai_report()
                )

                st.session_state.ai_report = report

                status.update(
                    label="✅ AI report ready!",
                    state="complete",
                    expanded=False,
                )

            except Exception as error:

                status.update(
                    label="❌ Gemini request failed",
                    state="error",
                    expanded=True,
                )

                st.error(
                    str(error)
                )

    if st.session_state.ai_report:

        st.divider()

        st.markdown(
            st.session_state.ai_report
        )

        st.download_button(
            "📄 Download AI Report",
            data=st.session_state.ai_report,
            file_name=(
                f"PY_BANK_AI_Report_"
                f"{st.session_state.account_no}.md"
            ),
            mime="text/markdown",
            width="stretch",
        )

    st.divider()

    st.markdown(
        "#### 💬 Ask PY BANK AI"
    )

    question = st.text_input(
        "Your question",
        placeholder=(
            "e.g. How much have I deposited so far?"
        ),
    )

    if st.button(
        "Ask Gemini →",
        use_container_width=True,
        type="primary",
    ):

        if not question.strip():

            st.warning(
                "Type a question first."
            )

        else:

            with st.status(
                "🤖 PY BANK AI is thinking...",
                expanded=True,
            ) as status:

                st.write(
                    "📚 Preparing your account context..."
                )

                time.sleep(.25)

                try:

                    answer = (
                        generate_ai_answer(
                            question.strip()
                        )
                    )

                    st.session_state.chat_answer = answer

                    status.update(
                        label="✅ Answer ready!",
                        state="complete",
                        expanded=False,
                    )

                except Exception as error:

                    status.update(
                        label="❌ Gemini request failed",
                        state="error",
                        expanded=True,
                    )

                    st.error(
                        str(error)
                    )

    if st.session_state.chat_answer:

        render_html(
            """
            <div class="glass">

                <div class="ai-badge">
                    🤖 GEMINI RESPONSE
                </div>

            </div>
            """
        )

        st.markdown(
            st.session_state.chat_answer
        )


# ============================================================
# PROFILE
# ============================================================

elif st.session_state.page == "Profile":

    account = get_account(
        st.session_state.account_no
    )

    history = get_history(
        st.session_state.account_no
    )

    balance = (
        get_balance(
            st.session_state.account_no
        )
        or 0
    )

    st.markdown(
        '<div class="section-title">👤 Profile</div>',
        unsafe_allow_html=True,
    )

    left, right = st.columns(
        [.85, 1.15],
        gap="large",
    )

    with left:

        profile_initials = safe(
            initials(
                st.session_state.name
            )
        )

        profile_name = safe(
            st.session_state.name
        )

        profile_account = safe(
            st.session_state.account_no
        )

        render_html(
            f"""
            <div class="balance-card"
                 style="text-align:center;">

                <div style="
                    width:76px;
                    height:76px;
                    border-radius:24px;
                    margin:0 auto 15px;
                    display:grid;
                    place-items:center;
                    font-size:26px;
                    font-weight:800;
                    background:
                        linear-gradient(
                            135deg,
                            rgba(93,228,255,.20),
                            rgba(159,124,255,.25)
                        );
                    border:
                        1px solid rgba(255,255,255,.11);
                ">
                    {profile_initials}
                </div>

                <div style="
                    font-size:1.35rem;
                    font-weight:800;
                ">
                    {profile_name}
                </div>

                <div style="
                    color:#91a6c1;
                    font-size:.76rem;
                    margin-top:5px;
                ">
                    PY BANK customer
                </div>

                <div style="margin-top:14px;">

                    <span class="account-chip">
                        A/C {profile_account}
                    </span>

                </div>

            </div>
            """
        )

    with right:

        customer = (
            account["name"]
            if account
            else st.session_state.name
        )

        render_html(
            f"""
            <div class="glass">

                <div class="section-title"
                     style="margin-top:0;">
                    Account Details
                </div>

                <div style="
                    line-height:2;
                    color:#b7c5d8;
                ">

                    <b>Customer:</b>
                    {safe(customer)}

                    <br>

                    <b>Account Number:</b>
                    {safe(st.session_state.account_no)}

                    <br>

                    <b>Current Balance:</b>
                    ₹{balance:,.2f}

                    <br>

                    <b>Total Transactions:</b>
                    {len(history)}

                    <br>

                    <b>Status:</b>

                    <span style="
                        color:#56edb1;
                    ">
                        ● Active
                    </span>

                </div>

            </div>
            """
        )

        render_html(
            """
            <div class="glass"
                 style="margin-top:14px;">

                <div class="section-title"
                     style="margin-top:0;">
                    🔒 Prototype security
                </div>

                <div style="
                    color:#91a6c1;
                    font-size:.8rem;
                    line-height:1.65;
                ">
                    This app preserves the Excel/PIN design
                    for learning and demonstration. For production
                    use, credentials should be hashed and sensitive
                    banking data should live in a transactional database.
                </div>

            </div>
            """
        )


# ============================================================
# FOOTER
# ============================================================

render_html(
    """
    <div class="footer">
        🏦 PY BANK • AI Banking Prototype •
        Powered by Streamlit + Google Gemini
    </div>
    """
)