import streamlit as st
import json
import re
import tempfile
import pandas as pd
import streamlit.components.v1 as components
from process_juicebox import extract, make_xlsx

st.set_page_config(page_title="HeroScouter – JSON → Sheet", layout="wide")
st.title("HeroScouter · JSON → Spreadsheet")
st.caption("Paste Juicebox JSON export → see a live table → copy to Google Sheets in one click.")


# ── CORE FIX: escape control chars inside JSON strings ───────────────────────
def fix_control_chars_in_strings(text: str) -> str:
    """
    Walk character-by-character. When inside a JSON string value, replace
    literal control characters (newlines, tabs, carriage returns, etc.) with
    their proper JSON escape sequences.

    This is the root cause of the "Invalid control character" error when
    pasting JSON from browser DevTools / network inspector — the server
    embeds raw \\n characters in summary/description fields without escaping.

    Regex cannot safely do this because it can't track 'inside vs outside a string'.
    """
    result = []
    in_string = False
    i = 0
    n = len(text)

    while i < n:
        c = text[i]

        if in_string:
            if c == '\\':
                # Pass both chars of the escape sequence through unchanged
                result.append(c)
                i += 1
                if i < n:
                    result.append(text[i])
                    i += 1
                continue
            elif c == '"':
                in_string = False
                result.append(c)
                i += 1
                continue
            else:
                # Inside a string — sanitise bare control characters
                if   c == '\n': result.append('\\n')
                elif c == '\r': result.append('\\r')
                elif c == '\t': result.append('\\t')
                elif ord(c) < 0x20:
                    result.append(f'\\u{ord(c):04x}')
                else:
                    result.append(c)
                i += 1
                continue
        else:
            if c == '"':
                in_string = True
            result.append(c)
            i += 1

    return ''.join(result)


# ── FULL JSON NORMALISER ──────────────────────────────────────────────────────
def normalize_json(text: str) -> str:
    """Clean all common JSON paste issues from network-exported data."""

    # 1. Strip BOM and surrounding whitespace
    text = text.strip().lstrip('\ufeff')

    # 2. Python literals → JSON (do BEFORE string scanner so booleans outside strings are fixed)
    text = re.sub(r'\bTrue\b',  'true',  text)
    text = re.sub(r'\bFalse\b', 'false', text)
    text = re.sub(r'\bNone\b',  'null',  text)

    # 3. Remove JS-style comments (from DevTools copy-as-fetch)
    text = re.sub(r'//[^\n"]*\n', '\n', text)
    text = re.sub(r'/\*.*?\*/', '', text, flags=re.DOTALL)

    # 4. THE MAIN FIX: escape literal control characters inside string values
    text = fix_control_chars_in_strings(text)

    # 5. Trailing commas before ] or }
    text = re.sub(r',(\s*[}\]])', r'\1', text)

    # 6. Single-quoted documents → double-quoted (only when no " present at all)
    if "'" in text and '"' not in text:
        text = text.replace("'", '"')

    return text


def safe_parse(text: str):
    """Try to parse JSON; return (data, error_msg)."""
    try:
        return json.loads(text), None
    except json.JSONDecodeError as e:
        # Give a pinpointed error message
        lines = text.splitlines()
        lineno = e.lineno
        col = e.colno
        snippet = lines[lineno - 1] if lineno <= len(lines) else ""
        pointer = " " * (col - 1) + "^"
        msg = (
            f"**JSON parse error** on line {lineno}, col {col}:\n\n"
            f"```\n{snippet}\n{pointer}\n```\n\n"
            f"*{e.msg}*"
        )
        return None, msg


# ── COPY-TO-CLIPBOARD COMPONENT ──────────────────────────────────────────────
def clipboard_button(tsv: str):
    """Render a button that copies TSV to the clipboard (works in modern browsers)."""
    # Escape backticks and backslashes so they survive the JS template literal
    safe = tsv.replace("\\", "\\\\").replace("`", "\\`").replace("$", "\\$")
    components.html(
        f"""
        <button id="copyBtn"
            style="margin-top:10px;padding:10px 24px;font-size:15px;
                   background:#1F4E79;color:white;border:none;
                   border-radius:6px;cursor:pointer;">
            📋 Copy all rows to clipboard (paste into Google Sheets)
        </button>
        <span id="msg" style="margin-left:14px;font-size:14px;color:green;"></span>
        <script>
        document.getElementById("copyBtn").addEventListener("click", function() {{
            const text = `{safe}`;
            navigator.clipboard.writeText(text).then(function() {{
                document.getElementById("msg").innerText = "✅ Copied! Open Google Sheets → Ctrl+V (or ⌘V)";
            }}, function() {{
                // Fallback for older browsers / HTTP
                const ta = document.createElement("textarea");
                ta.value = text;
                ta.style.position = "fixed";
                ta.style.opacity = "0";
                document.body.appendChild(ta);
                ta.focus(); ta.select();
                document.execCommand("copy");
                document.body.removeChild(ta);
                document.getElementById("msg").innerText = "✅ Copied! Open Google Sheets → Ctrl+V (or ⌘V)";
            }});
        }});
        </script>
        """,
        height=70,
    )


# ── UI ────────────────────────────────────────────────────────────────────────
json_text = st.text_area(
    "Paste your Juicebox JSON export here",
    height=280,
    placeholder='{ "contacts": [ ... ] }  or a raw list  [ { ... }, ... ]',
)

col1, col2 = st.columns([1, 5])
convert_btn = col1.button("▶ Convert", type="primary", use_container_width=True)

if convert_btn:
    if not json_text.strip():
        st.error("Nothing to convert — paste your JSON first.")
        st.stop()

    # 1. Clean
    cleaned = normalize_json(json_text)

    # 2. Parse
    data, err = safe_parse(cleaned)
    if err:
        st.error("Could not parse JSON. Details below:")
        st.markdown(err)

        with st.expander("🔍 Show cleaned text (first 800 chars)"):
            st.code(cleaned[:800])

        st.info(
            "**Common causes from network/DevTools exports:**\n"
            "- Literal newlines inside text fields (auto-fixed — "
            "if you still see this, the JSON may be truncated)\n"
            "- Make sure you copied the *entire* response "
            "(scroll to the bottom and verify it ends with `}` or `]`)\n"
            "- Python booleans `True`/`False` are auto-fixed\n"
            "- Trailing commas are auto-fixed"
        )
        st.stop()

    # 3. Extract rows
    try:
        rows = extract(data)
    except Exception as e:
        st.error(f"Extraction failed: {e}")
        st.stop()

    if not rows:
        st.warning("JSON parsed OK but no candidate rows were found. Check the structure.")
        st.stop()

    # 4. Build DataFrame
    df = pd.DataFrame(rows)

    st.success(f"✅ {len(df)} candidates processed")

    # 5. Preview table
    st.subheader("Preview")
    st.dataframe(df, use_container_width=True, height=420)

    # 6. Copy-to-clipboard (TSV → Google Sheets)
    st.subheader("Copy to Google Sheets")
    tsv = df.to_csv(sep="\t", index=False)
    clipboard_button(tsv)

    # 7. Optional Excel download (kept but de-emphasised)
    with st.expander("⬇ Download as .xlsx instead"):
        with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
            make_xlsx(rows, tmp.name)
            tmp_path = tmp.name
        with open(tmp_path, "rb") as f:
            st.download_button(
                label="Download Excel file",
                data=f,
                file_name="heroscouter_candidates.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
