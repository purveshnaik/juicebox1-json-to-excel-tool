import streamlit as st
import json
import tempfile
import pandas as pd
import re
import ast
import streamlit.components.v1 as components
from process_juicebox import extract, make_xlsx

st.set_page_config(page_title="JSON to Excel", layout="wide")

st.title("JSON → Spreadsheet Converter")

# --- CLEANING FUNCTION ---
def clean_json(text):
    text = text.strip()

    # Remove BOM / weird chars
    text = text.encode('utf-8', 'ignore').decode('utf-8')

    return text


# --- ROBUST PARSER ---
def parse_json_safely(text):
    # Attempt 1: Direct JSON
    try:
        return json.loads(text)
    except:
        pass

    # Attempt 2: Fix raw newlines
    try:
        fixed = re.sub(r'(?<!\\)\n', r'\\n', text)
        return json.loads(fixed)
    except:
        pass

    # Attempt 3: Handle stringified JSON
    try:
        unescaped = text.encode().decode('unicode_escape')
        return json.loads(unescaped)
    except:
        pass

    # Attempt 4: Python dict fallback (very useful)
    try:
        return ast.literal_eval(text)
    except:
        pass

    raise ValueError("Unable to parse JSON. Input is too malformed.")


# --- INPUT ---
json_text = st.text_area("Paste your JSON here", height=300)

if st.button("Convert"):
    if not json_text.strip():
        st.error("Paste some JSON first.")
    else:
        try:
            cleaned = clean_json(json_text)
            data = parse_json_safely(cleaned)

            if not isinstance(data, (list, dict)):
                st.error("Parsed data is not valid JSON structure.")
                st.stop()

            rows = extract(data)

            if not rows:
                st.warning("No data extracted.")
            else:
                df = pd.DataFrame(rows)

                st.success(f"Processed {len(df)} rows ✅")

                # --- PREVIEW ---
                st.dataframe(df, use_container_width=True)

                # --- DOWNLOAD EXCEL ---
                with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
                    make_xlsx(rows, tmp.name)
                    tmp_path = tmp.name

                with open(tmp_path, "rb") as f:
                    st.download_button(
                        label="Download Excel",
                        data=f,
                        file_name="output.xlsx"
                    )

                # --- COPY BUTTON ---
                tsv_data = df.to_csv(sep="\t", index=False)

                st.markdown("### Copy for Google Sheets")

                components.html(f"""
<textarea id="data" style="width:100%;height:200px;">{tsv_data}</textarea>
<br><br>
<button onclick="copyText()" style="padding:10px 20px;font-size:16px;cursor:pointer;">
Copy Entire Sheet
</button>

<script>
function copyText() {{
    var copyText = document.getElementById("data");
    copyText.select();
    copyText.setSelectionRange(0, 999999);
    document.execCommand("copy");
    alert("Copied to clipboard!");
}}
</script>
""", height=300)

        except Exception as e:
            st.error(f"❌ Failed to process JSON: {e}")

            st.markdown("### Debug Preview (first 500 chars)")
            st.code(json_text[:500])
