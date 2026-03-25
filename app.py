import streamlit as st
import json
import tempfile
import pandas as pd
import re
import streamlit.components.v1 as components
from process_juicebox import extract, make_xlsx

st.set_page_config(page_title="JSON to Excel", layout="wide")

st.title("JSON → Spreadsheet Converter")

# --- CLEANER FUNCTION (fix invalid JSON newlines) ---
def clean_json(text):
    import re

    # Remove leading/trailing whitespace
    text = text.strip()

    # Remove invisible BOM / weird chars
    text = text.encode('utf-8', 'ignore').decode('utf-8')

    # Fix unescaped newlines
    text = re.sub(r'(?<!\\)\n', r'\\n', text)

    return text

# --- INPUT ---
json_text = st.text_area("Paste your JSON here", height=300)

if st.button("Convert"):
    if not json_text.strip():
        st.error("Paste some JSON first.")
    else:
        try:
            # 🔥 CLEAN JSON BEFORE PARSING
            cleaned_text = clean_json(json_text)

            try:
                data = json.loads(cleaned_text)
            except json.JSONDecodeError as e:
                st.error(f"JSON Error: {e}")
                st.text_area("Problematic JSON (first 500 chars)", cleaned_text[:500])
                st.stop()

            if not isinstance(data, (list, dict)):
                st.error("Invalid JSON format.")
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
            st.error(f"Error: {e}")
