import streamlit as st
import json
import tempfile
import pandas as pd
import re
import streamlit.components.v1 as components
from process_juicebox import extract, make_xlsx

st.set_page_config(page_title="JSON to Excel", layout="wide")

st.title("JSON → Spreadsheet Converter")


# --- STRONG CLEANER ---
def normalize_json(text):
    text = text.strip()

    # Remove BOM / invisible chars
    text = text.encode('utf-8', 'ignore').decode('utf-8')

    # Fix unescaped newlines
    text = re.sub(r'(?<!\\)\n', r'\\n', text)

    # Remove trailing commas (VERY COMMON ISSUE)
    text = re.sub(r',(\s*[}\]])', r'\1', text)

    return text


# --- PARSER ---
def safe_parse(text):
    try:
        return json.loads(text)
    except Exception as e:
        raise ValueError(f"Still invalid JSON: {e}")


# --- INPUT ---
json_text = st.text_area("Paste your JSON here", height=300)

if st.button("Convert"):
    if not json_text.strip():
        st.error("Paste some JSON first.")
    else:
        try:
            cleaned = normalize_json(json_text)

            data = safe_parse(cleaned)

            if not isinstance(data, (list, dict)):
                st.error("Invalid JSON structure.")
                st.stop()

            rows = extract(data)

            if not rows:
                st.warning("No data extracted.")
            else:
                df = pd.DataFrame(rows)

                st.success(f"Processed {len(df)} rows ✅")

                # Preview
                st.dataframe(df, use_container_width=True)

                # Download Excel
                with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
                    make_xlsx(rows, tmp.name)
                    tmp_path = tmp.name

                with open(tmp_path, "rb") as f:
                    st.download_button(
                        label="Download Excel",
                        data=f,
                        file_name="output.xlsx"
                    )

                # Copy button
                tsv_data = df.to_csv(sep="\t", index=False)

                st.markdown("### Copy for Google Sheets")

                components.html(f"""
<textarea id="data" style="width:100%;height:200px;">{tsv_data}</textarea>
<br><br>
<button onclick="copyText()" style="padding:10px 20px;font-size:16px;">
Copy Entire Sheet
</button>

<script>
function copyText() {{
    var copyText = document.getElementById("data");
    copyText.select();
    document.execCommand("copy");
    alert("Copied!");
}}
</script>
""", height=300)

        except Exception as e:
            st.error(f"❌ Failed: {e}")

            st.markdown("### Debug (first 500 chars)")
            st.code(json_text[:500])
