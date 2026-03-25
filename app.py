import streamlit as st
import json
import tempfile
import pandas as pd
from process_juicebox import extract, make_xlsx

st.set_page_config(page_title="JSON to Excel", layout="wide")

st.title("JSON → Spreadsheet Converter")

# --- INPUT ---
json_text = st.text_area("Paste your JSON here", height=300)

if st.button("Convert"):
    if not json_text.strip():
        st.error("Paste some JSON first.")
    else:
        try:
            data = json.loads(json_text)
            rows = extract(data)

            if not rows:
                st.warning("No data extracted.")
            else:
                df = pd.DataFrame(rows)

                st.success(f"Processed {len(df)} rows ✅")

                # --- PREVIEW TABLE ---
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

                # --- COPY OPTION (TSV for clean paste into Sheets) ---
                tsv_data = df.to_csv(sep="\t", index=False)

                st.markdown("### Copy for Google Sheets")
                st.text_area(
                    "Copy this and paste into Google Sheets",
                    tsv_data,
                    height=200
                )

        except Exception as e:
            st.error(f"Invalid JSON or processing error: {e}")