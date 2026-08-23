import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
from PIL import Image
import easyocr

st.set_page_config(page_title="Image to Excel Converter", layout="wide")
st.title("📊 Image to Excel Converter (Offline / Free)")
st.write("Extracts text, numbers, and tabular data directly into Excel without external API keys.")

@st.cache_resource
def load_ocr_reader():
    return easyocr.Reader(['en'], gpu=False)

reader = load_ocr_reader()

uploaded_file = st.file_uploader("Upload Image (Table, Progress Report, Chart, Notes)", type=["jpg", "jpeg", "png"])

if uploaded_file and st.button("Convert to Excel"):
    img = Image.open(uploaded_file).convert('RGB')
    st.image(img, caption="Uploaded Image", width=450)
    
    with st.spinner("Extracting table data directly from image..."):
        img_np = np.array(img)
        
        # OCR Detection
        results = reader.readtext(img_np)
        
        if not results:
            st.warning("Image me koi readable text ya data nahi mila.")
        else:
            sorted_by_y = sorted(results, key=lambda r: r[0][0][1])
            
            data_rows = []
            current_row = []
            last_y = None
            row_threshold = 20
            
            for item in sorted_by_y:
                y_coord = item[0][0][1]
                x_coord = item[0][0][0]
                text = str(item[1]).strip()
                
                if last_y is None or abs(y_coord - last_y) < row_threshold:
                    current_row.append((x_coord, text))
                    if last_y is None:
                        last_y = y_coord
                else:
                    current_row.sort(key=lambda x: x[0])
                    data_rows.append([t[1] for t in current_row])
                    current_row = [(x_coord, text)]
                    last_y = y_coord
                    
            if current_row:
                current_row.sort(key=lambda x: x[0])
                data_rows.append([t[1] for t in current_row])
            
            # Normalize column counts
            max_cols = max(len(row) for row in data_rows) if data_rows else 1
            padded_rows = [row + [''] * (max_cols - len(row)) for row in data_rows]
            
            # Safe unique column headers to prevent PyArrow display error
            safe_columns = [f"Column_{i+1}" for i in range(max_cols)]
            df = pd.DataFrame(padded_rows, columns=safe_columns)
            
            # Clean string conversions
            df = df.astype(str)
            
            st.success("Extraction Complete!")
            st.write("### Data Preview:")
            st.table(df)
            
            # Excel export
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df.to_excel(writer, index=False, header=False)
            
            st.download_button(
                label="📥 Download Excel File (.xlsx)",
                data=output.getvalue(),
                file_name="extracted_table.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
