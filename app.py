import streamlit as st
import pandas as pd
from io import StringIO, BytesIO
from PIL import Image
import base64
from openai import OpenAI

api_key = st.secrets.get("OPENROUTER_API_KEY") if "OPENROUTER_API_KEY" in st.secrets else st.sidebar.text_input("OpenRouter Free API Key", type="password")

st.set_page_config(page_title="Image to Excel Converter", layout="wide")
st.title("📊 Image to Excel / Data Converter")

if not api_key:
    st.warning("Sidebar me OpenRouter Free API Key dalein (Get free key from openrouter.ai/keys)")
else:
    client = OpenAI(
        base_url="https://openrouter.ai/api/v1",
        api_key=api_key
    )

uploaded_file = st.file_uploader("Upload Image (Table, Chart, Notes)", type=["jpg", "jpeg", "png"])
user_prompt = st.text_input("Formatting Instructions", "Extract all table and chart data into clean comma-separated CSV format.")

if uploaded_file and api_key and st.button("Convert to Excel"):
    img = Image.open(uploaded_file)
    st.image(img, caption="Uploaded Image", width=400)
    
    with st.spinner("AI is analyzing and converting data..."):
        uploaded_file.seek(0)
        base64_image = base64.b64encode(uploaded_file.read()).decode('utf-8')
        
        prompt_text = f"""
        User instruction: '{user_prompt}'.
        Extract all table columns, rows, and flowchart steps accurately from the image.
        Format STRICTLY as valid comma-separated CSV text.
        Do NOT wrap in markdown codeblocks (no ```csv or ```).
        Do NOT write explanations or greetings. Only output CSV rows.
        """
        
        try:
            response = client.chat.completions.create(
                model="google/gemini-2.0-flash-lite-preview-02-05:free",
                messages=[
                    {
                        "role": "user",
                        "content": [
                            {"type": "text", "text": prompt_text},
                            {
                                "type": "image_url",
                                "image_url": {
                                    "url": f"data:image/jpeg;base64,{base64_image}"
                                }
                            }
                        ]
                    }
                ]
            )
            
            raw_text = response.choices[0].message.content.strip().replace("```csv", "").replace("```", "").strip()
            
            # CSV to DataFrame
            df = pd.read_csv(StringIO(raw_text))
            
            st.success("Extraction Complete!")
            st.write("### Data Preview:")
            st.dataframe(df, use_container_width=True)
            
            # Export to Excel
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df.to_excel(writer, index=False)
            
            st.download_button(
                label="📥 Download Excel (.xlsx)",
                data=output.getvalue(),
                file_name="extracted_data.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        except Exception as e:
            st.error(f"Error: {e}")
            if 'raw_text' in locals():
                st.text_area("Raw AI Text", raw_text)
