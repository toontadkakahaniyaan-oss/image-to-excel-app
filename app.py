import streamlit as st
import pandas as pd
import requests
from io import StringIO, BytesIO
from PIL import Image
import base64

api_key = st.secrets.get("HF_TOKEN") if "HF_TOKEN" in st.secrets else st.sidebar.text_input("Hugging Face Free Token", type="password")

st.set_page_config(page_title="Image to Excel Converter", layout="wide")
st.title("📊 Image to Excel / Data Converter")

if not api_key:
    st.warning("Sidebar me Hugging Face Free Token dalein.")

uploaded_file = st.file_uploader("Upload Image (Table, Chart, Notes)", type=["jpg", "jpeg", "png"])
user_prompt = st.text_input("Formatting Instructions", "Extract all table columns and rows into clean comma-separated CSV format.")

if uploaded_file and api_key and st.button("Convert to Excel"):
    img = Image.open(uploaded_file)
    st.image(img, caption="Uploaded Image", width=400)
    
    with st.spinner("AI is processing the table data..."):
        uploaded_file.seek(0)
        base64_image = base64.b64encode(uploaded_file.read()).decode('utf-8')
        
        prompt_text = f"""
        User instruction: {user_prompt}
        Task: Extract all structured table data from this image.
        Format STRICTLY as valid standard CSV text with comma separators.
        Do NOT wrap in markdown blocks (no ```csv).
        Do NOT add conversational text. Only raw CSV lines.
        """
        
        headers = {
            "Authorization": f"Bearer {api_key}",
            "Content-Type": "application/json"
        }
        
        API_URL = "[https://api-inference.huggingface.co/models/Qwen/Qwen2-VL-7B-Instruct/v1/chat/completions](https://api-inference.huggingface.co/models/Qwen/Qwen2-VL-7B-Instruct/v1/chat/completions)"
        
        payload = {
            "model": "Qwen/Qwen2-VL-7B-Instruct",
            "messages": [
                {
                    "role": "user",
                    "content": [
                        {"type": "text", "text": prompt_text},
                        {
                            "type": "image_url",
                            "image_url": {"url": f"data:image/jpeg;base64,{base64_image}"}
                        }
                    ]
                }
            ],
            "max_tokens": 2048
        }
        
        response = requests.post(API_URL, headers=headers, json=payload)
        
        if response.status_code == 200:
            res_data = response.json()
            raw_text = res_data["choices"][0]["message"]["content"].strip()
            raw_text = raw_text.replace("```csv", "").replace("```", "").strip()
            
            try:
                df = pd.read_csv(StringIO(raw_text))
                st.success("Extraction Complete!")
                st.write("### Data Preview:")
                st.dataframe(df, use_container_width=True)
                
                output = BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    df.to_excel(writer, index=False)
                
                st.download_button(
                    label="📥 Download Excel (.xlsx)",
                    data=output.getvalue(),
                    file_name="extracted_data.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            except Exception as parse_err:
                st.warning("Table auto-parse me issue aaya, raw CSV niche hai:")
                st.text_area("Extracted CSV Data", raw_text)
                st.download_button("Download Raw CSV", raw_text, file_name="data.csv")
        else:
            st.error(f"Inference Error ({response.status_code}): {response.text}")
