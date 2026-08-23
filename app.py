import streamlit as st
import pandas as pd
from io import StringIO, BytesIO
from PIL import Image
import base64
from huggingface_hub import InferenceClient

api_key = st.secrets.get("HF_TOKEN") if "HF_TOKEN" in st.secrets else st.sidebar.text_input("Hugging Face Free Token", type="password")

st.set_page_config(page_title="Image to Excel Converter", layout="wide")
st.title("📊 Image to Excel / Data Converter")

if not api_key:
    st.warning("Sidebar me Hugging Face Free Token dalein.")
else:
    client = InferenceClient(api_key=api_key)

uploaded_file = st.file_uploader("Upload Image (Table, Chart, Notes)", type=["jpg", "jpeg", "png"])
user_prompt = st.text_input("Formatting Instructions", "Extract all tabular and key data accurately into CSV table format.")

if uploaded_file and api_key and st.button("Convert to Excel"):
    img = Image.open(uploaded_file)
    st.image(img, caption="Uploaded Image", width=400)
    
    with st.spinner("AI is processing and creating Excel..."):
        uploaded_file.seek(0)
        base64_image = base64.b64encode(uploaded_file.read()).decode('utf-8')
        
        prompt_text = f"""
        User instruction: '{user_prompt}'.
        Extract all table columns, rows, and structured data accurately from this image.
        Format STRICTLY as valid standard CSV text with comma separators.
        Do NOT wrap in markdown code blocks (no ```csv or ```).
        Do NOT add intro sentences or conversation. Only raw CSV rows.
        """
        
        try:
            response = client.chat.completions.create(
                model="Qwen/Qwen2.5-VL-7B-Instruct",
                messages=[
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
                max_tokens=2048
            )
            
            raw_text = response.choices[0].message.content.strip().replace("```csv", "").replace("```", "").strip()
            
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
        except Exception as e:
            st.error(f"Error: {e}")
            if 'raw_text' in locals():
                st.text_area("Extracted Raw CSV", raw_text)
                st.download_button("Download Raw CSV", raw_text, file_name="data.csv")
