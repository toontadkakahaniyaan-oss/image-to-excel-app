import streamlit as st
import pandas as pd
from io import StringIO, BytesIO
from PIL import Image
import base64
from groq import Groq

# Groq API Key
api_key = st.secrets.get("GROQ_API_KEY") if "GROQ_API_KEY" in st.secrets else st.sidebar.text_input("Groq API Key (Free)", type="password")

st.set_page_config(page_title="Image to Excel Converter", layout="wide")
st.title("📊 Image to Excel / Data Converter")

if not api_key:
    st.warning("Please enter your Groq API Key in the sidebar to proceed (Get free key from console.groq.com).")
else:
    client = Groq(api_key=api_key)

uploaded_file = st.file_uploader("Upload Image (Table, Flowchart, Notes)", type=["jpg", "jpeg", "png"])
user_prompt = st.text_input("Formatting Instructions", "Extract all tabular and key data accurately into CSV format.")

if uploaded_file and api_key and st.button("Convert to Excel"):
    img = Image.open(uploaded_file)
    st.image(img, caption="Uploaded Image", width=400)
    
    with st.spinner("AI is analyzing and converting data..."):
        # Convert image to base64
        uploaded_file.seek(0)
        base64_image = base64.b64encode(uploaded_file.read()).decode('utf-8')
        
        prompt_text = f"""
        User instruction: '{user_prompt}'.
        Extract all table data, columns, and rows from this image.
        Format STRICTLY as plain comma-separated CSV text.
        Do NOT wrap in markdown code blocks (no ```csv or ```).
        Do NOT write any intro, conversational text, or explanations. Only pure CSV data.
        """
        
        try:
            chat_completion = client.chat.completions.create(
                messages=[
                    {
                        "role": "user",
                        "content": [
                            {"type": "text", "text": prompt_text},
                            {
                                "type": "image_url",
                                "image_url": {
                                    "url": f"data:image/jpeg;base64,{base64_image}",
                                },
                            },
                        ],
                    }
                ],
                model="llama-3.2-11b-vision-preview",
            )
            
            raw_text = chat_completion.choices[0].message.content.strip().replace("```csv", "").replace("```", "").strip()
            
            # CSV to DataFrame
            df = pd.read_csv(StringIO(raw_text))
            
            st.success("Extraction Complete!")
            st.write("### Data Preview:")
            st.dataframe(df, use_container_width=True)
            
            # Excel export
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
                st.text_area("AI Raw Text", raw_text)
