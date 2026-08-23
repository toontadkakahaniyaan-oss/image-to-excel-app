import streamlit as st
from google import genai
import pandas as pd
from io import StringIO, BytesIO
from PIL import Image

api_key = st.secrets.get("GEMINI_API_KEY") if "GEMINI_API_KEY" in st.secrets else st.sidebar.text_input("Gemini API Key", type="password")

st.set_page_config(page_title="Image to Excel Converter", layout="wide")
st.title("📊 Image to Excel / Data Converter")

if not api_key:
    st.warning("Please enter your Gemini API Key in the sidebar to proceed.")
else:
    client = genai.Client(api_key=api_key)

uploaded_file = st.file_uploader("Upload Image (Table, Flowchart, Notes)", type=["jpg", "jpeg", "png"])
user_prompt = st.text_input("Formatting Instructions", "Extract all key data and structure it into clean tabular CSV format.")

if uploaded_file and api_key and st.button("Convert to Excel"):
    img = Image.open(uploaded_file)
    st.image(img, caption="Uploaded Image", width=400)
    
    with st.spinner("AI is analyzing and converting data..."):
        prompt_text = f"""
        {user_prompt}
        CRITICAL: Return ONLY raw, valid comma-separated CSV text.
        Do NOT write markdown code blocks (no ```csv or ```).
        Do NOT write explanations or greetings.
        """
        
        try:
            response = client.models.generate_content(
                model='gemini-2.5-flash',
                contents=[img, prompt_text]
            )
            
            raw_text = response.text.strip().replace("```csv", "").replace("```", "").strip()
            
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
            if 'response' in locals() and hasattr(response, 'text'):
                st.text_area("AI Raw Text", response.text)
