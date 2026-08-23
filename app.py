import streamlit as st
import google.generativeai as genai
import pandas as pd
from io import StringIO, BytesIO
from PIL import Image

# Streamlit Secrets ya UI se API key lena
api_key = st.secrets.get("GEMINI_API_KEY") if "GEMINI_API_KEY" in st.secrets else st.sidebar.text_input("Gemini API Key", type="password")

st.set_page_config(page_title="Image to Excel Converter", layout="wide")
st.title("📊 Image to Excel / Data Converter")

if not api_key:
    st.warning("Please enter your Gemini API Key in sidebar or configure Secrets to proceed.")
else:
    genai.configure(api_key=api_key)

uploaded_file = st.file_uploader("Upload Image (Table, Chart, Notes, etc.)", type=["jpg", "jpeg", "png"])
user_prompt = st.text_input("Formatting Instructions", "Extract all tabular and key data accurately into CSV table format.")

if uploaded_file and api_key and st.button("Convert to Excel"):
    img = Image.open(uploaded_file)
    st.image(img, caption="Uploaded Image", width=400)
    
    with st.spinner("AI is analyzing and converting data..."):
        system_instruction = f"""
        Analyze the uploaded image according to the instruction: '{user_prompt}'.
        Extract and format the content STRICTLY into standard comma-separated CSV format.
        Do not add code blocks (no ```csv), no markdown explanations, and no intro sentences. Only raw CSV rows.
        """
        
        # Available models ki priority list
        candidate_models = ['gemini-2.5-flash', 'gemini-2.0-flash', 'gemini-flash-latest', 'gemini-1.5-flash']
        
        response = None
        for model_name in candidate_models:
            try:
                model = genai.GenerativeModel(model_name)
                response = model.generate_content([system_instruction, img])
                if response:
                    break
            except Exception:
                continue

        if response is None:
            st.error("Model connect nahi ho paya. Kripya check karein ki API Key sahi hai.")
        else:
            try:
                raw_text = response.text.strip().replace("```csv", "").replace("```", "").strip()
                
                # Convert CSV string to DataFrame
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
                st.error(f"Error structuring data: {e}")
                st.text_area("Raw Extracted Output", response.text)
            # Convert CSV string to DataFrame
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
            st.error(f"Error structuring data: {e}")
            st.text_area("Raw Extracted Output", response.text if 'response' in locals() else "No response")
