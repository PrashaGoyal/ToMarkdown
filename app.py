import streamlit as st
import os
import tempfile
from markitdown import MarkItDown
from pathlib import Path

# --- Configuration & UI Setup ---
st.set_page_config(page_title="MarkItDown Universal Converter", page_icon="📝", layout="wide")
st.theme = "light" # Streamlit defaults to user preference, but we aim for clean white/professional

# Initialize the MarkItDown Engine
# User-Agent handling is often managed by underlying engines (like requests for HTML), 
# but for internal stability, we initialize the converter here.
md = MarkItDown()

st.title("🚀 Universal Document-to-Text Converter")
st.markdown("""
Upload **Word, Excel, PPT, PDF, or HTML** files. 
We'll convert them into clean Markdown instantly.
""")

# --- File Upload Area ---
uploaded_files = st.file_uploader(
    "Drag and drop files here", 
    type=["docx", "xlsx", "pptx", "pdf", "html", "htm"], 
    accept_multiple_files=True
)

if uploaded_files:
    st.divider()
    
    for uploaded_file in uploaded_files:
        file_extension = Path(uploaded_file.name).suffix.lower()
        base_name = Path(uploaded_file.name).stem
        
        try:
            # MarkItDown requires a file path or stream. 
            # Using a temporary file ensures high resilience for complex MS formats.
            with tempfile.NamedTemporary_TemporaryFile(delete=False, suffix=file_extension) as tmp_file:
                tmp_file.write(uploaded_file.getvalue())
                tmp_path = tmp_file.name

            # Processing Step
            with st.spinner(f"Processing {uploaded_file.name}..."):
                # MarkItDown conversion logic
                result = md.convert(tmp_path)
                converted_text = result.text_content

            # --- Display & UX ---
            with st.expander(f"📄 Preview: {uploaded_file.name}", expanded=True):
                # Scrollable box for instant preview
                st.text_area(
                    label="Converted Content",
                    value=converted_text,
                    height=300,
                    key=f"text_{uploaded_file.name}"
                )
                
                # Download Options
                col1, col2 = st.columns(2)
                
                with col1:
                    st.download_button(
                        label="Download as Markdown (.md)",
                        data=converted_text,
                        file_name=f"{base_name}_converted.md",
                        mime="text/markdown",
                        key=f"md_{uploaded_file.name}"
                    )
                
                with col2:
                    st.download_button(
                        label="Download as Plain Text (.txt)",
                        data=converted_text,
                        file_name=f"{base_name}_converted.txt",
                        mime="text/plain",
                        key=f"txt_{uploaded_file.name}"
                    )
            
            # Clean up temp file
            os.remove(tmp_path)

        except Exception as e:
            st.error(f"⚠️ Could not read {uploaded_file.name}. Please check the format.")
            # Log error for developer but keep UI clean for user
            print(f"Error processing {uploaded_file.name}: {e}")

else:
    st.info("Waiting for your files... Upload them above to start the magic.")

# Footer
st.markdown("---")
st.caption("Built with Microsoft MarkItDown & Streamlit | 2026 Edition")
