import streamlit as st
from docx import Document
from docx.shared import Inches
from PIL import Image
import zipfile
import io

st.set_page_config(page_title="JPG to DOCX Converter", layout="centered")

st.title("📸 JPG to DOCX Converter")
st.write("Upload JPG images and convert them to Word documents!")

# Create two columns for better layout
col1, col2 = st.columns([1, 1])

with col1:
    st.subheader("Upload Images")
    uploaded_files = st.file_uploader(
        "Choose JPG files",
        type=["jpg", "jpeg"],
        accept_multiple_files=True
    )

with col2:
    st.subheader("Status")
    if uploaded_files:
        st.success(f"✓ {len(uploaded_files)} file(s) selected")
    else:
        st.info("📤 No files selected yet")

def convert_jpg_to_docx(jpg_file):
    """
    Convert a single JPG file to DOCX format.
    
    Args:
        jpg_file: Uploaded file object from Streamlit
    
    Returns:
        BytesIO object containing the DOCX file
    """
    try:
        # Create a new Document
        doc = Document()

        # Open image and get dimensions
        img = Image.open(jpg_file)
        width, height = img.size

        # Calculate width in inches (max 6 inches to fit on page)
        max_width = 6.0
        aspect_ratio = height / width

        if width > height:
            # Landscape orientation
            doc_width = max_width
        else:
            # Portrait orientation
            doc_width = max_width

        # Save image to bytes for adding to document
        jpg_file.seek(0)
        img_bytes = jpg_file.read()
        img_buffer = io.BytesIO(img_bytes)

        # Add image to document
        doc.add_picture(img_buffer, width=Inches(doc_width))

        # Save document to bytes
        docx_buffer = io.BytesIO()
        doc.save(docx_buffer)
        docx_buffer.seek(0)

        return docx_buffer

    except Exception as e:
        st.error(f"Error converting {jpg_file.name}: {str(e)}")
        return None


if uploaded_files:
    st.divider()

    if st.button("🔄 Convert All to DOCX", use_container_width=True, type="primary"):
        # Create a progress bar
        progress_bar = st.progress(0)
        status_text = st.empty()

        # Create ZIP file in memory
        zip_buffer = io.BytesIO()

        successful = 0
        failed = 0

        with zipfile.ZipFile(zip_buffer, 'w') as zip_file:
            for idx, file in enumerate(uploaded_files):
                # Update progress
                progress = (idx + 1) / len(uploaded_files)
                progress_bar.progress(progress)
                status_text.text(f"Converting {idx + 1}/{len(uploaded_files)}: {file.name}")

                # Convert the file
                docx_buffer = convert_jpg_to_docx(file)

                if docx_buffer:
                    # Get filename without extension and add .docx
                    filename = file.name.rsplit('.', 1)[0]
                    zip_file.writestr(f"{filename}.docx", docx_buffer.getvalue())
                    successful += 1
                else:
                    failed += 1

        zip_buffer.seek(0)

        # Clear progress indicators
        progress_bar.empty()
        status_text.empty()

        # Show results
        st.divider()
        col1, col2, col3 = st.columns(3)

        with col1:
            st.metric("Successful", successful)
        with col2:
            st.metric("Failed", failed)
        with col3:
            st.metric("Total", len(uploaded_files))

        # Download button
        st.download_button(
            label="⬇️ Download All DOCX Files (ZIP)",
            data=zip_buffer,
            file_name="converted_documents.zip",
            mime="application/zip",
            use_container_width=True
        )

        st.success("✅ Conversion complete! Click the button above to download.")
        