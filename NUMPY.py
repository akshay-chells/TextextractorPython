import os
from docx import Document
import streamlit as st


def extract_images_from_word(uploaded_file):
    st.write("Processing Word file...")
    document = Document(uploaded_file)

    # Create directory to save extracted images
    image_dir = "extracted_images"
    if not os.path.exists(image_dir):
        os.makedirs(image_dir)

    image_count = 0
    for rel in document.part.rels.values():
        if "image" in rel.target_ref:
            image_count += 1
            img = rel.target_part.blob
            img_path = os.path.join(image_dir, f"{uploaded_file.name}_image_{image_count}.png")
            with open(img_path, "wb") as img_file:
                img_file.write(img)
            st.write(f"Extracted image {image_count}: {img_path}")

    if image_count == 0:
        st.write("No images found in the document.")
    else:
        st.success(f"Successfully extracted {image_count} images from the document.")


# Streamlit app
st.title("Word Document Image Extractor")
uploaded_file = st.file_uploader("Upload a Word document", type=["docx"])

if uploaded_file:
    extract_images_from_word(uploaded_file)
