import os
import streamlit as st
from pptx import Presentation
from pptx.util import Inches, Pt
from pypdf import PdfReader
#Stremlit Page Configuration
st.set_page_config(
    page_title="PDF to PowerPoint Converter", page_icon="📊", layout="centered"
)

st.title("📄 PDF to PowerPoint Converter")
st.write("Convert your PDF documents into editable PowerPoint (.pptx) slides instantly.")

#File uploader
uploaded_file = st.file_uploader("Upload a PDF file", type=["pdf"])
#PDF checking and writing it into the temporary local file(buffer)
if uploaded_file is not None:

  input_pdf_path = "temp_input.pdf"
  output_ppt_path = "output_presentation.pptx"

   # Save uploaded PDF temporarily
  with open(input_pdf_path, "wb") as f:
    f.write(uploaded_file.getbuffer())

  # Conversion button
  if st.button("Convert to PPTX"):
    with st.spinner("Converting pages... Please wait."):
      try:
        # Check if file exists before reading
        if not os.path.exists(input_pdf_path):
          st.error("Uploaded file path not found. Please re-upload.")
        else:
          reader = PdfReader(input_pdf_path)
          presentation = Presentation()
          presentation.slide_width = Inches(13.333)
          presentation.slide_height = Inches(7.5)
          blank_layout = presentation.slide_layouts[6]

          for index, page in enumerate(reader.pages):
            slide = presentation.slides.add_slide(blank_layout)
            page_text = page.extract_text() or ""

            text_box = slide.shapes.add_textbox(
                Inches(1.0), Inches(1.0), Inches(11.333), Inches(5.5)
            )
            text_frame = text_box.text_frame
            text_frame.word_wrap = True

            lines = page_text.split("\n")
            for line_index, line_text in enumerate(lines):
              if line_index == 0:
                paragraph = text_frame.paragraphs[0]
              else:
                paragraph = text_frame.add_paragraph()
              paragraph.text = line_text.strip()
              paragraph.font.size = Pt(16)

          presentation.save(output_ppt_path)
          st.success("Conversion completed successfully!")

          # Provide download button
          with open(output_ppt_path, "rb") as file:
            st.download_button(
                label="📥 Download PowerPoint Presentation",
                data=file,
                file_name="converted_presentation.pptx",
                mime=(
                    "application/vnd.openxmlformats-officedocument.presentationml.presentation"
                ),
            )
      except Exception as e:
        st.error(f"An error occurred during conversion: {e}")
