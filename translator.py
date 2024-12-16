import boto3
import streamlit as st
import PyPDF2
from io import BytesIO
from docx import Document
from fpdf import FPDF
import os

# Initialize Amazon Translate client
def init_aws_translate():
    return boto3.client(
        service_name='translate',
        region_name='us-east-1',  # Adjust to your region
        aws_access_key_id=os.environ.get('aws_key_id'),  # Replace with your own
        aws_secret_access_key=os.environ.get('aws_secret_key')  # Replace with your own
    )

# Function to extract text from PDF
def extract_text_from_pdf(pdf_file):
    pdf_reader = PyPDF2.PdfReader(pdf_file)
    text = ''
    for page_num in range(len(pdf_reader.pages)):
        page = pdf_reader.pages[page_num]
        text += page.extract_text()
    return text

# Function to extract text from Word document
def extract_text_from_word(word_file):
    doc = Document(word_file)
    text = ''
    for para in doc.paragraphs:
        text += para.text + '\n'
    return text

# Function to translate text
def translate_text(client, text, source_language, target_language):
    response = client.translate_text(
        Text=text,
        SourceLanguageCode=source_language,
        TargetLanguageCode=target_language
    )
    return response['TranslatedText']

# Save translated content to Word
def save_as_word(translated_text, original_filename):
    doc = Document()
    doc.add_paragraph(translated_text)
    translated_filename = f"translated_{original_filename}.docx"
    doc.save(translated_filename)
    return translated_filename

# Save translated content to PDF
def save_as_pdf(translated_text, original_filename):
    # Initialize PDF
    pdf = FPDF()
    pdf.set_auto_page_break(auto=True, margin=15)
    pdf.add_page()
    
    # Add a custom font that supports Unicode
    pdf.add_font('DejaVu', '', 'DejaVuSans.ttf', uni=True)
    pdf.set_font('DejaVu', '', 12)
    
    # Add the translated text to the PDF
    pdf.multi_cell(0, 10, translated_text)
    
    # Save the document as a PDF file
    translated_filename = f"translated_{original_filename}.pdf"
    pdf.output(translated_filename)
    
    return translated_filename


# Save translated content as text file
def save_as_text(translated_text, original_filename):
    translated_filename = f"translated_{original_filename}.txt"
    with open(translated_filename, "w", encoding="utf-8") as f:
        f.write(translated_text)
    return translated_filename

# Streamlit app
def main():
    st.title("Document Translator with Amazon Translate")
    st.write("Upload a document (PDF, Word, or TXT), choose a target language, and translate!")

    # Translation settings
    target_language = st.selectbox(
        "Select Target Language", 
        ["fr", "ar", "es", "de", "it", "zh", "ja"], 
        index=0
    )

    source_language = st.text_input(
        "Enter Source Language Code (e.g., 'en' for English)", 
        value="en"
    )

    # File upload
    uploaded_file = st.file_uploader("Upload a document (PDF, Word, or TXT)", type=["pdf", "docx", "txt"])
    if uploaded_file is not None:
        st.success(f"Uploaded: {uploaded_file.name}")
        
        # Extract text based on file type
        if uploaded_file.type == "application/pdf":
            document_content = extract_text_from_pdf(uploaded_file)
        elif uploaded_file.type == "application/vnd.openxmlformats-officedocument.wordprocessingml.document":
            document_content = extract_text_from_word(uploaded_file)
        else:
            document_content = uploaded_file.read().decode("utf-8")
        
        # Display document preview
        st.subheader("Document Preview:")
        st.text_area("Preview", document_content[:500], height=200)

    # Translate button
    if st.button("Start Translation"):
        if uploaded_file is not None:
            with st.spinner("Translating document..."):
                try:
                    # Initialize Amazon Translate client
                    client = init_aws_translate()
                    
                    # Translate the content
                    translated_text = translate_text(client, document_content, source_language, target_language)
                    
                    # Display results
                    st.success("Translation completed!")
                    st.subheader("Translated Document:")
                    st.text_area("Translation Result", translated_text, height=300)
                    
                    # Save and allow downloading in original file format
                    if uploaded_file.type == "application/pdf":
                        translated_filename = save_as_pdf(translated_text, uploaded_file.name)
                    elif uploaded_file.type == "application/vnd.openxmlformats-officedocument.wordprocessingml.document":
                        translated_filename = save_as_word(translated_text, uploaded_file.name)
                    else:
                        translated_filename = save_as_text(translated_text, uploaded_file.name)
                    
                    # Provide download button
                    with open(translated_filename, "rb") as f:
                        st.download_button(
                            label="Download Translated Document",
                            data=f,
                            file_name=translated_filename,
                            mime="application/octet-stream"
                        )
                except Exception as e:
                    st.error(f"An error occurred: {e}")
        else:
            st.warning("Please upload a file to translate.")

if __name__ == "__main__":
    main()
