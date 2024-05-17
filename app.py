import os
import zipfile
from lxml import etree
import pandas as pd
import streamlit as st

# Ensure the uploads directory exists
if not os.path.exists("uploads"):
    os.makedirs("uploads")

def get_document_comments(docxFileName):
    comments_dict = {}
    comments_of_dict = {}
    docx_zip = zipfile.ZipFile(docxFileName)
    comments_xml = docx_zip.read('word/comments.xml')
    comments_of_xml = docx_zip.read('word/document.xml')
    et_comments = etree.XML(comments_xml)
    et_comments_of = etree.XML(comments_of_xml)
    ooXMLns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}

    comments = et_comments.xpath('//w:comment', namespaces=ooXMLns)
    comments_of = et_comments_of.xpath('//w:commentRangeStart', namespaces=ooXMLns)

    for c in comments:
        comment_text = c.xpath('string(.)', namespaces=ooXMLns)
        comment_id = c.xpath('@w:id', namespaces=ooXMLns)[0]
        comment_author = c.xpath('@w:author', namespaces=ooXMLns)[0]
        comments_dict[comment_id] = {'text': comment_text, 'author': comment_author}

    for c in comments_of:
        comments_of_id = c.xpath('@w:id', namespaces=ooXMLns)[0]
        parts = et_comments_of.xpath(
            "//w:r[preceding-sibling::w:commentRangeStart[@w:id=" + comments_of_id + "] and following-sibling::w:commentRangeEnd[@w:id=" + comments_of_id + "]]",
            namespaces=ooXMLns)
        comment_of = ''
        for part in parts:
            comment_of += part.xpath('string(.)', namespaces=ooXMLns)
            comments_of_dict[comments_of_id] = comment_of

    return comments_dict, comments_of_dict

def extract_comments_from_docx(docx_path):
    comments_data = []
    comments_dict, comments_of_dict = get_document_comments(docx_path)

    for comment_id, comment in comments_dict.items():
        referenced_text = comments_of_dict.get(comment_id, '')
        comments_data.append({
            'file_name': os.path.basename(docx_path),
            'comment': comment['text'],
            'author': comment['author'],
            'referenced_text': referenced_text
        })

    return comments_data

# Main UI
st.set_page_config(page_title="DOCX Comments Extractor", layout="wide")
st.title('Extract Qualitative Coding from Microsoft Word Documents')
st.markdown("""
    This application allows you to upload DOCX files and extracts the comments along with the author and referenced text. 
    <p> 
    You can upload multiple DOCX files at once.
    Ensure your filename is formatted correctly: [Interviewee Name]_[Coder Name].docx
    -Interviewee Name: The name of the person interviewed in the document.
    -Coder Name: The name of the person who performed the qualitative coding/theme identification. 
""")

uploaded_files = st.file_uploader("Choose DOCX files", type="docx", accept_multiple_files=True)

if uploaded_files:
    all_comments = []
    progress_bar = st.progress(0)
    for i, uploaded_file in enumerate(uploaded_files):
        file_path = os.path.join("uploads", uploaded_file.name)
        with open(file_path, "wb") as f:
            f.write(uploaded_file.getbuffer())
        
        comments = extract_comments_from_docx(file_path)
        all_comments.extend(comments)
        progress_bar.progress((i + 1) / len(uploaded_files))
    
    if all_comments:
        df = pd.DataFrame(all_comments)
        df[['Interviewee', 'Coder']] = df['file_name'].str.split('_', n=1, expand=True)
        df['Coder'] = df['Coder'].str.replace('.docx', '', regex=False)
        df = df[['Interviewee', 'Coder', 'comment', 'author', 'referenced_text', 'file_name']]
        
        st.success("File processing complete!")
        st.write(df)
        csv = df.to_csv(index=False)
        st.download_button(label="Download CSV", data=csv, file_name='comments_summary.csv', mime='text/csv')

# Additional Styling
st.markdown(
    """
    <style>
    .reportview-container {
        background-color: #f0f2f6;
    }
    .sidebar .sidebar-content {
        background-color: #f0f2f6;
    }
    </style>
    """,
    unsafe_allow_html=True
)
