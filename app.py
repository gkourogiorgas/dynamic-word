import streamlit as st
import pandas as pd
from docxtpl import DocxTemplate
from zipfile import ZipFile
import os
import base64

st.set_page_config(page_title="Dynamic Word Application", layout="wide")

HELP_TEXT = """
<p>Uses a word file as a template with "tags" in the form of {{ key1 }}. Then uses a csv file as a list where the headers of the columns are also "key1", "key2" etc.</p>
<p>You can use as many different tags as you want and you can use the same tag as many times as you want.</p>
<p>Please refer to the demo files to get a better understanding.</p>
<p>To create a csv file from an excel file, make sure that:
<ol>
    <li>The table with "key" headers is in cell A1 in a specific worksheet</li>
    <li>There is no other data in this worksheet apart from the table above</li>
    <li>Having the worksheet with the table active, click <em>save as</em> and choose <em>csv</em> as file type</li>
</ol>
</p>
<h3>Data privacy</h3>
<p>The uploaded files are stored in memory for as long as the session is active. The produced word files are deleted once they are added to the zip file. The zip file is deleted once downloaded</p>
<p>Made by <a href="https://www.linkedin.com/in/gkourogiorgas/" >Georgios Kourogiorgas</a></p>
"""

def get_binary_file_downloader_html(bin_file, file_label='File'):
    with open(bin_file, 'rb') as f:
        data = f.read()
    bin_str = base64.b64encode(data).decode()
    href = f'<a href="data:application/octet-stream;base64,{bin_str}" download="{os.path.basename(bin_file)}">Download a demo {file_label}</a>'
    return href

def dynamic_word(context,wordfile):
    if len(context)>0:
        zipObj = ZipFile('download.zip', 'w')
        for i in range(len(context)):
            doc = DocxTemplate(wordfile)
            doc.render(context.iloc[i,:].to_dict())
            endfilename = str(i)+"_"+wordfile.name
            doc.save(endfilename)
            zipObj.write(endfilename)
            os.remove(endfilename)
        zipObj.close()
        with open("download.zip", "rb") as fp:
            btn = st.download_button(label="Download ZIP",data=fp,file_name="download.zip",mime="application/zip")
            os.remove("download.zip")
def openfile(uploaded_file):
    if uploaded_file is not None:
        if uploaded_file.name.endswith('.csv'):
            processed_file = pd.read_csv(uploaded_file)
            st.write('List:')
            st.write(processed_file)
            return processed_file
        elif uploaded_file.name.endswith('.doc') or uploaded_file.name.endswith('.docx'):
            st.write('Word file uploaded')
        else:
            st.warning('There seems to be a problem with the file types')
col1, col2 = st.columns([10, 10])
with col1:
    st.title('Dynamic Word Application')
    # Working with File Upload
    st.markdown(get_binary_file_downloader_html('DemoData/list.csv', 'list'), unsafe_allow_html=True)
    listuploaded_file = st.file_uploader("Upload List File",type=["csv"])
    st.markdown(get_binary_file_downloader_html('DemoData/example.docx', 'example'), unsafe_allow_html=True)
    worduploaded_file = st.file_uploader("Upload Word File",type=["doc","docx"])


    if st.button("Submit"):
        listdf = openfile(listuploaded_file)
        openfile(worduploaded_file)
        dynamic_word(listdf,worduploaded_file)
with col2:
    st.title('Help:')
    st.markdown(HELP_TEXT, unsafe_allow_html=True)
