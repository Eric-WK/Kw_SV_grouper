#!/usr/bin/env python3

"""This is the main entry point of the program."""
## it will be a streamlit app 
import streamlit as st
import pandas as pd 
from helpers import get_fuzzy_match_score, to_excel





## Title of the app
st.title("Keyword Search Volume Grouper")

## two columns 
top, bottom = st.columns(2)

## topthreshold input
TOP_THRESHOLD = top.number_input("Top Threshold", min_value=0, max_value=100, value=100, step=1)

## bottomthreshold input
BOTTOM_THRESHOLD = bottom.number_input("Bottom Threshold", min_value=0, max_value=100, value=40, step=1)

st.markdown("Input either the sheet number or the sheet name but not both. ")

## Excel sheet number (0, 1, 2)
sheetnumber = st.number_input("Excel Sheet Number", min_value=0, max_value=10, value=1, step=1)

## or define the sheet name 
sheetname = st.text_input("Excel Sheet Name", value="")

## upload the excel file
uploaded_file = st.file_uploader("Choose a file", type="xlsx")

if uploaded_file is not None: 
    fn = uploaded_file.name
    ## split by the dot, and get the last element
    country_flag = fn.split('- ')[-1].split('.')[0]
    keyword_col_name = f"Keyword - {country_flag}"
    ## save to session_state
    if 'keyword_col_name' not in st.session_state:
        st.session_state.keyword_col_name = keyword_col_name

## input the column name for the keywords
kw_col = st.text_input("Keywords Column Name", value = st.session_state.keyword_col_name if uploaded_file is not None else "")

## if the file is uploaded
if uploaded_file is not None:
    ## either the sheetnumber or the sheetname is given, not both
    if sheetnumber is not None:
        ## read the excel file
        df = pd.read_excel(uploaded_file, sheet_name=sheetnumber)
    elif sheetname is not None:
        df = pd.read_excel(uploaded_file, sheet_name=sheetname)
    else:
        st.write("Please provide the sheet number or sheet name")


def process_dataframe(dataframe: pd.DataFrame) -> pd.DataFrame:
    """Groups the dataframe by Average MSV, and then performs the Fuzzy String Matching on the keywords in that list"""
    msv_col = "Average MSV"
    ## groupby Average MSV and join the keywords 
    df_grouped = df.groupby(msv_col)[kw_col].apply(lambda x: ', '.join(x)).reset_index()

    ## for each keyword in the keywords column, use the first keyword as the pivot, and calculate the fuzzy match score with the rest of the keywords
    ## then remove the ones with a score of 100, and one below 60 
    ## create a new column to hold the keywords that are similar to the pivot keyword, joined by a "/"
    df_grouped['similar_keywords'] = df_grouped[kw_col].apply(lambda x: '/'.join([kw for kw in x.split(',') if get_fuzzy_match_score(x.split(',')[0], kw) > BOTTOM_THRESHOLD and get_fuzzy_match_score(x.split(',')[0], kw) < TOP_THRESHOLD]))

    ## add the similar_keywords column to the original dataframe, by matching the Average MSV 
    df_merged = df.merge(df_grouped[[msv_col, 'similar_keywords']], on=msv_col, how='left')
    
    for index, row in df_merged.iterrows():
    ## if the keyword in the kw_col is in the similar_keywords col, then remove it from the similar_keywords col
        if row[kw_col] in row['similar_keywords']:
            ## split the similar_keywords col by "/"
            similar_keywords = row['similar_keywords'].split('/')
            ## remove the keyword from the similar_keywords col
            if row[kw_col] in similar_keywords:
                similar_keywords.remove(row[kw_col])
            ## join the similar_keywords col by "/"
            df_merged.at[index, 'similar_keywords'] = '/'.join(similar_keywords)
    return df_merged 


## make two columns: process and download file 
process = st.button("Process File")

if process:
    ## process the dataframe
    df_merged = process_dataframe(df)
    ## add df_merged to session_state
    if 'df_merged' not in st.session_state:
        st.session_state.df_merged = df_merged
    ## convert the dataframe to excel
    df_excel = to_excel(df_merged)
    ## download the processed dataframe
    # st.download_button(label='📥 Download Current Result',data = df_excel, file_name='Final_Results.xlsx')
    ## save the df_excel to a session_state
    if 'df_excel' not in st.session_state:
        st.session_state.df_excel = df_excel

## show a button to preview the dataframe 
if 'df_excel' in st.session_state:
    if st.button("Preview Processed File"):
        st.dataframe(st.session_state.df_merged)

## download button: 
if 'df_excel' in st.session_state:
    st.download_button(label='📥 Download Current Result',data = st.session_state.df_excel, file_name='Final_Results.xlsx')