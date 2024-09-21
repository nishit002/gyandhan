import streamlit as st
import pandas as pd
from io import BytesIO
from docx import Document
import matplotlib.pyplot as plt

# Function to filter the data based on user selection
def filter_data(df, college=None, course=None):
    # Drop unnecessary columns (e.g., 'ID')
    df = df.drop(columns=['ID', 'Course_link'], errors='ignore')

    if college:
        df = df[df['college'] == college]
    if course:
        df = df[df['Course_name'] == course]
    
    # Remove columns with all NA values, but keep those with some valid data
    df = df.dropna(axis=1, how='all')
    
    # Transpose the dataframe to display headers as rows
    df = df.T.reset_index()
    df.columns = ['Field'] + [f'Value_{i}' for i in range(1, df.shape[1])]
    return df

# Function to convert DataFrame to a Word document
def df_to_word(df):
    doc = Document()
    table = doc.add_table(rows=1, cols=df.shape[1])

    # Add headers to the Word table
    hdr_cells = table.rows[0].cells
    for i, column in enumerate(df.columns):
        hdr_cells[i].text = str(column)

    # Add DataFrame rows to the Word table
    for index, row in df.iterrows():
        row_cells = table.add_row().cells
        for i, value in enumerate(row):
            row_cells[i].text = str(value)

    # Save to a buffer
    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

# Function to convert relevant columns to numeric
def convert_to_numeric(df, columns):
    for col in columns:
        df[col] = pd.to_numeric(df[col], errors='coerce')
    return df

# Updated function to plot a chart based on data
def plot_chart(df, column):
    # Drop any NaN values and check if the column has valid data to plot
    valid_data = df[column].dropna()
    
    if not valid_data.empty:
        plt.figure(figsize=(10, 6))
        valid_data.plot(kind='bar', color='skyblue')
        plt.title(f'Distribution of {column}')
        plt.xlabel(column)
        plt.ylabel('Values')
        st.pyplot(plt)
    else:
        st.write(f"No valid data to plot for '{column}'.")

# Streamlit app
st.set_page_config(layout="wide")  # Wide layout for better display

st.title('College & Course Filter App')

# Sidebar for user selections
st.sidebar.header('Filter Options')
uploaded_file = st.sidebar.file_uploader("Upload your Excel file", type=['xlsx'])

if uploaded_file is not None:
    # Read Excel file
    df = pd.read_excel(uploaded_file)

    # Convert specific columns to numeric (assume these are the numeric ones)
    numeric_columns = ['Fees', 'Duration', 'TOEFL', 'IELTS', 'PTE']
    df = convert_to_numeric(df, numeric_columns)

    # Display available options for filtering
    college_options = df['college'].dropna().unique()
    course_options = df['Course_name'].dropna().unique()

    selected_college = st.sidebar.selectbox('Select College', options=[None] + list(college_options))
    selected_course = st.sidebar.selectbox('Select Course', options=[None] + list(course_options))

    # Filter data based on selection
    filtered_data = filter_data(df, college=selected_college, course=selected_course)

    # Display the filtered data
    if not filtered_data.empty:
        st.header(f'Data for {selected_college} - {selected_course}')
        st.write('Filtered Data:')
        st.dataframe(filtered_data)

        # Convert DataFrame to Word and provide download option
        word_file = df_to_word(filtered_data)
        st.download_button(
            label="Download as Word",
            data=word_file,
            file_name="filtered_data.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )

        # Adding a chart for visualization
        st.header("Data Visualization")

        # List the numeric columns available for chart
        numeric_columns = df.select_dtypes(include='number').columns.tolist()
        if numeric_columns:
            chart_column = st.selectbox('Select column for chart:', options=numeric_columns)
            plot_chart(df, chart_column)
        else:
            st.write("No numeric data available for charting.")
    else:
        st.write("No data available for the selected college and course.")
