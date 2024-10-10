import streamlit as st
import pandas as pd
import plotly.express as px
from docx import Document
import base64
from io import BytesIO
from jinja2 import Template

# Function to generate detailed content based on course and country selected
def generate_content(course_details, graph_data, college_collections, course_name, country):
    selected_course = course_details[(course_details['Course Name'] == course_name) & 
                                     (course_details['Country'] == country)]
    
    content = f"# {course_name} in {country}\n\n"
    
    # Overview Section
    content += f"## Overview\n"
    content += (f"The {course_name} in {country} is one of the most sought-after courses for its blend of practical and theoretical knowledge. "
                f"This course is typically offered by world-renowned universities like {selected_course.iloc[0]['University Name']}. "
                f"The field has seen substantial growth, especially in the specializations like {selected_course.iloc[0]['Top Specializations']}, "
                f"making it a strong choice for students looking to enhance their careers in technology.\n\n")
    
    # Cost and Living Expenses
    content += f"## Cost and Living Expenses\n"
    content += (f"Pursuing {course_name} in {country} comes with a cost. The tuition fees range between "
                f"{selected_course.iloc[0]['Tuition Fees (INR)']}, while living expenses amount to approximately "
                f"{selected_course.iloc[0]['Living Expenses (INR)']} per year. These costs, while substantial, are offset by the potential return on investment (ROI) "
                f"for graduates from universities like {selected_course.iloc[0]['University Name']} which boast a high ROI of {selected_course.iloc[0]['Net ROI (USD)']}.\n\n")
    
    # Admission Requirements
    content += f"## Admission Requirements\n"
    content += (f"Students looking to enroll in {course_name} in {country} must meet certain academic and standardized test criteria. "
                f"These include {selected_course.iloc[0]['Admission Requirements (GRE, GPA, etc.)']}. Admission is highly competitive, with acceptance rates as low as "
                f"{selected_course.iloc[0]['Acceptance Rate (%)']}% for top universities like {selected_course.iloc[0]['University Name']}.\n\n")
    
    # Job Prospects and ROI
    content += f"## Job Prospects and ROI\n"
    content += (f"Graduating from {course_name} in {country} opens up multiple opportunities in fields such as {selected_course.iloc[0]['Top Specializations']}. "
                f"Graduates often find employment with industry giants such as {selected_course.iloc[0]['Top Employers']}. "
                f"The median salary for professionals in this field stands at {selected_course.iloc[0]['Median Salary (USD)']}, making it a lucrative option. "
                f"With a net ROI of {selected_course.iloc[0]['Net ROI (USD)']}, this degree offers both immediate and long-term financial benefits.\n\n")

    return content

# Phase 1: Function to generate Word document
def generate_word(content):
    doc = Document()
    doc.add_heading(content.split('\n')[0], 0)
    
    paragraphs = content.split('\n\n')
    for para in paragraphs[1:]:
        doc.add_paragraph(para)
    
    # Save document to memory
    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    
    return buffer

# Phase 1: Function to generate HTML from content
def generate_html(content):
    template = Template("""
    <html>
    <head>
        <title>{{ title }}</title>
    </head>
    <body>
        {% for section in sections %}
            <h2>{{ section.title }}</h2>
            <p>{{ section.content }}</p>
        {% endfor %}
    </body>
    </html>
    """)
    
    sections = []
    paragraphs = content.split('\n\n')
    for para in paragraphs:
        title, *body = para.split('\n')
        sections.append({"title": title, "content": ' '.join(body)})
    
    html_content = template.render(title=sections[0]['title'], sections=sections)
    
    return html_content

# Phase 1: Function to download files
def download_link(file_data, file_name, file_label):
    b64 = base64.b64encode(file_data).decode()
    href = f'<a href="data:application/octet-stream;base64,{b64}" download="{file_name}">{file_label}</a>'
    return href

# Streamlit App: Phase 1
def main():
    st.title("Content Automation for Course Details")
    st.write("Upload an Excel file containing course and university details to generate a comprehensive content template.")

    # File upload
    uploaded_file = st.file_uploader("Upload Excel file", type="xlsx")
    
    if uploaded_file:
        # Read the uploaded Excel file
        course_details = pd.read_excel(uploaded_file, sheet_name="Course Details")
        graph_data = pd.read_excel(uploaded_file, sheet_name="Graph Data")
        college_collections = pd.read_excel(uploaded_file, sheet_name="College Collections")
        
        # User selects the course and country
        course_name = st.selectbox("Select Course Name", course_details['Course Name'].unique())
        country = st.selectbox("Select Country", course_details['Country'].unique())
        
        # Generate the content
        generated_content = generate_content(course_details, graph_data, college_collections, course_name, country)
        
        # Display generated content
        st.subheader("Generated Content")
        st.text_area("Content", value=generated_content, height=500)

        # Plotly graph: Tuition Fees
        selected_graph_data = graph_data[graph_data['University Name'].isin(course_details['University Name'])]
        fig = px.bar(selected_graph_data, x='University Name', y='Tuition Fees (INR)', title='Tuition Fees by University')
        st.plotly_chart(fig)
        
        # Download options: Word and HTML
        word_file = generate_word(generated_content)
        st.markdown(download_link(word_file.read(), f"{course_name}_content.docx", "Download Word Document"), unsafe_allow_html=True)
        
        html_content = generate_html(generated_content)
        st.markdown(download_link(html_content.encode(), f"{course_name}_content.html", "Download HTML Document"), unsafe_allow_html=True)

if __name__ == "__main__":
    main()
# Adding more detailed sections for College Collections, Scholarships, and Visa Information
def generate_extended_content(course_details, graph_data, college_collections, course_name, country):
    selected_course = course_details[(course_details['Course Name'] == course_name) & 
                                     (course_details['Country'] == country)]
    
    # Start with the existing content from Phase 1
    content = f"# {course_name} in {country}\n\n"
    
    # Overview Section (Extended)
    content += f"## Overview\n"
    content += (f"The {course_name} in {country} is a highly prestigious degree that attracts students from all over the world. "
                f"Not only does it offer top-tier education in {selected_course.iloc[0]['Top Specializations']}, "
                f"but universities like {selected_course.iloc[0]['University Name']} provide world-class research facilities.\n\n")
    
    # Adding more details about the academic rigor and flexibility of the program
    content += (f"This course typically spans 2 years with options to specialize in fields such as {selected_course.iloc[0]['Top Specializations']}. "
                f"Many institutions offer part-time study options or flexible course schedules to accommodate working professionals.\n\n")
    
    # Cost and Living Expenses (Extended)
    content += f"## Cost and Living Expenses\n"
    content += (f"The total cost of pursuing {course_name} in {country} includes tuition fees ranging between {selected_course.iloc[0]['Tuition Fees (INR)']} "
                f"and living expenses of approximately {selected_course.iloc[0]['Living Expenses (INR)']} annually. "
                f"Students should budget for additional expenses such as travel, accommodation, books, and supplies.\n\n")

    # Adding comparative analysis of universities with graphs
    content += "### Tuition Fees and ROI Comparison\n"
    content += "Here's a comparison of the top universities offering the best return on investment (ROI):\n\n"
    
    # Adding College Collection (Best ROI)
    content += "## Best ROI Universities for MS in Computer Science\n"
    selected_collections = college_collections[college_collections['Collection Name'] == 
                                               f"Best ROI for {course_name} ({country})"]
    content += "| University Name | Tuition Fees | Acceptance Rate | Application Link |\n"
    content += "|-----------------|--------------|-----------------|------------------|\n"
    for _, row in selected_collections.iterrows():
        content += (f"| {row['University Name']} | {row['Tuition Fees (INR)']} | "
                    f"{row['Acceptance Rate (%)']}% | [Apply Here]({row['Application Link']}) |\n")
    
    # Scholarships and Financial Aid Section
    content += "\n## Scholarships and Financial Aid\n"
    content += (f"To support international students, there are several scholarships available, including {selected_course.iloc[0]['Scholarships Available']}. "
                f"Many institutions also offer merit-based scholarships, while government-backed grants are often available for students in STEM fields.\n\n")
    
    # Visa Information
    content += "## Visa Information\n"
    content += (f"International students pursuing {course_name} in {country} typically apply for the {selected_course.iloc[0]['Visa Types']}. "
                f"The application process requires proof of admission, financial stability, and adherence to local immigration laws. "
                f"Students should plan to apply for a visa at least 3 months before the start of their program.\n\n")

    # Adding more graphs and analytics related to costs
    content += "## Cost Analysis and Budgeting\n"
    content += (f"While the tuition and living expenses may seem high, the return on investment for graduates of {selected_course.iloc[0]['University Name']} "
                f"justifies the cost. Graduates in fields like {selected_course.iloc[0]['Top Specializations']} see an average starting salary of "
                f"{selected_course.iloc[0]['Median Salary (USD)']}, and with a net ROI of {selected_course.iloc[0]['Net ROI (USD)']}, this degree "
                f"is a sound investment for the future.\n\n")

    # Adding YouTube guide
    content += "\n## Watch a Detailed Guide\n"
    content += f"[YouTube Video]({selected_course.iloc[0]['Relevant YouTube Video URL']})\n\n"

    return content

# Phase 2: Modified Streamlit App with more sections
def main_phase_2():
    st.title("Content Automation for Course Details - Extended Version")
    st.write("Upload an Excel file containing course and university details to generate a comprehensive content template with more analysis.")

    # File upload
    uploaded_file = st.file_uploader("Upload Excel file", type="xlsx")
    
    if uploaded_file:
        # Read the uploaded Excel file
        course_details = pd.read_excel(uploaded_file, sheet_name="Course Details")
        graph_data = pd.read_excel(uploaded_file, sheet_name="Graph Data")
        college_collections = pd.read_excel(uploaded_file, sheet_name="College Collections")
        
        # User selects the course and country
        course_name = st.selectbox("Select Course Name", course_details['Course Name'].unique())
        country = st.selectbox("Select Country", course_details['Country'].unique())
        
        # Generate the content (extended version)
        generated_content = generate_extended_content(course_details, graph_data, college_collections, course_name, country)
        
        # Display generated content
        st.subheader("Generated Content")
        st.text_area("Content", value=generated_content, height=500)
        
        # Plotly graph: Tuition Fees by University (Graph Data)
        selected_graph_data = graph_data[graph_data['University Name'].isin(course_details['University Name'])]
        fig = px.bar(selected_graph_data, x='University Name', y='Tuition Fees (INR)', title='Tuition Fees by University')
        st.plotly_chart(fig)

        # Word and HTML download options
        word_file = generate_word(generated_content)
        st.markdown(download_link(word_file.read(), f"{course_name}_extended_content.docx", "Download Word Document"), unsafe_allow_html=True)
        
        html_content = generate_html(generated_content)
        st.markdown(download_link(html_content.encode(), f"{course_name}_extended_content.html", "Download HTML Document"), unsafe_allow_html=True)

if __name__ == "__main__":
    main_phase_2()
# Adding detailed visualizations and career growth prospects analysis
def generate_final_content(course_details, graph_data, college_collections, course_name, country):
    selected_course = course_details[(course_details['Course Name'] == course_name) & 
                                     (course_details['Country'] == country)]
    
    content = f"# {course_name} in {country}\n\n"
    
    # Overview Section (Extended)
    content += f"## Overview\n"
    content += (f"The {course_name} in {country} is a highly prestigious degree that attracts students from all over the world. "
                f"Not only does it offer top-tier education in {selected_course.iloc[0]['Top Specializations']}, "
                f"but universities like {selected_course.iloc[0]['University Name']} provide world-class research facilities.\n\n")
    
    # Adding academic flexibility, and program highlights
    content += (f"This course typically spans 2 years with options to specialize in fields such as {selected_course.iloc[0]['Top Specializations']}. "
                f"Many institutions offer part-time study options or flexible course schedules to accommodate working professionals.\n\n")
    
    # Cost and Living Expenses (Extended)
    content += f"## Cost and Living Expenses\n"
    content += (f"The total cost of pursuing {course_name} in {country} includes tuition fees ranging between {selected_course.iloc[0]['Tuition Fees (INR)']} "
                f"and living expenses of approximately {selected_course.iloc[0]['Living Expenses (INR)']} annually. "
                f"Students should budget for additional expenses such as travel, accommodation, books, and supplies.\n\n")
    
    # Adding comparative analysis of universities with graphs
    content += "### Tuition Fees and ROI Comparison\n"
    content += "Here’s a comparison of the top universities offering the best return on investment (ROI):\n\n"
    
    # Adding College Collection (Best ROI)
    content += "## Best ROI Universities for MS in Computer Science\n"
    selected_collections = college_collections[college_collections['Collection Name'] == 
                                               f"Best ROI for {course_name} ({country})"]
    content += "| University Name | Tuition Fees | Acceptance Rate | Application Link |\n"
    content += "|-----------------|--------------|-----------------|------------------|\n"
    for _, row in selected_collections.iterrows():
        content += (f"| {row['University Name']} | {row['Tuition Fees (INR)']} | "
                    f"{row['Acceptance Rate (%)']}% | [Apply Here]({row['Application Link']}) |\n")
    
    # Scholarships and Financial Aid Section
    content += "\n## Scholarships and Financial Aid\n"
    content += (f"To support international students, there are several scholarships available, including {selected_course.iloc[0]['Scholarships Available']}. "
                f"Many institutions also offer merit-based scholarships, while government-backed grants are often available for students in STEM fields.\n\n")
    
    # Visa Information Section
    content += "## Visa Information\n"
    content += (f"International students pursuing {course_name} in {country} typically apply for the {selected_course.iloc[0]['Visa Types']}. "
                f"The application process requires proof of admission, financial stability, and adherence to local immigration laws. "
                f"Students should plan to apply for a visa at least 3 months before the start of their program.\n\n")
    
    # Cost Analysis and Budgeting Section
    content += "## Cost Analysis and Budgeting\n"
    content += (f"While the tuition and living expenses may seem high, the return on investment for graduates of {selected_course.iloc[0]['University Name']} "
                f"justifies the cost. Graduates in fields like {selected_course.iloc[0]['Top Specializations']} see an average starting salary of "
                f"{selected_course.iloc[0]['Median Salary (USD)']}, and with a net ROI of {selected_course.iloc[0]['Net ROI (USD)']}, this degree "
                f"is a sound investment for the future.\n\n")

    # Adding deeper analysis of job prospects and career growth
    content += "## Job Prospects and Career Growth\n"
    content += (f"Graduates from {course_name} in {country} often secure employment with global leaders like {selected_course.iloc[0]['Top Employers']}. "
                f"Their roles range from software developers to data scientists, with an average starting salary of {selected_course.iloc[0]['Median Salary (USD)']}. "
                f"Career growth in fields like {selected_course.iloc[0]['Top Specializations']} is exceptionally strong, with salaries typically increasing "
                f"by 10-15% annually for the first five years.\n\n")
    
    # Adding a YouTube guide
    content += "\n## Watch a Detailed Guide\n"
    content += f"[YouTube Video]({selected_course.iloc[0]['Relevant YouTube Video URL']})\n\n"

    return content

# Final Phase: Streamlit App with full content and additional visualizations
def main_final_phase():
    st.title("Content Automation for Course Details - Complete Version")
    st.write("Upload an Excel file containing course and university details to generate a comprehensive content template with full analysis and visualizations.")

    # File upload
    uploaded_file = st.file_uploader("Upload Excel file", type="xlsx")
    
    if uploaded_file:
        # Read the uploaded Excel file
        course_details = pd.read_excel(uploaded_file, sheet_name="Course Details")
        graph_data = pd.read_excel(uploaded_file, sheet_name="Graph Data")
        college_collections = pd.read_excel(uploaded_file, sheet_name="College Collections")
        
        # User selects the course and country
        course_name = st.selectbox("Select Course Name", course_details['Course Name'].unique())
        country = st.selectbox("Select Country", course_details['Country'].unique())
        
        # Generate the final, extended content
        generated_content = generate_final_content(course_details, graph_data, college_collections, course_name, country)
        
        # Display generated content
        st.subheader("Generated Content")
        st.text_area("Content", value=generated_content, height=500)
        
        # Visualizations: Tuition Fees by University
        selected_graph_data = graph_data[graph_data['University Name'].isin(course_details['University Name'])]
        fig = px.bar(selected_graph_data, x='University Name', y='Tuition Fees (INR)', title='Tuition Fees by University')
        st.plotly_chart(fig)

        # Visualizations: ROI by University
        fig2 = px.scatter(selected_graph_data, x='University Name', y='Net ROI (USD)', size='Tuition Fees (INR)', 
                          color='University Name', title='ROI vs Tuition Fees by University')
        st.plotly_chart(fig2)

        # Word and HTML download options
        word_file = generate_word(generated_content)
        st.markdown(download_link(word_file.read(), f"{course_name}_complete_content.docx", "Download Word Document"), unsafe_allow_html=True)
        
        html_content = generate_html(generated_content)
        st.markdown(download_link(html_content.encode(), f"{course_name}_complete_content.html", "Download HTML Document"), unsafe_allow_html=True)

if __name__ == "__main__":
    main_final_phase()
