# webui.py

import streamlit as st
import tempfile
import os
#from MarkdownToPPTX import MarkdownToPPTX
from MarkdownToPPTX.MarkdownToPPTX import MarkdownToPPTX

# Page configuration
st.set_page_config(
    page_title="mykLabs Markdown to PowerPoint Converter",
    #page_icon="📝",
    page_icon="./assets/images/k600.ico",
    layout="wide"
)

# Set the server to only listen for local loopback addresses 
# and allow local access only
st.config.set_option('server.address', '127.0.0.1')

# set page icon and logo
st.logo(image="./assets/images/k600.svg", size="large")

#st.markdown("---")
# Page header
st.title("Markdown to PowerPoint Converter by mykLabs")
#st.markdown("---")

template_file = st.file_uploader(
    "Choose a PowerPoint template file (.pptx)",
    type=["pptx"],
    key="template_uploader"
)

# Save template to temporary file if uploaded
template_path = None
if template_file is not None:
    try:
        with tempfile.NamedTemporaryFile(delete=False, suffix=".pptx") as tmp_template:
            tmp_template.write(template_file.read())
            template_path = tmp_template.name
        st.success("Template uploaded successfully!")
    except Exception as e:
        st.error(f"Error uploading template: {str(e)}")
        template_path = None


# Create tabs
tab1, tab2 = st.tabs(["📝 Text Input", "📁 File Upload"])

# Tab 1: Text input
with tab1:
    st.header("Enter Markdown Text")
    markdown_text = st.text_area(
        "Paste your markdown content here:",
        height=300,
        placeholder="# Your Markdown Content\n\n---\n\n## Slide Title\n\n- Bullet point 1\n- Bullet point 2\n\n---\n\n## Another Slide\n\n| Column 1 | Column 2 |\n|----------|----------|\n| Data 1   | Data 2   |"
    )
    
    if st.button("Convert to PowerPoint", key="text_convert"):
        if markdown_text.strip():
            try:
                # Create converter instance with template if provided
                converter = MarkdownToPPTX(template_path) if template_path else MarkdownToPPTX()
                
                # Parse markdown content
                slides_data = converter.parse_markdown(markdown_text)
                
                if slides_data:
                    # Create a new presentation instead of clearing slides
                    # Create slides
                    first_header_slide = True
                    slide_count = 0
                    for i, slide_data in enumerate(slides_data):
                        if first_header_slide and slides_data[0]['title']:
                            # Create title slide for the first main header
                            converter.create_title_slide(slide_data['title'])
                            first_header_slide = False
                            slide_count += 1
                            
                            # If this slide has content, create a content slide too
                            if slide_data['content']:
                                converter.create_content_slide(slide_data['title'], slide_data['content'])
                                slide_count += 1
                        else:
                            converter.create_content_slide(slide_data['title'], slide_data['content'])
                            slide_count += 1
                    
                    # Save to temporary file
                    with tempfile.NamedTemporaryFile(delete=False, suffix=".pptx") as tmp_file:
                        converter.presentation.save(tmp_file.name)
                        tmp_file_path = tmp_file.name
                    
                    # Provide download button
                    with open(tmp_file_path, "rb") as file:
                        st.download_button(
                            label="Download PowerPoint Presentation",
                            data=file,
                            file_name="presentation.pptx",
                            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                        )
                    
                    # Clean up temporary file
                    os.unlink(tmp_file_path)

                    # Clean up template file if it was uploaded
                    if template_path and os.path.exists(template_path):
                        os.unlink(template_path)
                    
                    st.success("Conversion successful! Click the download button to get your PowerPoint file.")
                else:
                    st.warning("No valid slide data found in the markdown content.")
            except Exception as e:
                # Clean up template file if it was uploaded
                if template_path and os.path.exists(template_path):
                    os.unlink(template_path)
                st.error(f"An error occurred during conversion: {str(e)}")
        else:
            st.warning("Please enter some markdown content.")

# Tab 2: File upload
with tab2:
    st.header("Upload Markdown File")
    uploaded_file = st.file_uploader(
        "Choose a markdown file",
        type=["md", "markdown"],
        key="file_uploader"
    )
    
    if uploaded_file is not None:
        try:
            # Read the file content
            markdown_content = uploaded_file.read().decode("utf-8")
            st.text_area(
                "File content preview:",
                value=markdown_content[:500] + ("..." if len(markdown_content) > 500 else ""),
                height=200,
                disabled=True
            )
            
            if st.button("Convert to PowerPoint", key="file_convert"):
                try:
                    # Create converter instance with template if provided
                    converter = MarkdownToPPTX(template_path) if template_path else MarkdownToPPTX()
                    
                    # Parse markdown content
                    slides_data = converter.parse_markdown(markdown_content)
                    
                    if slides_data:
                        # Create slides
                        first_header_slide = True
                        slide_count = 0
                        for i, slide_data in enumerate(slides_data):
                            if first_header_slide and slides_data[0]['title']:
                                # Create title slide for the first main header
                                converter.create_title_slide(slide_data['title'])
                                first_header_slide = False
                                slide_count += 1
                                
                                # If this slide has content, create a content slide too
                                if slide_data['content']:
                                    converter.create_content_slide(slide_data['title'], slide_data['content'])
                                    slide_count += 1
                            else:
                                converter.create_content_slide(slide_data['title'], slide_data['content'])
                                slide_count += 1
                        
                        # Save to temporary file
                        with tempfile.NamedTemporaryFile(delete=False, suffix=".pptx") as tmp_file:
                            converter.presentation.save(tmp_file.name)
                            tmp_file_path = tmp_file.name
                        
                        # Provide download button
                        with open(tmp_file_path, "rb") as file:
                            st.download_button(
                                label="Download PowerPoint Presentation",
                                data=file,
                                file_name="presentation.pptx",
                                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                            )
                        
                        # Clean up temporary file
                        os.unlink(tmp_file_path)

                        # Clean up template file if it was uploaded
                        if template_path and os.path.exists(template_path):
                            os.unlink(template_path)
                        
                        st.success("Conversion successful! Click the download button to get your PowerPoint file.")
                    else:
                        st.warning("No valid slide data found in the markdown file.")
                except Exception as e:
                    # Clean up template file if it was uploaded
                    if template_path and os.path.exists(template_path):
                        os.unlink(template_path)
                    st.error(f"An error occurred during conversion: {str(e)}")
        except Exception as e:
            # Clean up template file if it was uploaded
            if template_path and os.path.exists(template_path):
                os.unlink(template_path)
            st.error(f"Error reading file: {str(e)}")
    else:
        st.info("Please upload a markdown file.")

# Usage instructions
st.markdown("---")
st.header("Usage Instructions")
st.markdown("""
1. **PowerPoint Template (Optional)**:
   - Upload a .pptx template file to use as the base for your presentation
   - If no template is provided, a default presentation will be created
            
2. **Text Input Tab**:
   - Paste your markdown content directly into the text area
   - Click "Convert to PowerPoint" to generate your presentation
   - Download the generated .pptx file

3. **File Upload Tab**:
   - Upload a .md or .markdown file from your computer
   - Review the content preview
   - Click "Convert to PowerPoint" to generate your presentation
   - Download the generated .pptx file

**Markdown Formatting Supported**:
- Slide separators: `---`
- Headers: `#` for title slides, `##` for section headers, `###` etc. for content headers
- Bullet points: `-` or `*` with indentation support
- Bold text: `**bold text**`
- Tables: Using `|` and `:---` format
- Regular paragraphs

**Example**:
```markdown
# Presentation Title

---

## First Slide

- Bullet point 1
- Bullet point 2
  - Indented bullet
- **Bold text** example

---

## Data Table

| Column 1 | Column 2 | Column 3 |
| :--- | :--- | :--- |
| Data 1 | Data 2 | Data 3 |
| Data 4 | Data 5 | Data 6 |
            """)

st.markdown("---")
st.markdown("""
##### by mykLabs
[myklabs](mailto:mikkel03@gmail.com) is a developer community focused on building and sharing tools for developers. We are a group of developers who are passionate about building tools that make life easier.
""")
st.markdown("""
 Developer Contact: If you have any questions or feedback, please feel free to contact me at [github](https://github.com/myklabs).
""")
