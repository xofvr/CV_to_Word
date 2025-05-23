"""
CV Generator using Office-Word-MCP-Server
Streamlit app for generating professional CV Word documents
"""
import streamlit as st
import os
import sys
from PIL import Image
import io
import asyncio
import json

# Check if we can import the MCP tools
MCP_AVAILABLE = False
try:
    # Add path to MCP server
    mcp_path = os.path.join(os.getcwd(), 'Office-Word-MCP-Server')
    if mcp_path not in sys.path:
        sys.path.insert(0, mcp_path)
    
    # Try importing the MCP tools
    import word_document_server.tools.document_tools as doc_tools
    import word_document_server.tools.content_tools as content_tools
    import word_document_server.tools.format_tools as format_tools
    
    MCP_AVAILABLE = True
    print("✅ MCP Word tools loaded successfully!")
except Exception as e:
    print(f"❌ MCP tools not available: {e}")
    print("Will use basic functionality")
    MCP_AVAILABLE = False

class CVGenerator:
    def __init__(self):
        self.output_dir = "generated_cvs"
        os.makedirs(self.output_dir, exist_ok=True)
    
    def analyze_cv_template(self, image):
        """
        Placeholder: Returns a static template structure for now.
        """
        template_structure = {
            "header": {
                "name": "{NAME_PLACEHOLDER}",
                "title": "{TITLE_PLACEHOLDER}",
                "contact": "{CONTACT_PLACEHOLDER}"
            },
            "sections": [
                {
                    "title": "Professional Summary",
                    "type": "paragraph",
                    "content": "{SUMMARY_PLACEHOLDER}"
                },
                {
                    "title": "Experience",
                    "type": "table",
                    "rows": 3,
                    "cols": 2,
                    "data": [
                        ["Position", "Company"],
                        ["{POSITION_1}", "{COMPANY_1}"],
                        ["{POSITION_2}", "{COMPANY_2}"]
                    ]
                },
                {
                    "title": "Education",
                    "type": "paragraph",
                    "content": "{EDUCATION_PLACEHOLDER}"
                },
                {
                    "title": "Skills",
                    "type": "paragraph",
                    "content": "{SKILLS_PLACEHOLDER}"
                }
            ]
        }
        return template_structure

    def __init__(self):
        self.output_dir = "generated_cvs"
        os.makedirs(self.output_dir, exist_ok=True)
        
        # Pre-defined CV templates based on common layouts
        self.templates = {
            "Modern Professional": {
                "header": {
                    "name": "{FULL_NAME}",
                    "title": "{PROFESSIONAL_TITLE}",
                    "contact": "{EMAIL} | {PHONE} | {LOCATION}"
                },
                "sections": [
                    {
                        "title": "Professional Summary",
                        "type": "paragraph",
                        "content": "{PROFESSIONAL_SUMMARY}"
                    },
                    {
                        "title": "Work Experience",
                        "type": "experience",
                        "items": [
                            {
                                "role": "{JOB_TITLE_1}",
                                "company": "{COMPANY_1}",
                                "duration": "{DURATION_1}",
                                "description": "{JOB_DESCRIPTION_1}"
                            },
                            {
                                "role": "{JOB_TITLE_2}",
                                "company": "{COMPANY_2}",
                                "duration": "{DURATION_2}",
                                "description": "{JOB_DESCRIPTION_2}"
                            }
                        ]
                    },
                    {
                        "title": "Education",
                        "type": "paragraph",
                        "content": "{EDUCATION_DETAILS}"
                    },
                    {
                        "title": "Key Skills",
                        "type": "paragraph",
                        "content": "{SKILLS_LIST}"
                    }
                ]
            },
            "Classic Format": {
                "header": {
                    "name": "{FULL_NAME}",
                    "title": "{PROFESSIONAL_TITLE}",
                    "contact": "{EMAIL} | {PHONE} | {LOCATION}"
                },
                "sections": [
                    {
                        "title": "Objective",
                        "type": "paragraph",
                        "content": "{CAREER_OBJECTIVE}"
                    },
                    {
                        "title": "Experience",
                        "type": "table",
                        "rows": 4,
                        "cols": 3,
                        "data": [
                            ["Position", "Company", "Duration"],
                            ["{POSITION_1}", "{COMPANY_1}", "{DURATION_1}"],
                            ["{POSITION_2}", "{COMPANY_2}", "{DURATION_2}"],
                            ["{POSITION_3}", "{COMPANY_3}", "{DURATION_3}"]
                        ]
                    },
                    {
                        "title": "Education & Certifications",
                        "type": "paragraph",
                        "content": "{EDUCATION_AND_CERTS}"
                    },
                    {
                        "title": "Technical Skills",
                        "type": "paragraph",
                        "content": "{TECHNICAL_SKILLS}"
                    }
                ]
            }
        }
    
    def get_template_by_name(self, template_name):
        """Get template structure by name"""
        return self.templates.get(template_name, self.templates["Modern Professional"])
    
    def analyze_cv_template(self, image=None, template_name="Modern Professional"):
        """
        Generate CV structure based on selected template
        Note: Image analysis removed - focusing on template-based approach
        """
        return self.get_template_by_name(template_name)
    
    async def _generate_word_document_async(self, template_structure, filepath):
        """
        Async function to generate Word document using MCP tools
        """
        try:
            # Create new document
            await doc_tools.create_document(
                filename=filepath,
                title="Professional CV",
                author="CV Generator"
            )
            
            # Add header section
            header = template_structure["header"]
            await content_tools.add_heading(filepath, header["name"], level=1)
            await content_tools.add_paragraph(filepath, header["title"])
            await content_tools.add_paragraph(filepath, header["contact"])
            
            # Add a line break
            await content_tools.add_paragraph(filepath, "")
            
            # Add sections
            for section in template_structure["sections"]:
                # Add section heading
                await content_tools.add_heading(filepath, section["title"], level=2)
                
                if section["type"] == "paragraph":
                    await content_tools.add_paragraph(filepath, section["content"])
                elif section["type"] == "table":
                    await content_tools.add_table(
                        filepath, 
                        rows=section["rows"], 
                        cols=section["cols"], 
                        data=section["data"]
                    )
                elif section["type"] == "experience":
                    # Handle experience items as formatted text
                    for item in section["items"]:
                        await content_tools.add_paragraph(
                            filepath, 
                            f"**{item['role']}** at {item['company']} ({item['duration']})"
                        )
                        await content_tools.add_paragraph(filepath, item['description'])
                        await content_tools.add_paragraph(filepath, "")  # Add spacing
                
                # Add spacing between sections
                await content_tools.add_paragraph(filepath, "")
            
            # Apply formatting to make it look professional
            await self._apply_formatting_async(filepath)
            
            return True
            
        except Exception as e:
            st.error(f"Error generating document: {str(e)}")
            return False
    
    async def _apply_formatting_async(self, filepath):
        """Apply formatting to make the document look professional"""
        try:
            # Create custom styles for professional look
            await format_tools.create_custom_style(
                filename=filepath,
                style_name="CVHeader",
                bold=True,
                font_size=18,
                font_name="Arial",
                color="blue"
            )
            
            await format_tools.create_custom_style(
                filename=filepath,
                style_name="SectionHeader",
                bold=True,
                font_size=14,
                font_name="Arial",
                color="darkblue"
            )
            
        except Exception as e:
            st.warning(f"Formatting warning: {str(e)}")
    
    def generate_word_document(self, template_structure, filename="cv_template.docx"):
        """
        Generate Word document using MCP tools based on analyzed template
        """
        filepath = os.path.join(self.output_dir, filename)
        
        if not MCP_AVAILABLE:
            st.error("MCP tools not available. Cannot generate Word document.")
            return None
        
        try:
            # Run the async function
            loop = asyncio.new_event_loop()
            asyncio.set_event_loop(loop)
            success = loop.run_until_complete(
                self._generate_word_document_async(template_structure, filepath)
            )
            loop.close()
            
            if success:
                return filepath
            else:
                return None
            
        except Exception as e:
            st.error(f"Error generating document: {str(e)}")
            return None

def main():
    st.set_page_config(
        page_title="CV Template to Word Generator",
        page_icon="📄",
        layout="wide"
    )
    
    st.title("CV Template to Word Generator")
    st.markdown("Upload a CV template image and generate a Word document based on the layout")
    
    # Show MCP status
    if MCP_AVAILABLE:
        st.success("✅ MCP Word tools are available!")
    else:
        st.error("❌ MCP Word tools not available. Please check your setup.")
    
    # Initialize CV generator
    cv_gen = CVGenerator()
    
    # Sidebar for template selection
    st.sidebar.header("Template Options")

    # Add tabs for input methods: image, JSON, and DOCX-to-JSON
    tabs = st.tabs(["Image Input", "JSON Input", "DOCX to JSON"])
    tab_image, tab_json, tab_docx2json = tabs

    # JSON Input tab
    with tab_json:
        st.subheader("JSON Template Input")
        json_text = st.text_area(
            "Paste CV Template JSON",
            value="",
            height=200,
            help="Paste JSON structure for CV template"
        )
        if json_text:
            try:
                template_json = json.loads(json_text)
                st.json(template_json)
                st.session_state.json_template_structure = template_json
            except Exception as e:
                st.error(f"Invalid JSON: {e}")

        if 'json_template_structure' in st.session_state:
            st.subheader("Generate Word Document from JSON")
            json_filename = st.text_input(
                "Output filename", 
                value="cv_from_json.docx",
                help="Name for the generated Word document from JSON"
            )
            if st.button("Generate Word Document from JSON"):
                if not MCP_AVAILABLE:
                    st.error("Cannot generate document - MCP tools not available")
                else:
                    with st.spinner("Generating Word document..."):
                        filepath = cv_gen.generate_word_document(
                            st.session_state.json_template_structure,
                            json_filename
                        )
                        if filepath and os.path.exists(filepath):
                            st.success(f"Document generated successfully: {filepath}")
                            with open(filepath, "rb") as file:
                                st.download_button(
                                    label="Download Word Document",
                                    data=file.read(),
                                    file_name=json_filename,
                                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                                )
                        else:
                            st.error("Failed to generate document")

    # Image Input tab
    with tab_image:
        # File uploader
        uploaded_file = st.file_uploader(
            "Upload CV Template Image",
            type=["png", "jpg", "jpeg"],
            help="Upload an A4 CV template image"
        )
        
        if uploaded_file is not None:
            # Display uploaded image
            image = Image.open(uploaded_file)
            
            col1, col2 = st.columns(2)
            
            with col1:
                st.subheader("Uploaded Template")
                st.image(image, caption="CV Template", use_column_width=True)
            
            with col2:
                st.subheader("Template Analysis")
                
                if st.button("Analyze Template", type="primary"):
                    with st.spinner("Analyzing template layout..."):
                        # Analyze the template
                        structure = cv_gen.analyze_cv_template(image)
                        
                        # Display structure
                        st.json(structure)
                        
                        # Store in session state
                        st.session_state.template_structure = structure
            
            # Generate document section
            if 'template_structure' in st.session_state:
                st.subheader("Generate Word Document")
                
                filename = st.text_input(
                    "Output filename", 
                    value="cv_template.docx",
                    help="Name for the generated Word document"
                )
                
                if st.button("Generate Word Document", type="primary"):
                    if not MCP_AVAILABLE:
                        st.error("Cannot generate document - MCP tools not available")
                    else:
                        with st.spinner("Generating Word document..."):
                            filepath = cv_gen.generate_word_document(
                                st.session_state.template_structure,
                                filename
                            )
                            
                            if filepath and os.path.exists(filepath):
                                st.success(f"Document generated successfully: {filepath}")
                                
                                # Provide download button
                                with open(filepath, "rb") as file:
                                    st.download_button(
                                        label="Download Word Document",
                                        data=file.read(),
                                        file_name=filename,
                                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                                    )
                                
                                # Preview document info
                                try:
                                    async def get_preview():
                                        doc_info = await doc_tools.get_document_info(filepath)
                                        doc_text = await doc_tools.get_document_text(filepath)
                                        return doc_info, doc_text
                                    
                                    loop = asyncio.new_event_loop()
                                    asyncio.set_event_loop(loop)
                                    doc_info, doc_text = loop.run_until_complete(get_preview())
                                    loop.close()
                                    
                                    st.subheader("Document Preview")
                                    st.json(doc_info)
                                    
                                    # Show document text
                                    st.text_area("Document Content", doc_text, height=200)
                                    
                                except Exception as e:
                                    st.warning(f"Preview error: {str(e)}")
                            else:
                                st.error("Failed to generate document")

    # DOCX to JSON tab
    with tab_docx2json:
        st.subheader("DOCX to JSON Converter")
        uploaded_docx = st.file_uploader(
            "Upload a Word (.docx) file to extract as JSON",
            type=["docx"],
            help="Upload a .docx CV file to extract its structure as JSON"
        )
        if uploaded_docx is not None:
            # Save uploaded file to a temp location
            temp_docx_path = os.path.join("generated_cvs", uploaded_docx.name)
            with open(temp_docx_path, "wb") as f:
                f.write(uploaded_docx.read())
            try:
                if not MCP_AVAILABLE:
                    st.error("MCP tools not available. Cannot extract JSON.")
                else:
                    # Use MCP tools to extract document info and text
                    import asyncio
                    async def extract_docx_json():
                        doc_info = await doc_tools.get_document_info(temp_docx_path)
                        doc_text = await doc_tools.get_document_text(temp_docx_path)
                        return {"info": doc_info, "text": doc_text}
                    loop = asyncio.new_event_loop()
                    asyncio.set_event_loop(loop)
                    docx_json = loop.run_until_complete(extract_docx_json())
                    loop.close()
                    st.subheader("Extracted JSON Structure")
                    st.json(docx_json)
                    st.download_button(
                        label="Download JSON",
                        data=json.dumps(docx_json, indent=2),
                        file_name=uploaded_docx.name.replace(".docx", ".json"),
                        mime="application/json"
                    )
            except Exception as e:
                st.error(f"Error extracting JSON: {e}")

if __name__ == "__main__":
    main()
