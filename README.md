# CV Template to Word Generator

**Work In Progress**

A powerful Streamlit application that converts CV template images into editable Word documents using the Model Context Protocol (MCP).

## 🚀 Features

- **📷 Image Input**: Upload A4 CV template images (PNG, JPG, JPEG)
- **📝 JSON Input**: Paste or edit CV structure as JSON
- **📄 DOCX to JSON**: Extract structure from existing Word documents
- **🔧 MCP Integration**: Uses Office-Word-MCP-Server for document generation
- **⬬ Download**: Generate and download professional .docx files
- **👀 Preview**: Real-time document preview before download
- **🎨 Professional Formatting**: Automatic styling and layout

## 🛠️ Installation

### Prerequisites
- Python 3.8+
- pip

### Setup Steps

1. **Clone the repository**
```bash
git clone <your-repo-url>
cd CV_to_Word
```

2. **Create and activate virtual environment**
```bash
python -m venv cv_env
source cv_env/bin/activate  # On Windows: cv_env\Scripts\activate
```

3. **Install dependencies**
```bash
pip install -r requirements.txt
```

4. **Clone MCP Server (if not included)**
```bash
git clone https://github.com/GongRzhe/Office-Word-MCP-Server.git
cd Office-Word-MCP-Server
pip install -r requirements.txt
cd ..
```

## 🎯 Usage

### Running the Application
```bash
streamlit run cv_generator.py
```

The app will open in your browser at `http://localhost:8501`

### Using the Interface

#### 📷 Image Input Tab
1. Upload an A4 CV template image
2. Click "Analyze Template" to extract structure
3. Enter filename for output
4. Click "Generate Word Document"
5. Download the generated .docx file

#### 📝 JSON Input Tab
1. Paste your CV structure as JSON in the text area
2. Preview the JSON structure
3. Enter filename for output
4. Click "Generate Word Document from JSON"
5. Download the generated .docx file

#### 📄 DOCX to JSON Tab
1. Upload an existing .docx CV file
2. View the extracted JSON structure
3. Download the JSON for future use

### JSON Structure Example
```json
{
  "header": {
    "name": "John Doe",
    "title": "Software Engineer",
    "contact": "john@email.com | +1234567890 | New York, NY"
  },
  "sections": [
    {
      "title": "Professional Summary",
      "type": "paragraph",
      "content": "Experienced software engineer with 5+ years..."
    },
    {
      "title": "Work Experience",
      "type": "experience",
      "items": [
        {
          "role": "Senior Developer",
          "company": "Tech Corp",
          "duration": "2020-2024",
          "description": "Led development of..."
        }
      ]
    }
  ]
}
```

## 📁 Project Structure

```
CV_to_Word/
├── cv_generator.py              # Main Streamlit application
├── requirements.txt             # Python dependencies
├── README.md                   # This file
├── .gitignore                  # Git ignore rules
├── Office-Word-MCP-Server/     # MCP server (ignored in git)
├── cv_env/                     # Virtual environment (ignored)
└── generated_cvs/              # Output directory (ignored)
```

## 🔧 Dependencies

### Main Dependencies
- `streamlit` - Web interface
- `Pillow` - Image processing
- `python-docx` - Word document handling (via MCP)

### MCP Server Dependencies
- `fastmcp` - MCP framework
- `python-docx` - Word document manipulation
- `asyncio` - Async operations

## 🚨 Troubleshooting

### MCP Tools Not Available
If you see "❌ MCP Word tools not available":

1. Ensure the `Office-Word-MCP-Server` directory exists
2. Install MCP server dependencies:
   ```bash
   cd Office-Word-MCP-Server
   pip install -r requirements.txt
   ```
3. Restart the Streamlit application

### Import Errors
If you encounter import errors:
```bash
pip install --upgrade streamlit pillow python-docx
```

### File Path Issues
Ensure you're running the app from the project root directory:
```bash
cd CV_to_Word
streamlit run cv_generator.py
```

## 🤝 Contributing

1. Fork the repository
2. Create a feature branch: `git checkout -b feature-name`
3. Make your changes
4. Commit: `git commit -am 'Add new feature'`
5. Push: `git push origin feature-name`
6. Submit a pull request

## 📝 License

This project is licensed under the MIT License - see the LICENSE file for details.

## 🙏 Acknowledgments

- [Office-Word-MCP-Server](https://github.com/GongRzhe/Office-Word-MCP-Server) - MCP server for Word document operations
- [Streamlit](https://streamlit.io/) - Web framework
- [Model Context Protocol](https://modelcontextprotocol.io/) - Protocol for AI tool integration

## 🔗 Links

- [MCP Documentation](https://modelcontextprotocol.io/docs)
- [Streamlit Documentation](https://docs.streamlit.io/)
- [Office-Word-MCP-Server GitHub](https://github.com/GongRzhe/Office-Word-MCP-Server)
