# PPTX Translation Service

A web application that translates PowerPoint presentations from English to Arabic. Upload a PPTX file, and get an Excel file with all extracted text and translations.

## Features

- **In-Place PPTX Translation**: Upload PowerPoint files and receive translated PPTX files with Arabic text
- **RTL Layout Mirroring**: Automatically mirrors slide layouts from left-to-right to right-to-left for Arabic
- **Flexible Output**: Choose between translated PPTX, Excel, or both formats
- **Slide Range Selection**: Translate specific slides (e.g., "1-10", "1,3,5", "1-5,8,10-12")
- **Semantic Dictionary**: Uses an LLM to find similar translations from a dictionary for better context
- **Dictionary Builder**: Auto-build translation dictionaries from parallel English/Arabic PPTX files with smart heuristics
- **LLM Validation**: Validates translation pairs using AI to ensure accuracy
- **Caching**: In-memory caching to avoid duplicate API calls

## Project Structure

```
Translation_site/
├── app/
│   ├── __init__.py
│   ├── main.py                 # FastAPI application
│   ├── services/
│   │   ├── __init__.py
│   │   ├── pptx_parser.py      # PPTX text extraction
│   │   ├── pptx_translator.py  # In-place translation & RTL mirroring
│   │   ├── translator.py       # Translation with LLM API
│   │   ├── excel_writer.py     # Excel file generation
│   │   ├── dictionary.py       # Dictionary management
│   │   └── alignment.py        # Parallel PPTX alignment
│   └── static/
│       └── index.html          # Web UI
├── data/
│   └── dictionary.json         # Translation dictionary
├── uploads/                    # Temporary upload storage
├── outputs/                    # Generated output files
├── requirements.txt
└── README.md
```

## Installation

1. **Clone the repository**
   ```bash
   git clone https://github.com/Abdulrahman-im/Translation.git
   cd Translation
   ```

2. **Install dependencies**
   ```bash
   pip install -r requirements.txt
   ```

3. **Configure the API**

   Edit `app/services/translator.py` and update the API configuration (lines 9-10):
   ```python
   API_URL = "https://your-api-url"  # Your LLM API endpoint
   API_KEY = "your-api-key"          # Your API key
   ```

## Running the Application

Start the server:
```bash
uvicorn app.main:app --reload
```

Open your browser and navigate to:
```
http://localhost:8000
```

## Usage

### Translate PPTX

1. Go to the **"Translate"** tab
2. Drag and drop a PowerPoint file (or click to browse)
3. Configure options:
   - **Slide Range**: Specify which slides to translate (e.g., "1-10", leave empty for all)
   - **Output Format**: Choose Translated PPTX, Excel Only, or Both
   - **Mirror Layout**: Enable/disable RTL layout mirroring for Arabic
4. Click **"Upload & Translate"**
5. Download your translated files

### Build Dictionary

1. Go to the **"Build Dictionary"** tab
2. Upload an English PPTX file and its Arabic counterpart
3. Click **"Build Dictionary"**
4. The system will:
   - Extract text from both files
   - Align texts by slide number
   - Validate each pair using the LLM
   - Add validated pairs to the dictionary

## API Endpoints

| Endpoint | Method | Description |
|----------|--------|-------------|
| `/api/upload` | POST | Upload PPTX and get translations |
| `/api/download/{file_id}` | GET | Download generated Excel file |
| `/api/dictionary` | GET | Get all dictionary entries |
| `/api/dictionary/add` | POST | Add a single dictionary entry |
| `/api/dictionary/build` | POST | Build dictionary from parallel PPTXs |
| `/api/dictionary/stats` | GET | Get dictionary statistics |
| `/api/health` | GET | Health check |

## How Translation Works

1. **Exact Match**: Check if the text exists in the dictionary
2. **Cache Check**: Return cached translation if available
3. **Semantic Search**: Find similar entries in the dictionary using LLM
4. **Translation**: Call LLM API with similar translations as context
5. **Cache Result**: Store translation for future use

## Dependencies

- FastAPI - Web framework
- Uvicorn - ASGI server
- python-pptx - PowerPoint file processing
- openpyxl - Excel file generation
- requests - HTTP client for API calls

## License

MIT
