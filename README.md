# PPTX Translation Service

A web application that translates PowerPoint presentations from English to Arabic. Upload a PPTX file, and get an Excel file with all extracted text and translations.

## Features

- **PPTX Translation**: Upload PowerPoint files and receive translations in Excel format
- **Semantic Dictionary**: Uses an LLM to find similar translations from a dictionary for better context
- **Dictionary Builder**: Auto-build translation dictionaries from parallel English/Arabic PPTX files
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
│   │   ├── translator.py       # Translation with LLM API
│   │   ├── excel_writer.py     # Excel file generation
│   │   ├── dictionary.py       # Dictionary management
│   │   └── alignment.py        # Parallel PPTX alignment
│   └── static/
│       └── index.html          # Web UI
├── data/
│   └── dictionary.json         # Translation dictionary
├── uploads/                    # Temporary upload storage
├── outputs/                    # Generated Excel files
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

1. Go to the **"Translate PPTX"** tab
2. Drag and drop a PowerPoint file (or click to browse)
3. Click **"Upload & Translate"**
4. Download the Excel file with translations

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
