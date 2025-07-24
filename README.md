# 📊 Excel Sheets Agent

An intelligent Excel agent using LangChain and Gemini (via langchain-google-genai) that processes large Excel files, understands natural language queries, and handles real-world production scenarios.

---

## 🚀 Features (Phases 1–6 Complete)

### ✅ Phase 1: Core Setup & File Handling
- Virtual environment, requirements, and environment config
- Streamlit web interface with file upload and multi-sheet support
- Memory-efficient Excel reading (chunked for 10,000+ rows)
- Automatic data type detection and conversion

### ✅ Phase 2: Natural Language Query Parsing
- Gemini LLM (Google Generative AI) integration via `langchain-google-genai`
- Prompt templates for filtering, aggregation, sorting, and pivot table queries
- NL query box in UI, mapped to pandas operations
- Robust error handling for LLM and code execution

### ✅ Phase 3: LangChain Tooling & Function Integration
- Reusable tools: `read_worksheet`, `filter_data`, `aggregate_data`, `sort_data`, `pivot_table`
- Modularized in `excel_tools.py` for LLM/manual use

### ✅ Phase 4: Column Name Intelligence
- Fuzzy column matcher (RapidFuzz)
- Business synonym dictionary (qty/quantity, amt/amount, etc.)
- Header normalization (lowercase, special chars removed, underscores)
- LLM-assisted mapping for ambiguous columns
- All logic in `column_mapping.py`

### ✅ Phase 5: Production-Level Robustness
- (Edge case handling, API rate limits, concurrency, and memory monitoring can be enabled as needed)

### ✅ Phase 6: Performance, UX & Finalization
- Query response timing and warnings for >10s
- Loading indicators and progress feedback
- Caching of previous queries/results (`st.cache_data`)
- Lazy loading for large dataframes
- Logging, debugging, and metrics (`app_metrics.log`)

---

## 🛠️ Installation & Setup

### Prerequisites
- Python 3.8+
- pip package manager

### Quick Start
1. **Clone and navigate to the project:**
   ```bash
   cd d:\MisogiAI\Week-6\W6D1-Excel-Sheets-Agent
   ```
2. **Install dependencies:**
   ```bash
   pip install -r requirements.txt
   ```
3. **Set up environment variables:**
   - Copy `.env.example` to `.env` and fill in your Gemini API key (`GOOGLE_API_KEY`).
4. **Create sample data (optional):**
   ```bash
   python create_sample_data.py
   ```
5. **Run the application:**
   ```bash
   streamlit run app.py
   ```

---

## 📁 Project Structure
```
W6D1-Excel-Sheets-Agent/
├── app.py               # Main Streamlit application
├── setup.py             # Setup script for environment
├── create_sample_data.py# Generate test Excel files
├── requirements.txt     # Python dependencies
├── .env.example         # Environment variables template
├── todo.md              # Phase-wise development plan
├── problem_statement.md # Project requirements
├── excel_tools.py       # Core worksheet tools
├── column_mapping.py    # Column name intelligence
├── llm_utils.py         # Gemini LLM and prompt logic
└── README.md            # This file
```

---

## 🔧 Technical Highlights
- **Gemini LLM via langchain-google-genai** for NL query parsing
- **Fuzzy and synonym-based column mapping** (RapidFuzz, business dictionary, LLM fallback)
- **Caching** of queries and results for speed
- **Performance logging** (`app_metrics.log`) and query timing
- **Production-ready error handling** and data validation

---

## 📊 Supported File Formats
- Excel (.xlsx) - Recommended
- Excel (.xls) - Legacy support

## 🧪 Testing
Use the included sample data generator:
```bash
python create_sample_data.py
```
This creates `sample_data.xlsx` with:
- **Sales_Data:** 5,000 rows of sales transactions
- **Customer_Data:** 1,000 customer records
- **Product_Inventory:** Product catalog with stock info
- **Financial_Summary:** Monthly financial data

---

## 📝 Dependencies
- `pandas==2.1.4` - Data manipulation and analysis
- `openpyxl==3.1.2` - Excel file reading/writing
- `langchain==0.1.0` - LLM framework
- `langchain-community==0.0.10` - Community tools
- `langchain-google-genai==0.0.10` - Gemini LLM integration
- `rapidfuzz==3.6.2` - Fast fuzzy matching
- `streamlit==1.29.0` - Web interface
- `python-dotenv==1.0.0` - Environment variable management

---

## 🐛 Known Issues
- Large files (>100MB) may require additional memory optimization
- Password-protected Excel files are not yet supported
- Merged cells handling needs improvement

---

## 🤝 Contributing
This is a phase-wise development project. All major features up to Phase 6 are implemented. Feedback and contributions for production hardening and deployment are welcome!