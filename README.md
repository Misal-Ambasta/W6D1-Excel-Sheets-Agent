# 📊 Excel Sheets Agent

An intelligent Excel agent using LangChain and Gemini (via langchain-google-genai) that processes large Excel files, understands natural language queries, and handles real-world production scenarios.

---

## 🛠️ Installation & Setup

### Prerequisites
- Python 3.8+
- pip package manager

### Quick Start
1. **Clone and navigate to the project:**
   ```bash
   cd cloned-repo
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

## 🐛 Known Issues
- Large files (>100MB) may require additional memory optimization
- Password-protected Excel files are not yet supported
- Merged cells handling needs improvement

---
