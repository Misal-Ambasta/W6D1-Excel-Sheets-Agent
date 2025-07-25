# Q: 2

## Excel Sheets Agent

Build an intelligent Excel agent using LangChain that processes large Excel files, understands natural language queries, and handles production scenarios including inconsistent column naming and edge cases.

## Requirements

### 1. Large File & Multi-Tab Handling

- Support 10,000+ rows and multiple worksheets
- Implement memory-efficient chunking strategies
- Handle different data types and worksheet navigation

### 2. Natural Language Processing

- Integrate LLM to interpret user queries
- Generate data operations from natural language
- Support complex analysis (filtering, aggregations, pivoting)

**Query Examples:**
- "Show sales data for Q3 2024 where revenue > 50000"
- "Create pivot table showing total sales by region and product"
- "Find customers who haven't ordered in 6 months"

### 3. Technology Stack (Choose One)

**Python:** pandas + openpyxl + langchain-community + streamlit

### 4. LangChain Tools

**Core Tools:**
- `read_worksheet()`, `filter_data()`, `aggregate_data()`, `sort_data()`, `pivot_table()`, `write_results()`

**Advanced Tools:**
- `merge_worksheets()`, `data_validation()`, `formula_evaluation()`, `chart_generation()`

### 5. LLM Integration

- **Choose:** OpenAI GPT-4, Anthropic Claude, Google Gemini Pro, or Ollama
- Implement prompt engineering, error handling, and context management

### 6. Column Name Mapping

**Handle Variations:**
- Different naming conventions (snake_case, camelCase, "Proper Case")
- Synonyms ("qty" vs "quantity", "amt" vs "amount")
- Special characters and multilingual headers

**Implementation:**
- Fuzzy matching algorithm for column names
- Synonym dictionary for business terms
- LLM-assisted column mapping suggestions

### 7. Production Edge Cases

**File Issues:** Corrupted files, password protection, merged cells, memory limits

**Data Issues:** Empty sheets, inconsistent data types, missing values, date format inconsistencies

**User Input:** Ambiguous queries, non-existent columns, conflicting conditions

**System Issues:** API rate limits, memory exhaustion, concurrent access, file locking

## Technical Specs

- Handle files up to 100MB
- Process queries within 10 seconds
- Support concurrent users
- Follow coding standards
- Implement security measures (input validation, file scanning)
