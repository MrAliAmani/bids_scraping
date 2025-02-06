# 🚀 Multi-Source Bid Scraping Project

Welcome to our advanced bid scraping project! This system is designed to extract bid information from multiple government and procurement websites, providing a comprehensive dataset of current bidding opportunities.

## 🌟 Features

- 🕷️ Automated web scraping from multiple procurement sources
- 🤖 Intelligent data extraction and parsing
- 📅 Date-based filtering for recent bids (usually within 1-2 days)
- 📊 Standardized data output in Excel format
- 📁 Organized file structure for downloaded attachments
- 🔐 Secure handling of login credentials
- 🔄 Robust error handling and retry mechanisms
- 📤 Automated file upload to MinIO/S3 storage
- 🎯 Parallel script execution with status monitoring

## 🧠 What This Project Does

This project automates the process of gathering bid information from various government and procurement websites. It performs the following tasks:

1. Executes multiple scrapers in parallel (up to 4 concurrent scripts)
2. Each scraper:
   - Navigates to its target website
   - Logs in (if required)
   - Applies necessary filters to find recent bids
   - Extracts detailed bid information
   - Downloads associated attachments
   - Organizes data into standardized Excel format
3. Automatically uploads completed data to MinIO/S3 storage
4. Provides real-time status monitoring and logging

## 🛠️ Technology Stack

- **Language**: Python 3.8+
- **Web Scraping**: 
  - Selenium WebDriver
  - BeautifulSoup4
  - Selenium Stealth
- **Data Processing**: 
  - Pandas
  - openpyxl
- **Storage**: 
  - MinIO (development)
  - AWS S3 (production)
- **Automation**:
  - Threading for parallel execution
  - Boto3 for S3 operations
- **Monitoring**:
  - Real-time status dashboard
  - Comprehensive logging

## 🚀 Getting Started

### Prerequisites

- Python 3.8+
- Chrome browser
- ChromeDriver (matching your Chrome version)
- MinIO server (for local testing)
- AWS S3 access (for production)

### Installation

1. Clone the repository:
   ```bash
   git clone https://github.com/yourusername/bid-scraping-project.git
   cd bid-scraping-project
   ```

2. Set up a virtual environment:
   ```bash
   python -m venv venv
   source venv/bin/activate  # On Windows, use `venv\Scripts\activate`
   ```

3. Install required packages:
   ```bash
   pip install -r requirements.txt
   ```

4. Set up environment variables:
   Create a `.env` file with required credentials:
   ```bash
   # Website Credentials
   HARTFORD_EMAIL=your_email@example.com
   HARTFORD_PASSWORD=your_password
   NYC_EMAIL=your_email@example.com
   NYC_PASSWORD=your_password
   
   # MinIO/S3 Configuration
   MINIO_ACCESS_KEY=minioadmin
   MINIO_SECRET_KEY=minioadmin
   MINIO_ENDPOINT=http://localhost:9000
   MINIO_BUCKET_NAME=bids-data
   ```

### Running the System

#### Single Scraper Mode
To run a specific scraper:
```bash
python scrapers/[scraper_name].py --days 2
```

#### Master Script Mode
To run all scrapers with parallel execution:
```bash
python master_script_v3.py --days 2
```

The master script provides:
- Concurrent execution (up to 4 scrapers)
- Real-time status dashboard
- Automatic upload to MinIO/S3
- Error recovery and retry mechanisms

## 📊 Output Structure

```
YYYY-MM-DD/                     # Date folder
├── [scraper_name]_COMPLETED/   # Scraper-specific folder
│   ├── [scraper_name].xlsx    # Consolidated bid data
│   └── [bid_number]/          # Bid-specific folders
│       └── attachments/       # Downloaded bid documents
```

## 🔄 Upload System

The project includes an automated upload system (`upload_bids.py`) that:
- Monitors for completed scraper folders
- Removes empty folders automatically
- Uploads data to MinIO (development) or S3 (production)
- Supports versioning and error recovery
- Provides detailed upload status reporting

## 🚧 Error Handling

- Comprehensive error recovery in scrapers
- Automatic retry mechanisms
- Session management
- Detailed logging
- CAPTCHA detection and handling
- Network error recovery

## 📝 Logging

- Real-time status dashboard
- Detailed execution logs
- Error tracking and reporting
- Upload status monitoring
- Performance metrics

## 🔒 Security

- Secure credential management
- SSL/TLS encryption
- Anti-bot detection measures
- Secure file handling
- Role-based access control

## 💡 Best Practices

1. Development
   - Use virtual environments
   - Test with MinIO locally
   - Follow PEP 8 guidelines
   - Document code changes

2. Deployment
   - Update environment variables
   - Verify storage configuration
   - Check folder permissions
   - Monitor initial execution

## 🤝 Contributing

We welcome contributions! Please see our [Contributing Guidelines](CONTRIBUTING.md) for more details.

## 📄 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

Happy scraping! 🎉

## 🧮 Data Processing Methods

The project uses advanced AI and ML techniques for accurate data matching and classification:

### 🎯 Matching Methods

1. **Agency Matching**
   - Batched similarity calculations
   - URL analysis
   - Enhanced confidence scoring
   - Pattern recognition

2. **Notice Type Matching**
   - Pattern-based recognition
   - Fallback to common types
   - Better variation handling
   - Context-aware classification

3. **Category Matching**
   - Domain-specific pattern recognition
   - Weighted semantic similarity
   - Code and keyword matching
   - Confidence thresholds
   - Ensemble method for improved accuracy

### 🔄 Processing Enhancements

1. **Advanced Preprocessing**
   - Stop word removal
   - Stemming and lemmatization
   - Special character handling
   - Advanced text cleaning
   - Punctuation normalization

2. **Embedding Models**
   - BERT base model
   - RoBERTa base model
   - Sentence Transformers
   - Domain-specific fine-tuning

3. **Contextual Matching**
   - Full document context analysis
   - Transformer-based models
   - Section-aware processing
   - Hierarchical matching

### ⚡ Performance Optimizations

- Parallel chunk processing
- Vectorized DataFrame operations
- Efficient API response caching
- Embedding result caching
- Reduced memory usage
- Minimized I/O operations

## 🛠️ Using excel_processor.py

The script supports various execution modes and configurations:

### Basic Usage

```bash
# Process all COMPLETED folders (default)
python utils/excel_processor.py

# Process specific folder
python utils/excel_processor.py [optional_base_path]
```

### API Options

```bash
# Use test API
python utils/excel_processor.py --test-api

# Test API with specific folder
python utils/excel_processor.py --test-api [optional_base_path]
```

### Matching Methods

```bash
# Use ensemble method
python excel_processor.py --method ensemble

# Use hybrid method
python excel_processor.py --method hybrid

# Use contextual method
python excel_processor.py --method contextual
```

### Embedding Models

```bash
# Use RoBERTa
python excel_processor.py --embedding-model roberta-base

# Use BERT
python excel_processor.py --embedding-model bert-base-uncased
```

### Combined Options

```bash
# Multiple options
python excel_processor.py --method ensemble --test-api --base-path /path/to/files
```

### Available Arguments

- `--method`: Choose matching method (similarity, ensemble, hybrid, contextual)
- `--test-api`: Use test API endpoints
- `--base-path`: Specify folder path to process
- `--embedding-model`: Choose embedding model (paraphrase-MiniLM-L6-v2, roberta-base, bert-base-uncased)
