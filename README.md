# 📦 Inventory Transfer Optimization System

A Streamlit-based retail inventory transfer recommendation generation system that intelligently analyzes inventory data and provides optimized transfer suggestions between stores.

## 🚀 Features

- **Smart Data Processing**: Automatic data validation and preprocessing
- **Dual Transfer Strategies**: 
  - Option A: Conservative Transfer (20% surplus limit)
  - Option B: Enhanced Transfer (50% surplus limit)
- **Priority-based Matching**: Intelligent matching of transfer-out and receive candidates
- **Comprehensive Analytics**: Detailed statistics and visualizations
- **Excel Export**: Multi-sheet Excel report generation
- **User-friendly Interface**: Streamlit-based web interface

## 📋 Required Data Format

### Mandatory Columns:
- `Article`: Product code (string)
- `Article Description`: Product description (string)
- `RP Type`: Replenishment type (ND/RF)
- `Site`: Store code (string)
- `OM`: Operational management unit (string)
- `MOQ`: Minimum order quantity (numeric)
- `SaSa Net Stock`: Current inventory quantity (numeric)
- `Pending Received`: In-transit order quantity (numeric)
- `Safety Stock`: Safety stock quantity (numeric)
- `Last Month Sold Qty`: Last month sales quantity (numeric)
- `MTD Sold Qty`: Month-to-date sales quantity (numeric)

## 🔧 Installation

1. Clone the repository:
```bash
git clone <repository-url>
cd inventory-transfer-system
```

2. Install dependencies:
```bash
pip install -r requirements.txt
```

3. Run the application:
```bash
streamlit run app.py
```

Or use the batch file:
```bash
run.bat
```

## 🎯 Transfer Strategies

### Option A: Conservative Transfer
- **ND Type**: Complete transfer-out of all available stock
- **RF Type**: Surplus transfer-out with 20% upper limit (minimum 2 pieces)

### Option B: Enhanced Transfer  
- **ND Type**: Complete transfer-out of all available stock
- **RF Type**: Enhanced transfer-out with 50% upper limit (minimum 2 pieces)
- **Sales-based prioritization**: Transfer from lowest sales locations first

## 📊 Output Features

- **Transfer Recommendations**: Detailed transfer suggestions with quantities
- **Statistical Analysis**: By product, by OM, transfer type distributions
- **Visualizations**: Matplotlib bar charts showing transfer vs receive analysis
- **Excel Export**: Comprehensive multi-sheet Excel report

## 🛠️ Technical Stack

- **Frontend**: Streamlit (>=1.28.0)
- **Data Processing**: pandas (>=2.0.0), numpy (>=1.24.0)
- **Excel Handling**: openpyxl (>=3.1.0)
- **Visualization**: matplotlib (>=3.7.0), seaborn (>=0.12.0)

## 📝 Version History

- **v1.7**: Added dual transfer strategies, enhanced analytics, and improved UI
- **v1.6**: Initial release with basic transfer logic

## 👨‍💻 Developer

**Ricky** - Inventory Optimization Specialist

## 📄 License

This project is proprietary software developed for internal use.