# 📊 Python + Excel Data Analysis Project  
### Household Expense Tracker  

This project demonstrates **data analysis and visualization** using **Python** and **Excel**.  
It is a simple side project to practice handling real-world style datasets, cleaning data, performing analysis, and visualizing results using charts.  
A **Streamlit GUI** is also included for easy interaction.  

---

## 🚀 Features
- **Excel Dataset** (`household_expenses.xlsx`)  
  - 65 rows, 6 columns (Date, Category, Item, Payment Mode, Amount, Notes).  
- **Data Cleaning**  
  - Handle missing values, remove duplicates, convert types.  
  - Add validation for `amount` range and `payment_mode`.  
- **Data Analysis**  
  - Average, max, min expenses.  
  - Group totals & averages by category.  
  - Count by payment mode.  
  - Monthly totals.  
  - Top 5 expensive items.  
- **Visualization (Matplotlib/Seaborn)**  
  - Bar chart: Expenses by Category.  
  - Pie chart: Payment Mode Distribution.  
  - Line chart: Monthly Expenses.  
  - Histogram: Expense Distribution.  
- **Excel Report Export** (`cleaned_expenses.xlsx`)  
  - Multiple sheets (Summary, Category Totals, Category Averages, Payment Modes, Monthly Totals, Top 5 Items, Above 5000, Sorted Expenses, Charts).  
  - Charts embedded directly into Excel.  
- **Streamlit GUI (`gui_app.py`)**  
  - View dataset.  
  - Generate analysis report & download Excel.  
  - View charts interactively.  

---

## 📂 Project Structure
```
├── household_expenses.xlsx      # Raw dataset (input)
├── main.py                      # Main analysis script
├── gui_app.py                   # Streamlit GUI
├── cleaned_expenses.xlsx        # Output Excel report (auto-generated once run)
├── charts/                      # Folder containing generated chart images
├── README.md                    # Project documentation
└── requirements.txt             # Python dependencies
```

---

## ⚙️ Installation
1. Clone or download this repository.  
2. (Optional) Create a virtual environment:
   ```
   python -m venv .venv
   .venv\Scripts\activate   # Windows
   source .venv/bin/activate  # Linux/Mac
   ```
3. Install dependencies:
   ```
   pip install -r requirements.txt
   ```

---

## ▶️ Usage

### 1. Run the Analysis Script
```
python main.py
```
This will:
- Clean the dataset.  
- Perform analysis.  
- Export `cleaned_expenses.xlsx`.  
- Generate charts inside the `charts/` folder.  

### 2. Run the Streamlit GUI
```
streamlit run gui_app.py
```
Features:
- **View Dataset**: Preview expense records.  
- **Generate Report**: Create Excel report and download.  
- **View Charts**: See bar, pie, line, and histogram charts.  

---

## 📦 Requirements
- Python 3.8+  
- pandas  
- seaborn  
- matplotlib  
- openpyxl  
- streamlit  

Install via:
```
pip install -r requirements.txt
```

---

## 📸 Sample Outputs
- **Excel Report** → `cleaned_expenses.xlsx` with charts embedded.  
- **Charts** → Saved in `charts/` folder and displayed in Streamlit GUI.  

---

## 🧑‍💻 Author
This project was created as a **practice project for Python + Excel Data Analysis**.  
It is not a final-year submission, but a learning side project.  
