# 📊 Project Analytics Dashboard (Streamlit)

An interactive **Project Monitoring & Analytics Dashboard** built using **Streamlit**, designed to track project progress, visualize timelines (Gantt charts), and identify delays in real-time.

---

## 🚀 Features

- 📊 **Interactive Gantt Chart**
  - Visualize project timelines by *project, vendor, or status*
  
- 📈 **Progress Tracking**
  - Compare **planned vs actual progress**
  
- ⚠️ **Delay & Risk Analysis**
  - Identify:
    - Overdue activities
    - Behind-schedule tasks
    - Lag in percentage points

- 📌 **Project Summary**
  - Weighted metrics based on activity duration
  - Export summary as CSV

- 🔍 **Advanced Filters**
  - Filter by:
    - Vendor
    - Project

- ⏱️ **Auto Refresh**
  - Real-time updates with configurable refresh interval

---

## 🛠️ Tech Stack

- **Frontend & App Framework**: Streamlit  
- **Data Processing**: Pandas, NumPy  
- **Visualization**: Plotly  
- **Excel Handling**: OpenPyXL  

---

## 📂 Project Structure
├── app.py # Main dashboard (enhanced UI + features)
├── gantt.py # Simpler version of dashboard
├── sample.xlsx # Sample input data
├── requirements.txt # Dependencies
├── README.md


---

## 📊 Input Data Format

Your Excel file must contain the following columns:

| Column Name   | Description |
|--------------|------------|
| vendor       | Vendor name |
| project      | Project name |
| activity     | Task/activity |
| start        | Start date |
| end          | End date |
| %complete    | Completion percentage |

✅ Notes:
- Column names are **case-insensitive**
- `%complete` can be like `45%` or `45`
- Vendor/Project can be blank (auto forward-filled)

---

## ▶️ How to Run

### 1. Clone the repository
```bash
git clone https://github.com/your-username/project-analytics-dashboard.git
cd project-analytics-dashboard
```

### 2. Install dependencies
```bash
pip install -r requirements.txt
```

### 3. Run the app
```bash
streamlit run app.py
```

---

## ⚙️ Usage

1. Upload or specify your Excel file path  
2. Select:
   - Sheet name (optional)
   - Filters (vendor/project)
3. Adjust:
   - “Today” date for lag calculation
   - Visualization preferences
4. Explore:
   - Gantt Chart
   - Risk Analysis
   - Project Summary

---

## 📈 Key Metrics Explained

- **Planned Progress**  
  Based on time elapsed between start and end date  

- **Actual Progress**  
  Based on `%complete`  

- **Behind Schedule**  
  When:
  ```
  Planned Progress > Actual Progress (by > 5%)
  ```

- **Overdue**  
  When:
  ```
  %Complete < 100 AND Today > End Date
  ```

---

## 🧠 Future Improvements

- Integrate with databases (PostgreSQL / Firebase) for scalable data storage  
- Implement multi-user authentication and role-based access  
- Enable API-based data ingestion for real-time updates  
- Add automated alerts and notifications for delays  
- Deploy on cloud platforms (AWS / Streamlit Cloud) for wider accessibility   
