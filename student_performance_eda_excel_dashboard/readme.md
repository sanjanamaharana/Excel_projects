# 📊 Student Performance Exploratory Data Analysis Dashboard

![Excel](https://img.shields.io/badge/Microsoft%20Excel-Analysis%20Dashboard-green)
![EDA](https://img.shields.io/badge/Project-EDA-blue)
![Status](https://img.shields.io/badge/Status-Completed-success)

## Overview

The **Student Performance Exploratory Data Analysis (EDA) Dashboard** is an interactive Excel-based analytics solution designed to evaluate student academic performance, attendance patterns, study groups, and extracurricular participation.

The workbook transforms raw student data into actionable insights through PivotTables, PivotCharts, KPI cards, and interactive slicers, enabling users to explore trends and make data-driven decisions.

### Intended Users

- Students
- Educators
- Academic Coordinators
- Data Analysts
- Excel Learners
- Educational Institutions

---

# 🎯 Project Objectives

This project aims to:

- Analyze student academic performance
- Track attendance trends
- Compare performance across genders
- Evaluate grade-level performance
- Analyze extracurricular activity participation
- Provide an interactive dashboard for quick decision-making

---

# ✨ Features & Functionality

## Data Analysis Features

- Student performance tracking
- Attendance analysis
- Gender-wise comparison
- Grade-level analysis
- Extracurricular activity analysis
- Study group analysis

## Dashboard Features

- KPI Cards
  - Total Student Count
  - Average Math Score
  - Average Science Score
  - Average English Score
  - Average Attendance %

- Interactive Slicers
  - Gender
  - Grade Level
  - Study Group
  - Extracurricular Activity

- Dynamic Pivot Charts
- Real-time filtering
- Visual performance comparison

## Excel Components Used

- Pivot Tables
- Pivot Charts
- Slicers
- Structured Tables
- Aggregation Functions
- Conditional Formatting
- Dashboard Design Techniques

---

# 📁 File Structure & Sheet Overview

| Sheet Name | Description |
|------------|-------------|
| student_performance_datasheet | Raw student dataset containing scores, attendance, demographics, and extracurricular activities |
| pivot_tables | Intermediate PivotTables used for dashboard calculations and chart generation |
| EDA Dashboard | Interactive dashboard with KPIs, charts, and slicers |

## Workflow

```text
Raw Dataset
      ↓
 Pivot Tables
      ↓
Pivot Charts
      ↓
 Dashboard
```

---

# 🗂 Dataset Structure

| Column Name | Description |
|-------------|-------------|
| student_id | Unique student identifier |
| math_score | Mathematics score |
| science_score | Science score |
| english_score | English score |
| attendance_percent | Attendance percentage |
| gender | Male/Female |
| grade_level | Student grade |
| study_group | Assigned study group |
| extra_curricular | Extracurricular activity |

---

# 📊 Dashboard Metrics

The dashboard tracks the following KPIs:

| KPI | Description |
|------|------------|
| Student Count | Total number of students |
| Avg. Math Score | Average mathematics score |
| Avg. Science Score | Average science score |
| Avg. English Score | Average English score |
| Avg. Attendance % | Average attendance percentage |

---

# 📈 Visualizations Included

### Attendance Analysis

- Attendance by Gender
- Attendance by Grade Level
- Attendance by Extracurricular Activity

### Performance Analysis

- Subject Scores by Gender
- Subject Scores by Grade Level
- Comparative Subject Performance

### Interactive Filters

- Gender
- Grade Level
- Study Group
- Extracurricular Activities

---

# 🔄 Data Sources

## Source Type

Sample educational dataset created for exploratory data analysis and dashboard development.

## Data Frequency

| Data Type | Update Frequency |
|------------|----------------|
| Student Records | Manual Update |
| Attendance Data | As Required |
| Academic Scores | As Required |

## External Connections

Currently:

```text
No external database connections
No Power Query connections
No API integrations
```

---

# ⚙️ Setup & Usage Instructions

## Requirements

- Microsoft Excel 2016 or later
- Microsoft Excel 365 (Recommended)

## Installation Steps

### 1. Download Workbook

Download:

```text
student_performance_dashboard.xlsx
```

### 2. Open in Excel

Open using:

```text
Microsoft Excel 2016+
```

### 3. Enable Editing

If prompted:

```text
Enable Editing
```

### 4. Refresh Data

Navigate:

```text
Data → Refresh All
```

### 5. Use Dashboard Filters

Interact with:

- Gender Slicer
- Grade Level Slicer
- Study Group Slicer
- Extracurricular Activity Slicer

Dashboard visuals will update automatically.

---

# 📝 Input Requirements

## Required Fields

| Field | Required |
|---------|---------|
| student_id | Yes |
| math_score | Yes |
| science_score | Yes |
| english_score | Yes |
| attendance_percent | Yes |
| gender | Yes |
| grade_level | Yes |
| study_group | Yes |
| extra_curricular | Yes |

---

## Data Validation Rules

### Score Columns

```text
Minimum: 0
Maximum: 100
```

### Attendance

```text
Minimum: 0%
Maximum: 100%
```

### Gender

Accepted Values:

```text
Male
Female
```

### Grade Level

Example Values:

```text
Grade 9
Grade 10
Grade 11
Grade 12
```

### Study Groups

```text
Group A
Group B
Group C
```

---

# 🧮 Key Formulas & Logic

## Average Score Calculation

```excel
=AVERAGE(B2:B51)
```

---

## Average Attendance

```excel
=AVERAGE(E2:E51)
```

---

## Student Count

```excel
=COUNTA(A2:A51)
```

---

## Pivot Table Aggregations

Examples:

```text
Average Math Score by Gender
Average Attendance by Grade Level
Average Science Score by Study Group
```

---

# 🔒 Assumptions

- Scores are measured out of 100.
- Attendance values are percentages.
- Each student has a unique Student ID.
- Dataset contains no duplicate records.

---

# 📋 Version History

| Version | Date | Author | Changes |
|----------|------------|---------|----------|
| 1.0 | Aug 2026 | Sanjana Maharana | Initial dataset creation |
| 1.1 | Aug 2026 | Sanjana Maharana | Added Pivot Tables |
| 1.2 | Aug 2026 | Sanjana Maharana | Added Interactive Dashboard |
| 1.3 | Aug 2026 | Sanjana Maharana | Added Slicers and KPI Cards |

---

# ⚠️ Known Limitations

| Limitation | Impact |
|------------|---------|
| Small sample dataset | Results may not represent large populations |
| Manual data updates | Requires manual refresh |
| No database connection | Cannot update automatically |
| Excel dependency | Requires Microsoft Excel |

---

# 🛠 Troubleshooting

## Dashboard Not Updating

Solution:

```text
Data → Refresh All
```

---

## Slicers Not Filtering

Solution:

```text
Right-click Slicer
→ Report Connections
→ Verify Pivot Table Connections
```

---

## Charts Display Blank Values

Possible Causes:

- Empty records
- Invalid data types
- Incorrect Pivot Table references

Solution:

```text
Refresh Pivot Tables
Check Source Data
```

---

# 🚀 Future Enhancements

Planned improvements:

- Power Query Integration
- Power BI Dashboard Version
- Automated Data Refresh
- Predictive Analytics
- Student Performance Forecasting
- Advanced KPI Monitoring

---

# 👩‍💻 Author

**Sanjana Maharana**

MCA Student | Data Analytics Enthusiast | Excel Dashboard Developer

### Skills

- Microsoft Excel
- Data Analysis
- Data Visualization
- Python
- Pandas
- SQL
- Power BI

---

# 📬 Contact & Support

For questions, suggestions, or collaboration:

- GitHub: https://github.com/sanjanamaharana
- LinkedIn: *(Add your LinkedIn Profile URL)*
- Email: *(Add your Email Address)*

---

# 📄 License

This project is intended for educational and portfolio purposes.

Feel free to fork, learn from, and improve the project with proper attribution.

---

⭐ If you found this project useful, consider giving the repository a star.
