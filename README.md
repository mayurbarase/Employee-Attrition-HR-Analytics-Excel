# 👥 Employee Attrition & HR Analytics Dashboard (Excel)

An interactive HR analytics dashboard built in Microsoft Excel that analyzes employee attrition across department, job role, age, salary, tenure, gender, and job satisfaction.

Filter by **Department, Job Role, Gender, Age Group, or Salary Range**, and every chart and KPI updates instantly.

![Employee Attrition Dashboard](images/dashboard-full.png)

---

## 📌 Why I Built This

Attrition is one of the most important things an HR team tracks, but raw employee records don't make it obvious *where* people are leaving from or *why*. This is my second Data Analytics project, and I wanted to go beyond static numbers and build something that lets you actually explore the patterns:

- What's the overall attrition rate?
- Which departments and job roles lose the most people?
- Does attrition change with age, tenure, or salary?
- Is there a link between job satisfaction and attrition?
- How does attrition differ by gender?

The dashboard turns the raw HR dataset into a single, filterable view built entirely with **Excel Tables, formulas, PivotTables, PivotCharts, and Slicers** — no Power Query, no VBA.

---

## ✨ Key Features

- **5 KPI cards** for Total Employees, Employees Left, Attrition Rate, Average Salary, and Average Tenure
- **7 analytical views** covering department, job role, age, salary, tenure, gender, and job satisfaction
- **5 interactive slicers** (Department, Job Role, Gender, Age Group, Salary Range) linked to every PivotTable
- **Calculated categories** (Age Group, Salary Range) that make the raw data easier to read
- **Clean HR-analytics visual theme** with a dark navy sidebar and consistent card styling

---

## 📈 Headline Numbers

| KPI | Value |
|---|---|
| Total Employees | 1,470 |
| Employees Left | 237 |
| Attrition Rate | 16.12% |
| Average Salary | ≈ ₹6,503 |
| Average Years at Company | ≈ 7.0 years |

---

## 🗂️ Dataset

Built on the **IBM HR Analytics Employee Attrition & Performance** dataset — a fictional HR dataset created by IBM data scientists, commonly used for practicing HR and people-analytics workflows.

- **Records:** 1,470 employees
- **Features:** 35 columns
- **Target variable:** Attrition (Yes / No)

The dataset covers employee demographics, department, job role, monthly income, job satisfaction, years at company, business travel, education, gender, overtime, performance, work-life balance, and other HR attributes.

---

## 🧮 KPI Calculations

```excel
Total Employees             =COUNTA(EmployeeData[EmployeeNumber])
Employees Left               =COUNTIF(EmployeeData[Attrition],"Yes")
Attrition Rate                =B4/B3        // Employees Left / Total Employees
Average Salary                =AVERAGE(EmployeeData[MonthlyIncome])
Average Years at Company      =AVERAGE(EmployeeData[YearsAtCompany])
```

### Calculated Columns

To make age and income easier to analyze, two custom categories were added directly to the employee table:

```excel
Age Group
=IF([@Age]<25,"Under 25",IF([@Age]<35,"25-34",IF([@Age]<45,"35-44",IF([@Age]<55,"45-54","55+"))))

Salary Range
=IF([@MonthlyIncome]<3000,"Below 3K",IF([@MonthlyIncome]<6000,"3K-6K",IF([@MonthlyIncome]<9000,"6K-9K",IF([@MonthlyIncome]<12000,"9K-12K","12K+"))))
```

Grouping ages and salaries into bands makes the PivotTables and charts far easier to read than plotting every individual value.

---

## 📊 Dashboard Analysis

| View | PivotTable Setup | Chart Type | What It Shows |
|---|---|---|---|
| Attrition by Department | Rows: Department · Columns: Attrition · Values: Count of EmployeeNumber | Clustered Column | Stayed vs. left, compared across departments |
| Attrition by Job Role | Rows: JobRole · Columns: Attrition · Values: Count of EmployeeNumber | Stacked Column | Retention and attrition by role |
| Attrition by Age Group | Rows: Age Group · Columns: Attrition · Values: Count of EmployeeNumber | Line Chart with Markers | How attrition shifts across age brackets |
| Attrition by Salary Range | Rows: Salary Range · Columns: Attrition · Values: Count of EmployeeNumber | Doughnut Chart | Attrition split across income levels |
| Attrition by Years at Company | Rows: YearsAtCompany · Columns: Attrition · Values: Count of EmployeeNumber | Horizontal Bar | Attrition by tenure |
| Attrition by Gender | Rows: Gender · Columns: Attrition · Values: Count of EmployeeNumber | Pie Chart | Gender split of attrition |
| Job Satisfaction vs. Attrition | Rows: JobSatisfaction (1–4) · Columns: Attrition · Values: Count of EmployeeNumber | 100% Stacked Column | Whether lower satisfaction correlates with leaving |

### 🎛️ Slicers

Five slicers — **Department, Job Role, Gender, Age Group, and Salary Range** — are connected to the dashboard's PivotTables. Selecting a department, for instance, instantly filters every chart down to that department's employee population.

![Department Filter](images/dashboard-department-filter.png)
![Gender Filter](images/dashboard-gender-filter.png)

---

## 🛠️ Tools & Skills

**Microsoft Excel** · Excel Tables · Structured References · PivotTables · PivotCharts · Slicers · Conditional Category Building · KPI Development · Dashboard Design · HR Analytics

Functions used: `COUNTA`, `COUNTIF`, `AVERAGE`, `IF`, `TEXT`

> No VBA/macros or Power Query were used. Everything is built with native Excel features.

---

## 🎨 Design Approach

The dashboard uses a clean HR/business analytics theme:

- Dark navy sidebar with a light blue dashboard background
- White KPI and chart cards for contrast
- A consistent blue and orange color scheme across visuals
- Custom-styled slicers and KPI icons
- A dedicated Key Insights panel
- Consistent typography and spacing throughout

The aim was a dashboard that reads clearly even to someone who's never seen the underlying data.

---

## 📁 Repository Structure

```text
Employee-Attrition-HR-Analytics-Excel/
│
├── Employee_Attrition_HR_Analytics.xlsx
├── README.md
└── images/
    ├── dashboard-full.png
    ├── dashboard-department-filter.png
    └── dashboard-gender-filter.png
```

---

## ▶️ How to Use

1. Download `Employee_Attrition_HR_Analytics.xlsx`.
2. Open it in Microsoft Excel (desktop is recommended for full slicer support).
3. Go to the **Dashboard** sheet and use the slicers to explore attrition by department, role, age, salary, or gender.

---

## 💡 What I Learned

This project pushed me past just calculating numbers and into building something people can actually explore. Working with a real-style HR dataset meant thinking carefully about how to group continuous data (age, salary) into meaningful categories, how to choose a chart that fits the question being asked, and how to keep five slicers all talking to the right PivotTables without breaking the layout.

The biggest takeaway: a dashboard is only useful if someone unfamiliar with the data can look at it and immediately understand what's being measured and why it matters.

---

## 🚀 Project Outcome

The dashboard lets you explore attrition across:

**Department → Job Role → Age → Salary → Tenure → Gender → Job Satisfaction**

It's part of my ongoing journey of learning data analytics through hands-on projects, building practical skills across **Excel, SQL, Python, and Power BI**.

---

## 👤 Author

**Mayur**
Aspiring Data Analyst | Excel · SQL · Python · Power BI

> Learning by building, analyzing, and continuously improving.

- **GitHub:** [your-username](https://github.com/your-username)
- **LinkedIn:** [your-name](https://linkedin.com/in/your-profile)

---

**Status:** Completed ✅

⭐ If you found this project useful, consider giving the repo a star!
