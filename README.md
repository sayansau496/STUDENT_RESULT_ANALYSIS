# 🎓 Student Result Analysis and Report Card Generator

<img width="493" height="356" alt="image" src="https://github.com/user-attachments/assets/909968d7-da3e-4319-b288-5dea501c3267" />


An interactive **Advanced Excel-based Student Result Analysis and Report Card Generator** designed to analyze academic performance, attendance, subject-wise marks, grades, ranks, and section-wise results.

The project covers **503 students** across Sections **A, B, and C**. It includes automated report-card generation using VLOOKUP, Pivot Tables, Pivot Charts, dashboards, and performance summaries.

---

## 📌 Project Objectives

* Analyze student academic performance across multiple subjects.
* Calculate total marks, percentage, grade, rank, and pass/fail status.
* Generate individual student report cards automatically using Roll Number.
* Compare section-wise performance and attendance.
* Analyze subject-wise and gender-wise performance.
* Build an interactive dashboard using Pivot Tables and Pivot Charts.
* Identify toppers, weak-performing sections, and grade distribution.

---

## 📊 Dashboard Preview

<img width="493" height="356" alt="image" src="https://github.com/user-attachments/assets/ff255c39-e333-4dfa-b084-80b0e93fbe3a" />


---

## 📁 Workbook Structure

| Sheet Name                  | Description                                                                                  |
| :-------------------------- | :------------------------------------------------------------------------------------------- |
| **raw_data**                | Original student records including Roll Number, Name, Section, Gender, Marks, and Attendance |
| **clean data**              | Processed student data with Total, Percentage, Grade, Rank, and Result                       |
| **dashbord**                | Interactive dashboard with Pivot Charts and navigation links                                 |
| **ATTENDANCE**              | Section-wise average attendance summary                                                      |
| **sec wise subj**           | Section-wise average marks for each subject                                                  |
| **section_wise_grade_card** | Grade distribution by section                                                                |
| **avg of gen_wise marks**   | Gender-wise average marks for all subjects                                                   |
| **vlookup_page**            | Automated individual student report card generator                                           |
| **SUMMARY**                 | Class-wise performance summary including topper, pass/fail count, and grade distribution     |
| **Sheet_review**            | Gender-based student performance review                                                      |
| **INSTUCTION**              | Instructions for data cleaning and dashboard development                                     |

---

## 🗂️ Dataset Overview

### Raw Data Columns

| Column             | Description                      |
| :----------------- | :------------------------------- |
| **Roll No**        | Unique student ID such as `S001` |
| **Student Name**   | Full name of the student         |
| **Section**        | Student section: A, B, or C      |
| **Gender**         | Male, Female, or No Gender       |
| **Mathematics**    | Marks out of 100                 |
| **Science**        | Marks out of 100                 |
| **English**        | Marks out of 100                 |
| **Hindi**          | Marks out of 100                 |
| **Social Science** | Marks out of 100                 |
| **Attendance (%)** | Student attendance percentage    |

### Calculated Columns

| Column         | Description                                  |
| :------------- | :------------------------------------------- |
| **Total**      | Sum of marks in all five subjects out of 500 |
| **Percentage** | Overall percentage obtained by the student   |
| **Grade**      | Grade based on percentage                    |
| **Rank**       | Student rank based on percentage             |
| **Result**     | Pass or Fail status                          |

---

## 📈 Key Statistics

### Overall Performance Summary

| Metric                         | Value       |
| :----------------------------- | :---------- |
| **Total Students**             | 503         |
| **Pass Count**                 | 358         |
| **Fail Count**                 | 115         |
| **Overall Pass Rate**          | 75.7%       |
| **Overall Average Percentage** | 69.78%      |
| **Highest Score**              | 98.6%       |
| **Top Scorer**                 | Rishi Sahoo |

### Section-wise Performance

| Section  | Students | Pass | Fail | Pass Rate | Average % | Topper       |
| :------- | :------- | :--- | :--- | :-------- | :-------- | :----------- |
| **10-A** | 153      | 106  | 47   | 69.3%     | 67.25%    | Anjali Naidu |
| **10-B** | 169      | 134  | 35   | 79.3%     | 69.8%     | Pari Sahoo   |
| **10-C** | 151      | 118  | 33   | 78.1%     | 79.05%    | Dev Jha      |

### Section-wise Attendance

| Section     | Average Attendance |
| :---------- | :----------------- |
| **A**       | 81.34%             |
| **B**       | 82.94%             |
| **C**       | 82.85%             |
| **Overall** | 82.37%             |

### Grade Distribution

| Grade   | Student Count | Percentage of Students |
| :------ | :------------ | :--------------------- |
| **A++** | 40            | 7.95%                  |
| **A**   | 122           | 24.25%                 |
| **B**   | 106           | 21.07%                 |
| **C**   | 113           | 22.47%                 |
| **D**   | 119           | 23.66%                 |

### Subject-wise Average Marks

| Subject            | Section A | Section B | Section C | Overall Average |
| :----------------- | :-------- | :-------- | :-------- | :-------------- |
| **Mathematics**    | 67.0      | 70.0      | 74.5      | 69.7            |
| **Science**        | 68.5      | 70.7      | 76.8      | 69.6            |
| **English**        | 68.3      | 68.0      | 73.8      | 70.4            |
| **Hindi**          | 65.0      | 70.3      | 76.5      | 68.9            |
| **Social Science** | 67.5      | 70.0      | 76.5      | 69.9            |

### Gender-wise Performance

| Gender     | Social Science | Science | Mathematics | English | Hindi |
| :--------- | :------------- | :------ | :---------- | :------ | :---- |
| **Female** | 69.52          | 69.91   | 69.35       | 67.79   | 68.35 |
| **Male**   | 72.63          | 72.38   | 71.12       | 70.19   | 70.85 |

---

## 🔧 Excel Features Used

* Advanced Excel Formulas
* VLOOKUP
* IF Function
* COUNTBLANK Function
* SUM Function
* AVERAGE Function
* RANK Function
* Conditional Formatting
* Pivot Tables
* Pivot Charts
* Dashboard Navigation Links
* Data Cleaning
* Report Card Automation

---

## 🧮 Important Excel Formulas

### Total Marks

```excel
=SUM(E2:I2)
```

### Percentage

```excel
=J2/500*100
```

### Grade Calculation

```excel
=IF(K2>=90,"A++",
IF(K2>=80,"A",
IF(K2>=70,"B",
IF(K2>=60,"C","D"))))
```

### Result Calculation

```excel
=IF(MIN(E2:I2)>=35,"Pass","Fail")
```

### Rank Calculation

```excel
=RANK(K2,$K$2:$K$504,0)
```

### Student Name Lookup for Report Card

```excel
=VLOOKUP(B3,'clean data'!$A$2:$N$504,2,FALSE)
```

> In the report card generator, enter the Roll Number in the designated input cell. The VLOOKUP formulas automatically retrieve student details, marks, percentage, grade, rank, and result.

---

## 🖥️ How to Use

### 1. Generate an Individual Report Card

1. Open the `vlookup_page` sheet.
2. Enter a Roll Number in the input cell, for example: `S001`.
3. The report card automatically displays:

   * Student Name
   * Section
   * Subject-wise Marks
   * Total Marks
   * Percentage
   * Grade
   * Rank
   * Result
   * Teacher Remarks

### 2. Explore the Dashboard

1. Open the `dashbord` sheet.
2. Review key performance indicators and charts.
3. Use navigation buttons to access:

   * Section-wise Grade Count
   * Section-wise Subject Average
   * Attendance Summary
   * Class Summary
   * Individual Report Card Generator

### 3. Explore Pivot Summaries

| Sheet                       | Analysis Available                 |
| :-------------------------- | :--------------------------------- |
| **ATTENDANCE**              | Section-wise attendance averages   |
| **sec wise subj**           | Section-wise average subject marks |
| **section_wise_grade_card** | Grade count by section             |
| **avg of gen_wise marks**   | Gender-wise average subject marks  |
| **SUMMARY**                 | Class-level performance overview   |

---

## 🧹 Data Cleaning Process

| Task       | Activity                                                             |
| :--------- | :------------------------------------------------------------------- |
| **Task 1** | Download and import raw student data                                 |
| **Task 2** | Remove duplicate records using **Data → Remove Duplicates**          |
| **Task 3** | Identify and handle blank or null values using `IF` and `COUNTBLANK` |
| **Task 4** | Standardize student details, marks, sections, and gender values      |
| **Task 5** | Create Pivot Tables and Pivot Charts for analysis                    |
| **Task 6** | Build the interactive dashboard and report card generator            |

---

## 📋 Grading Scale

| Grade   | Percentage Range |
| :------ | :--------------- |
| **A++** | 90% and above    |
| **A**   | 80% – 89%        |
| **B**   | 70% – 79%        |
| **C**   | 60% – 69%        |
| **D**   | Below 60%        |

---

## 📂 Recommended Repository Structure

```text
STUDENT-RESULT-ANALYSIS-AND-REPORT-CARD-GENERATOR/
│
├── README.md
├── Student_Result_Analysis.xlsx
│
├── images/
│   ├── dashboard-preview.png
│   └── report-card-preview.png
│
└── docs/
    └── project-description.pdf
```

---

## 💡 Key Insights

* Section **10-C** has the highest average percentage among all sections.
* Section **10-B** has the highest pass rate.
* Mathematics and Social Science show strong overall performance.
* Attendance is consistently above 80% across all sections.
* Grade **A** has the highest number of students.
* The automated report card generator reduces manual work and improves reporting accuracy.

---

## 🚀 Future Improvements

* Add student search using `XLOOKUP`.
* Add dynamic drop-down lists for Roll Number selection.
* Create printable report cards using Excel page layout settings.
* Add parent contact and teacher feedback sections.
* Add conditional formatting for weak subject identification.
* Connect the workbook to Power BI for advanced visualization.
* Automate report-card export to PDF using VBA or Office Scripts.

---

## 👤 Author

**Sayan Sau**

* GitHub: https://github.com/sayansau496
* LinkedIn: https://www.linkedin.com/in/sayan-sau-8239216b/
* Project Repository: https://github.com/sayansau496/STUDENT_RESULT_ANALYSIS

---

## ⭐ Support

If you found this project useful, please consider giving the repository a star.

Repository: https://github.com/sayansau496/STUDENT_RESULT_ANALYSIS

---

## 📄 License

This project is created for **educational and portfolio purposes**. Please ensure that any student data used in the project is anonymized and handled responsibly.
