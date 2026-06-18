# STUDENT-RESULT-ANALYSIS-AND-REPORT-CARD-GENERATOR
Project useing advance Excel to generate report card and analysis performance.

# 📊 ABC School — Student Result Analysis Dashboard

An interactive Excel-based student academic performance analysis system covering 503 students across Sections A, B, and C, with automated report card generation, pivot-driven dashboards, and subject-wise insights.

---

## 📁 Workbook Structure
The workbook contains the following sheets:

| Sheet | Description |
| :--- | :--- |
| **raw_data** | Original unprocessed student records (Roll No, Name, Section, Gender, Subject Marks, Attendance) |
| **clean data** | Cleaned and processed student data with computed Total, Percentage, Grade, Rank, and Result |
| **dashbord** | Interactive dashboard with pivot charts and navigation links |
| **ATTENDANCE** | Section-wise average attendance pivot summary |
| **sec wise subj** | Section-wise average subject marks pivot summary |
| **section_wise_grade_card** | Section-wise grade count pivot table |
| **avg of gen_wise marks** | Gender-wise average marks per subject |
| **vlookup_page** | Individual student report card generator (enter Roll No to auto-generate) |
| **SUMMARY** | Class-wise performance summary with topper, pass/fail counts, grade distribution |
| **Sheet_review** | Gender-based student review pivot |
| **INSTUCTION** | Step-by-step task instructions for data cleaning and dashboard building |

---

## 🗂️ Data Overview

### Raw Data Columns
| Column | Description |
| :--- | :--- |
| **Roll No** | Unique student identifier (e.g. S001) |
| **Student Name** | Full name of the student |
| **Section** | Class section — A, B, or C |
| **Gender** | Male / Female / No gender |
| **Mathematics** | Marks out of 100 |
| **Science** | Marks out of 100 |
| **English** | Marks out of 100 |
| **Hindi** | Marks out of 100 |
| **Social Science** | Marks out of 100 |
| **Attendance (%)** | Attendance percentage |

### Computed Columns (Clean Data)
| Column | Description |
| :--- | :--- |
| **Total** | Sum of all 5 subject marks (out of 500) |
| **Percentage** | Overall percentage score |
| **Grade** | A++, A, B, C, or D |
| **Rank** | Class rank based on percentage |
| **Result** | Pass / Fail |

---

## 📈 Key Statistics

### Overall Summary
| Metric | Value |
| :--- | :--- |
| **Total Students** | 503 |
| **Pass Count** | 358 |
| **Fail Count** | 115 |
| **Overall Pass Rate** | 75.7% |
| **Overall Average %** | 69.78% |
| **Highest Score** | 98.6% (Rishi Sahoo) |

### Section-wise Performance
| Section | Students | Pass | Fail | Pass Rate | Avg % | Topper |
| :--- | :--- | :--- | :--- | :--- | :--- | :--- |
| **10-A** | 153 | 106 | 47 | 69.3% | 67.25% | Anjali Naidu |
| **10-B** | 169 | 134 | 35 | 79.3% | 69.8% | Pari Sahoo |
| **10-C** | 151 | 118 | 33 | 78.1% | 79.05% | Dev Jha |

### Section-wise Attendance
| Section | Average Attendance |
| :--- | :--- |
| **A** | 81.34% |
| **B** | 82.94% |
| **C** | 82.85% |
| **Overall** | 82.37% |

### Grade Distribution (All Students)
| Grade | Count | % of Students |
| :--- | :--- | :--- |
| **A++** | 40 | 7.95% |
| **A** | 122 | 24.25% |
| **B** | 106 | 21.07% |
| **C** | 113 | 22.47% |
| **D** | 119 | 23.66% |

### Subject-wise Average Marks
| Subject | Section A | Section B | Section C | Overall |
| :--- | :--- | :--- | :--- | :--- |
| **Mathematics** | 67.0 | 70.0 | 74.5 | 69.7 |
| **Science** | 68.5 | 70.7 | 76.8 | 69.6 |
| **English** | 68.3 | 68.0 | 73.8 | 70.4 |
| **Hindi** | 65.0 | 70.3 | 76.5 | 68.9 |
| **Social Science** | 67.5 | 70.0 | 76.5 | 69.9 |

### Gender-wise Performance
| Gender | Social Science | Science | Mathematics | English | Hindi |
| :--- | :--- | :--- | :--- | :--- | :--- |
| **Female** | 69.52 | 69.91 | 69.35 | 67.79 | 68.35 |
| **Male** | 72.63 | 72.38 | 71.12 | 70.19 | 70.85 |

---

## 🔧 How to Use

### 1. Generate a Student Report Card
* Go to the `vlookup_page` sheet.
* Enter a Roll Number in the designated cell (e.g., `S001`).
* The report card will auto-populate with the student's name, section, rank, percentage, subject-wise marks, grade, and teacher's remarks.
<img width="1039" height="699" alt="generate" src="https://github.com/user-attachments/assets/61df8784-9953-41f5-b95c-5aede37046cf" />


### 2. Explore the Dashboard
* Open the `dashbord` sheet.
* Use the navigation buttons to jump to:
  * Section-wise Grade Count
  * Section-wise Subject Averages
  * Class Summary
  * Result Card Generator
 
<img width="1329" height="742" alt="report_card" src="https://github.com/user-attachments/assets/1bb2a8a9-1176-48f5-800f-3f62fed23148" />



### 3. Explore Pivot Summaries
Navigate directly to individual summary sheets for deeper insights:
* `ATTENDANCE` — Section-wise attendance averages
* `sec wise subj` — Section-wise subject performance
* `section_wise_grade_card` — Grade distribution by section
* `avg of gen_wise marks` — Gender-based subject averages
* `SUMMARY` — Full class-level performance overview

---

## 🧹 Data Cleaning Tasks
*(As per the `INSTUCTION` sheet)*

| Task | Action |
| :--- | :--- |
| **Task 1** | Download raw data |
| **Task 2** | Remove duplicates (Data tab → Remove Duplicates); fill null/blank fields using `IF` and `COUNTBLANK` functions |
| **Task 3 & 4** | Apply cleaning procedures thoughtfully |
| **Task 5** | Build interactive dashboard using Pivot Tables and Pivot Charts |
| **Task 6** | Final workbook must contain: Raw Data, Cleaning Task, Cleaned Data, and Dashboard sheet |

---

## 📋 Grading Scale

| Grade | Percentage Range |
| :--- | :--- |
| **A++** | 90% and above |
| **A** | 80% – 89% |
| **B** | 70% – 79% |
| **C** | 60% – 69% |
| **D** | Below 60% |

---


