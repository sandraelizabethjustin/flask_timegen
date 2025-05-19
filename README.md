# 🗓️ TIMEGEN – Automated Timetable Generator

A smart web-based timetable generator designed to automate class scheduling for academic institutions. Built using Python, Flask, and Pandas, TIMEGEN processes Excel-based inputs to generate optimized, conflict-free class timetables.

## 🚀 Project Overview

Creating academic timetables manually is time-consuming and error-prone. TIMEGEN simplifies this by:
- Automatically allocating subjects to time slots
- Considering faculty availability and course-hour mapping
- Preventing back-to-back scheduling of the same subject
- Generating and exporting timetables in Excel format

This project was developed as part of our **B.Tech Mini Project** at **Rajiv Gandhi Institute of Technology, Kottayam**.

## ✨ Features

- 📥 Upload Excel sheets for:
  - Course-Faculty Mapping
  - Course-Hour Requirements
  - Partially Filled Timetable        <!-- ⚙️ Intelligent timetable optimization algorithm-->
- 🔄 Avoids same-subject consecutive sessions
- 🧑‍🏫 Generates individual teacher-wise timetables
- 📤 Export the final timetable to Excel
- 🌐 User-friendly web interface using HTML/CSS

## 🛠️ Tech Stack

| Tech | Purpose |
|------|---------|
| **Python** | Core programming logic |
| **Flask** | Backend framework |
| **Pandas** | Excel data manipulation |
| **HTML/CSS** | Frontend design |
| **MS Excel** | Input/Output format |

<!--## 📂 Folder Structure

flask_timegen/
├── app.py # Main Flask application
├── templates/ # HTML files
├── static/ # CSS files
├── uploads/ # Uploaded Excel sheets
├── outputs/ # Generated timetable Excel files
└── README.md # Project documentation -->
## 🧠 How It Works

1. Upload three Excel files:
   - `faculty_course.xlsx`
   - `course_hours.xlsx`
   - `partially_filled_timetable.xlsx` 

2. TIMEGEN reads and parses the data using **Pandas**.

3. The **optimization algorithm** assigns courses to time slots based on:
   - Teacher availability
   - Required teaching hours
   - Slot distribution (avoiding subject repetition)

4. The final timetable is generated and can be downloaded as an Excel file.

## ✅ Sample Inputs

You’ll need to provide the following Excel sheets:

- **Course-Faculty Mapping:** which faculty handles which course.
- **Course-Hour Mapping:** how many hours per week each course needs.
- **Partially Filled Timetable:** pre-assigned slots (labs, external sessions, etc.)

> 📌 *Ensure all Excel sheets follow the correct format as outlined in the Instructions page.*

## 🖼️ Screenshots

### 🔹 Home Page
![Home Page](Downloads/home.jpeg)

### 🔹 Upload Excel Files
![Upload](Downloads/how to upload.jpeg)

### 🔹 Generated Timetable
![Timetable](C:\Users\HP\Pictures\Screenshots\generated_timetable.png)

<!--## 🧪 Testing

- **Unit Testing:** Python functions tested for slot allocation logic
- **Black Box Testing:** Verified inputs and generated output correctness
- **Integration Testing:** End-to-end flow from upload → generation → download-->

## 🧑‍💻 Authors

Developed by a team of four CSE students at RIT Kottayam:

- [Malavika S](https://github.com/MalavikaS2002)
- [Minnu Shaji](https://github.com/Minnu-shaji)
- [P Lakshmi Priya](https://github.com/Lakshmi086)
- [Sandra Elizabeth Justin](https://github.com/sandraelizabethjustin)

## 📄 License

This project is for academic and learning purposes. Free to use under the MIT License.


<!--## 💡 Future Improvements

- Faculty-specific constraints (e.g. no classes on Friday afternoons)
- Tutorial and elective handling
- Admin login and database integration
- Mobile-responsive interface-->

## 🌟 Support

If you found this project helpful, consider giving it a ⭐️ on GitHub!
