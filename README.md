# 🏆 HR Certificate Generation System

This project automates the generation of internship or HR certificates for students or employees using data from an Excel sheet and a Word template. It generates personalized `.docx` certificates in bulk using Python, making HR documentation efficient and error-free.

---

## 📌 Key Features

- ⚡ Bulk generation of internship or work certificates
- 📄 Auto-fill data like name, domain, dates, and duration
- 📅 Date formatting and month calculation
- 🧠 Intelligent Word template replacement (preserving formatting)
- 💾 Outputs clean `.docx` certificates named after each student

---

## 🧠 Technologies Used

| Tool/Library       | Purpose                                         |
|--------------------|-------------------------------------------------|
| **Python 3**       | Core programming language                       |
| **pandas**         | Data manipulation and Excel reading             |
| **python-docx**    | Editing `.docx` templates with dynamic content  |
| **os**             | File handling and directory management          |
| **datetime**       | Date formatting and duration calculation        |

---

## 🗂️ Project Structure

```
HR_Certificate_Generation/
│
├── .dist/                     # Auto-generated build/cache folder (optional)
├── data/
│   └── student_data.xlsx      # Excel file with student info (Name, Domain, Dates)
├── output/
│   └── *.docx                 # Output certificates saved here
├── templates/
│   └── internship_template.docx  # Certificate Word template with placeholders
├── generate_certificates.py   # Legacy/basic generator script (optional)
├── generate_certificate_final.py # ✅ Final working script for generating certificates
├── requirements.txt           # All required Python dependencies
└── README.md                  # This file
```

---

## 📥 Excel Input Format (`data/student_data.xlsx`)

Make sure the Excel file has the following column headers (case-insensitive):

| NAME | DOMAIN | START DATE | END DATE | ISSUE DATE |
|------|--------|-------------|----------|-------------|

Example:

| NAME          | DOMAIN         | START DATE | END DATE   | ISSUE DATE |
|---------------|----------------|------------|------------|-------------|
| John Doe      | Web Development| 2024-01-10 | 2024-03-10 | 2024-03-11  |

---

## 📄 Word Template Format (`templates/internship_template.docx`)

Use placeholders in your `.docx` template:

- `{Student Name}`
- `{Domain Name}`
- `{Start date}`
- `{End date}`
- `{Issue date}`
- `{No of months}`

> These will be replaced dynamically for each student.

---

## ⚙️ How to Run

### ✅ 1. Install dependencies

```bash
pip install -r requirements.txt
```

Or manually:

```bash
pip install pandas python-docx openpyxl
```

### ✅ 2. Place your Excel file and `.docx` template in correct folders

- Excel → `data/student_data.xlsx`
- Template → `templates/internship_template.docx`

### ✅ 3. Run the certificate generator

```bash
python generate_certificate_final.py
```

✅ All certificates will be saved in the `output/` folder.

---

## 🖼️ Sample Certificate Output

The generated certificates will be saved like:

```
output/
├── John_Doe_Internship_Certificate.docx
├── Alice_Smith_Internship_Certificate.docx
```

Each `.docx` contains fully customized data pulled from the Excel file.

---

## 🧑‍💻 Author

**Solomon Goodwin Alexander**  
📍 Nagpur, India  
🔗 [GitHub Profile](https://github.com/Solomon-Alexander1)  
🔗 [LinkedIn Profile](https://www.linkedin.com/in/solomon-alexander-184733170/)

---

## 🛡️ License

This project is open-source and available under the **MIT License**.  
Feel free to use, modify, or contribute for educational or professional use.

---

## 🚀 Future Enhancements (Optional)

- Export directly as `.pdf` format  
- Add logo/signature support  
- Web interface for HR/Training managers  
- Emailing certificates to recipients  

---

📢 **If you find this project useful, don't forget to give it a ⭐ on GitHub!**
