# 📊 محول قوائم الامتحانات | Algerian Exam List Converter

أداة بايثون لتحويل قوائم امتحانات المستخرجة من الديوان من صيغة Markdown إلى ملفات Excel منسقة بشكل احترافي مع دعم كامل للغة العربية واتجاه الكتابة من اليمين إلى اليسار (RTL).

A Python tool to convert Algerian Education Ministry exam lists from Markdown format to professionally formatted Excel files with full Arabic RTL support.

![Python](https://img.shields.io/badge/Python-3.7+-blue.svg)
![License](https://img.shields.io/badge/License-MIT-green.svg)

---

## 📋 المميزات | Features

- ✅ **تحويل Markdown إلى Excel**
  تحويل قوائم الامتحانات من Markdown إلى ملفات Excel بسهولة
  **Markdown to Excel Conversion** – Transform exam lists into Excel spreadsheets

- ✅ **دعم اللغة العربية (RTL)**
  تنسيق كامل للنصوص العربية من اليمين إلى اليسار
  **Arabic RTL Support** – Full right-to-left layout for Arabic text

- ✅ **تنظيم حسب اللجان**
  تقسيم التلاميذ تلقائيًا حسب لجان الامتحان
  **Committee-Based Organization** – Automatically separate students by committees

- ✅ **ملف واحد أو عدة أوراق**
  إمكانية إنشاء ملف واحد أو عدة أوراق حسب اللجان
  **Single & Multi-Sheet Output** – Generate one file or multiple sheets

- ✅ **تنسيق احترافي**
  عناوين منسقة، حدود، ألوان، وأعمدة مضبوطة
  **Professional Formatting** – Styled headers, borders, colors, and widths

- ✅ **واجهتان (CLI + Web)**
  استعمال عبر سطر الأوامر أو واجهة ويب
  **Dual Interface** – CLI and web-based interfaces available

- ✅ **رؤوس ديناميكية**
  استخراج تلقائي لمعلومات المركز، الشعبة، والولاية
  **Dynamic Headers** – Auto-extract center, branch, and wilaya info

---

## 📦 التثبيت | Installation

### المتطلبات | Prerequisites

- Python 3.7 أو أحدث
- pip (مدير الحزم)

### الخطوة 1: تحميل المشروع | Clone or Download

```bash
git clone <repository-url>
cd ExamList_Converter
```

---

## 🚀 الاستخدام | Usage

### عبر سطر الأوامر | CLI

```bash
python md_to_excel.py input.md output.xlsx
```

### عبر الويب | Web Interface

```bash
python md_to_excel_web.py
```

ثم افتح المتصفح:
Then open in your browser:
👉 http://localhost:5000

---

## 📁 هيكلة المشروع | Project Structure

```
ExamList_Converter/
│── md_to_excel.py
│── md_to_excel_web.py
│── templates/
│── static/
│── requirements.txt
│── README.md
```

---

## 🧑‍💻 المطور | Developer

تم تطوير هذا المشروع لتسهيل العمل الإداري داخل مديريات التربية ومراكز الامتحانات في الجزائر.

This project was developed to simplify administrative workflows in Algerian education directorates and exam centers.

---

## 📜 الرخصة | License

هذا المشروع مفتوح المصدر تحت رخصة MIT.
This project is open-source under the MIT License.

---

## ⭐ ملاحظة | Note

إذا أعجبك المشروع، لا تنسَ وضع ⭐ على المستودع!
If you like this project, don’t forget to ⭐ the repository!
