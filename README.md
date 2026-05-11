# 📊 محول قوائم المناداة باللجان | Algerian Exam Call Lists Converter

أداة بايثون لتحويل **قوائم المناداة باللجان** المستخرجة من **الديوان الوطني للإمتحانات والمسابقات** من صيغة Markdown إلى ملفات Excel منسقة بشكل احترافي مع دعم كامل للغة العربية واتجاه الكتابة من اليمين إلى اليسار (RTL).

> ⚠️ ملاحظة: يتم استخراج ملف Markdown (MD) انطلاقًا من ملف PDF عبر تحويله باستخدام الموقع التالي:
> https://pdf2md.morethan.io/

A Python tool to convert **exam call lists by committees** from Markdown format (extracted from PDF files) into professionally formatted Excel files with full Arabic RTL support.

> ⚠️ Note: The Markdown file (MD) is generated from a PDF file using:
> https://pdf2md.morethan.io/

![Python](https://img.shields.io/badge/Python-3.7+-blue.svg)
![License](https://img.shields.io/badge/License-MIT-green.svg)

---

## 📋 المميزات | Features

* ✅ **تحويل Markdown إلى Excel**
  تحويل قوائم المناداة باللجان من Markdown إلى ملفات Excel بسهولة
  **Markdown to Excel Conversion** – Transform call lists into Excel spreadsheets

* ✅ **دعم اللغة العربية (RTL)**
  تنسيق كامل للنصوص العربية من اليمين إلى اليسار
  **Arabic RTL Support** – Full right-to-left layout for Arabic text

* ✅ **تنظيم حسب اللجان**
  تقسيم المترشحين تلقائيًا حسب لجان الامتحان
  **Committee-Based Organization** – Automatically separate candidates by committees

* ✅ **ملف واحد أو عدة أوراق**
  إمكانية إنشاء ملف واحد أو عدة أوراق حسب اللجان
  **Single & Multi-Sheet Output** – Generate one file or multiple sheets

* ✅ **تنسيق احترافي**
  عناوين منسقة، حدود، ألوان، وأعمدة مضبوطة
  **Professional Formatting** – Styled headers, borders, colors, and widths

* ✅ **واجهتان (CLI + Web)**
  استعمال عبر سطر الأوامر أو واجهة ويب
  **Dual Interface** – CLI and web-based interfaces available

* ✅ **رؤوس ديناميكية**
  استخراج تلقائي لمعلومات المركز، الشعبة، والولاية
  **Dynamic Headers** – Auto-extract center, branch, and wilaya info

---

## 📦 التثبيت | Installation

### المتطلبات | Prerequisites

* Python 3.7 أو أحدث
* pip (مدير الحزم)

### الخطوة 1: تحميل المشروع | Clone or Download

```bash
git clone https://github.com/chebichebkheireddine/ExamList_Converter.git
cd ExamList_Converter
```

---

## 🚀 الاستخدام | Usage

### 🔹 الخطوة 1: تحويل PDF إلى Markdown

قم أولاً بتحويل ملف PDF إلى صيغة MD عبر الموقع:
👉 https://pdf2md.morethan.io/

### 🔹 الخطوة 2: استعمال الأداة

#### عبر سطر الأوامر | CLI

```bash
python md_to_excel.py text.md output.xlsx
```

#### عبر الويب | Web Interface

```bash
python app.py
```

ثم افتح المتصفح:
Then open in your browser:
👉 http://localhost:5000
<img width="1363" height="599" alt="image" src="https://github.com/user-attachments/assets/62267ec8-4e64-43f4-b41b-47f7ae21d133" />

---

## 📁 هيكلة المشروع | Project Structure

```
ExamList_Converter/
│── md_to_excel.py
│── app.py
│── requirements.txt
│── README.md
```

---

## 🧑‍💻 المطور | Developer

تم تطوير هذا المشروع لتسهيل العمل الإداري داخل مديريات التربية ومراكز الامتحانات، خاصة في معالجة قوائم المناداة باللجان.

This project was developed to simplify administrative workflows in education directorates and exam centers, especially for handling committee-based call lists.

---

## 📜 الرخصة | License

هذا المشروع مفتوح المصدر تحت رخصة MIT.
This project is open-source under the MIT License.

---

## ⭐ ملاحظة | Note

إذا أعجبك المشروع، لا تنسَ وضع ⭐ على المستودع!
If you like this project, don’t forget to ⭐ the repository!
