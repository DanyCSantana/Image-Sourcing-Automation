# 🎬 Python Tool for Sourcing and Organizing Movie Images

**File:** `ImageManagementScript.py`  
**Category:** Portfolio Project – Process Automation for Inflight Entertainment  
**Role:** Creator & Sole Developer  
**Technologies:** Python, pandas, openpyxl, pathlib, tqdm, logging  

---

## 📌 Objective

Efficiently source poster and still images requested by airlines by searching a large database of nearly **30,000 images** spread across **800+ distributor folders for movies and TV, Posters and Stills**, using Python to automate what used to be a fully manual process. This tool reduced the workload by over **90%**, saving hours of work each cycle update.

---

## ⚙️ What the Script Does

- 🔍 **Recursively searches** through 300+ distributor folders to locate relevant images based on titles listed in an Excel tracker  
- 🧠 Uses **intelligent string matching** (regex and fuzzy logic) to increase accuracy even when filenames vary  
- 📂 **Copies and renames** matched images into a clean, structured output folder (separated by posters and stills)  
- 📊 **Updates the Excel sheet** with the results for tracking and documentation  
- 📭 **Generates missing image reports**, including ready-to-send email drafts for contacting distributors  
- 📈 Includes **logging and progress tracking** with full traceability

---

## 🛠️ Skills Demonstrated

- Python scripting for automation  
- File system operations and optimization  
- Data manipulation with `pandas` and `openpyxl`  
- Regex for pattern recognition  
- Clean and structured logging  
- User-friendly feedback with `tqdm` progress bars  
- Workflow thinking: connecting tools, users, and deliverables  

---

## 📉 Before vs After

| Task                        | Manual Process | With Python Script |
|-----------------------------|----------------|---------------------|
| Finding all required images | 2–3 hours      | < 10 minutes        |
| Folder navigation           | Tedious        | Fully automated     |
| Updating Excel tracker      | Manual         | Auto-filled         |
| Emailing for missing items  | Time-consuming | Draft generated     |
| Error tracking              | Prone to loss  | Logged & traceable  |

---

## 🧩 Real-World Context

This tool was built to solve a real challenge in the inflight entertainment industry: organizing high volumes of image assets (posters and stills) provided by over 300 distributors. The script supports the metadata and delivery process by making sourcing visual assets significantly faster and more reliable.

---

## 🚀 Impact

- 📁 Automated processing of ~30,000 images across 800+ folders  
- ⏱️ Reduced image retrieval time from hours to minutes  
- 💼 Created a scalable solution used monthly across airline projects  
- 📎 Made follow-ups with content providers faster and more consistent  

---
