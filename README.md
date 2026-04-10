# 🚀 Excelisor POC – AI-Powered Data Analysis Engine

An intelligent full-stack application that enables users to query and analyze tabular/Excel data using natural language.

Built using FastAPI and LLM-driven planning, the system translates user queries into executable analytical workflows, performs data processing, and returns structured insights.

---

## 🎯 Key Highlights

- 🧠 Natural Language → Data Analysis (LLM-powered)
- ⚙️ Multi-step execution planning engine
- 📊 Automated data aggregation & transformation
- 🔍 Built-in data auditing (missing values, duplicates, outliers)
- 🧱 Full-stack architecture (backend + frontend)
- 🚀 Designed for scalability and extensibility

---

## 🧠 Core Problem Solved

Non-technical users struggle to analyze data using traditional tools.

This system bridges that gap by:
➡️ Converting plain English queries into structured data operations  
➡️ Automating analysis without requiring Excel formulas or SQL  

---

## ⚙️ System Architecture

User Query (Natural Language)
        ↓
Intent Classification + Entity Extraction
        ↓
Execution Plan Generation (LLM)
        ↓
Python Execution Engine (Pandas)
        ↓
Processed Data Output
        ↓
Frontend Display

---

## 🧱 Project Structure

excelisor-poc/
├── backend/      # FastAPI + AI planning + execution engine
├── frontend/     # User interface
├── .gitignore
├── README.md

---

## ⚙️ Tech Stack

Backend:
- FastAPI (API layer)
- Pandas (data processing)
- OpenAI API (LLM reasoning)
- Python

Frontend:
- (Update with your framework: React / JS / etc.)

---

## 🚀 Features

### 🔹 AI Query Engine
- Converts natural language into structured data operations
- Handles aggregation, filtering, ranking, and comparisons

### 🔹 Data Audit Module
- Detects:
  - Missing values
  - Duplicate records
  - Outliers (IQR method)
- Provides data quality insights

### 🔹 Execution Planner
- Generates step-by-step analytical plans
- Supports multi-step transformations

### 🔹 Scalable Design
- Modular architecture (planner, executor, audit)
- Easily extendable for new analytics features

---

## 🔍 Example Queries

- "Show total sales by region"
- "Top 5 products by revenue"
- "Audit this dataset"
- "Compare average sales across categories"

---

## 🖥️ Setup Instructions

### 1. Clone Repository

git clone https://github.com/harshBABU/excelisor-poc.git
cd excelisor-poc

---

### 2. Backend Setup

cd backend
python -m venv venv
venv\Scripts\activate

pip install -r requirements.txt

---

### 3. Environment Variables

Create `.env` inside backend/:

OPENAI_API_KEY=your_openai_api_key

---

### 4. Run Backend

uvicorn main:app --reload

Access API docs:
http://localhost:8000/docs

---

### 5. Frontend Setup

cd frontend
npm install
npm run start

---

## 📌 Engineering Highlights

- Designed a **multi-stage AI pipeline** (intent → plan → execution)
- Implemented **robust data audit system** for real-world datasets
- Built **modular architecture** enabling independent scaling of components
- Handled **large datasets (~300K+ rows)** efficiently using Pandas

---

## 🚀 Future Enhancements

- Visualization layer (charts, dashboards)
- Real-time data connectors
- Advanced query understanding
- Multi-agent architecture

---

## 👨‍💻 Author

Harsh Agarwal

---

## 💡 Why This Project Stands Out

This project demonstrates:
✔ Applied AI in real-world data problems  
✔ Backend system design + data engineering  
✔ Full-stack development capability  
✔ Ability to translate business problems into technical solutions  

---

## 📜 License

For learning and demonstration purposes.
