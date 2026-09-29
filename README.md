# 🚀 AI-Infused Executive Reporting & Analytics Suite

[![Google Apps Script](https://img.shields.io/badge/Google%20Apps%20Script-4285F4?style=flat-square&logo=google&logoColor=white)](https://developers.google.com/apps-script)
[![Google Gemini API](https://img.shields.io/badge/Gemini%203.8%20Flash-8E75C2?style=flat-square&logo=googlegemini&logoColor=white)](https://ai.google.dev/)
[![Google Sheets](https://img.shields.io/badge/Google%20Sheets-34A853?style=flat-square&logo=googlesheets&logoColor=white)](https://sheets.google.com/)
[![clasp](https://img.shields.io/badge/Deployed%20with-clasp-blue?style=flat-square)](https://github.com/google/clasp)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg?style=flat-square)](https://opensource.org/licenses/MIT)

An enterprise-grade business intelligence and financial reporting pipeline built inside Google Sheets. This suite bridges transactional data pipelines with Google's latest **Gemini 3.8 Flash Interactions API** via Google Apps Script, automating data hygiene, performance digest computation, forensic margin auditing, executive brief generation, and real-time conversational analysis.
---

## 📌 Executive Summary & Architecture
In traditional workflows, turning raw transactional datasets into leadership-level insights takes hours of manual filtering, aggregation, memorandum writing, and email distribution. 

This project completely automates that lifecycle into a **4-stage intelligent reporting pipeline**:
[ Raw Transaction Ingestion ] (15,000+ line items: POS, ERP, Web)
│
▼
[ Metric Aggregation & Hygiene ] ➔ Top-line KPIs, Regional Splits, Margin Loss Logs
│
▼
[ Gemini 3.8 Flash AI Layer ] ➔ Strategic CFO Synthesis & Forensic Audit
│
▼
[ Multi-Channel Delivery ] ➔ In-Sheet Memo (E3) + HTML Email Briefing + In-Sheet Chat Sidebar
---

## ✨ Key Features & Script Breakdown
| File | Module | Description |
| :--- | :--- | :--- |
| `05_AiCommentary.js` | **CFO Executive Synthesis** | Extracts top-line KPIs, regional breakdowns, and top product performers from `proj3_exec_summary_digest`. Calls the Gemini 3.8 Flash Interactions API to generate a CFO-level memorandum directly into cell `E3`. |
| `06_EmailBriefing.js` | **Executive Email Dispatch** | Transforms live dashboard metrics and AI commentary into a responsive HTML briefing template and delivers it to stakeholders via `GmailApp`. |
| `06_LossAuditAI.js` | **Forensic Loss & Margin Auditor** | Scans 1,546 unprofitable order records in `proj1_audit_loss_log`, aggregates top margin leakage drivers by product category, channel, and country, and prompts Gemini to prescribe operational remediation strategies. |
| `06_AIChat.js` & `06_AIChatModal.html` | **"Ask Your Data" In-Sheet Chatbot** | Implements an interactive custom HTML sidebar inside Google Sheets. Allows leadership to query real-time numbers in natural language. |
| `Menu.js` | **Enterprise Menu Integration** | Binds native UI controls into the Google Sheets toolbar under custom actions. |
---

## 🛠️ Technology Stack
- **Platform:** Google Sheets, Google Apps Script (V8 Runtime)
- **AI Model:** Google Gemini 3.8 Flash (`v1beta/interactions`)
- **Developer Tooling:** `@google/clasp` (Google Command Line Apps Script), Git, Node.js
- **Google Services Used:**
  - `PropertiesService` (Encrypted credential management)
  - `UrlFetchApp` (RESTful API interaction)
  - `GmailApp` (Automated HTML email dispatch)
  - `HtmlService` (Client-side sidebar UI & bi-directional RPC)
  - `SpreadsheetApp` (Spreadsheet DOM manipulation)
---

## 📁 Repository Structure
ai-executive-reporting-sheets/

├── .clasp.json             # clasp configuration mapping local code to Apps Script project ID
├── .gitignore              # Protects sensitive files, credentials, and clasp tokens
├── appsscript.json         # Apps Script manifest and OAuth scope definitions
│
├── Menu.js                 # Toolbar menu initialization (onOpen hook)
├── 05_AiCommentary.js      # Core Gemini API prompt engineering and cell rendering
├── 06_EmailBriefing.js     # Responsive HTML email compilation and dispatch
├── 06_LossAuditAI.js       # Margin leakage aggregation and forensic prompt logic
├── 06_AIChat.js            # Server-side API handler for chatbot queries
├── 06_AIChatModal.html     # Client-side UI stylesheet, markup, and JavaScript for sidebar
│
└── README.md               # Project architecture and technical documentation

🔐 Setup & Installation
1. Prerequisites
Node.js installed on your machine.
A free Google Gemini API key from Google AI Studio.
Google Apps Script API enabled in your User Settings.

2. Clone the Repository
git clone https://github.com/dansierrasam79/ai-executive-reporting-sheets.git
cd ai-executive-reporting-sheets

3. Install clasp & Authenticate
npm install -g @google/clasp
clasp login
4. Link Your Google Sheet
Update .clasp.json with your spreadsheet's unique Script ID (found under Project Settings ⚙️ > IDs > Script ID in Apps Script):
{
  "scriptId": "YOUR_APPS_SCRIPT_ID_HERE",
  "rootDir": "."
}
Push local files to Google Drive:
clasp push

5. Secure Credential Configuration
To follow enterprise security standards, the API key is never hardcoded in source files.
Instead, configure it via Script Properties:
In the Google Apps Script editor, navigate to Project Settings (gear icon ⚙️).
Scroll to Script Properties and click Add script property.
Set Property to: GEMINI_API_KEY
Set Value to: Your_Gemini_API_Key_Here
Click Save script properties.

📊 Sample AI Output (Executive Memorandum)
Below is an excerpt of the strategic analysis generated by 05_AiCommentary.js and rendered in cell E3:> ### **MEMORANDUM**
>
> **TO:** Executive Leadership Team
> **FROM:** Office of the Chief Financial Officer
> **SUBJECT:** Executive Financial Performance Commentary & Strategic Outlook
>
> **1. Executive Summary & Top-Line Performance**
> Total Net Revenue reached **$55.39M**, delivering Net Profit of **$35.59M** (an elite **64.25% net profit margin**). Cost of Goods Sold settled at $19.80M (35.75% of revenue), confirming that supply chain controls have insulated unit economics against inflationary pressures.
>
> **2. Regional Dynamics & Profitability**
>
> - **North America:** Core engine contributing $32.78M in revenue (59.2% share) at a 64.50% profit margin.
> - **Asia:** Highest-margin territory at **65.71%**, outpacing the global benchmark by +146 bps.
> - **Europe:** Generated $10.20M in revenue, trailing corporate margin benchmarks by \~258 bps (61.67%), indicating localized fulfillment or promotional friction.
>
> **3. Product Concentration & Strategic Directives**
> Revenue concentration risk is minimal: top 5 revenue-generating SKUs account for just 3.58% ($1.98M) of global volume, confirming resilient catalog diversification. Directives include increasing discretionary marketing by 15% in Asia to capture high-margin demand and executing a fulfillment margin review across Europe.
🛡️ Security & Best Practices
Zero API Key Leakage: Keys are stored exclusively in Google Cloud's PropertiesService environment and ignored by Git via .gitignore.
Stateless Execution: All transactional aggregation happens in-memory; no client data is persisted externally.
Fail-Safe Exception Handling: API calls use muteHttpExceptions: true with error-parsing wrappers to prevent spreadsheet execution crashes.

👤 Author
Daniel Chakraborty

📄 License
This project is licensed under the MIT License - see the LICENSE file for details.
