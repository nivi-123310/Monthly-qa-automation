**📊 Automated QA Validation System for Media Campaign Data**
🚀 Overview
This project is a production-grade automated QA system built using Python and Pandas to validate large-scale media campaign datasets across multiple markets, channels, brands, and products.
It eliminates manual QA effort and ensures data accuracy, consistency, and compliance with business rules before reporting.

**🎯 Why This Project Matters**
Manual QA of campaign data is:


Time-consuming (hours per market)


Error-prone


Not scalable


**This solution:**


Automates validation across multiple markets


Detects data inconsistencies instantly


Reduces QA time from hours → minutes


Ensures high data reliability for dashboards & reporting



**⚙️ What the Script Does**
**1. 📂 Reads Data**


Scans folder structure:
Market → Channel → Excel Files


Reads all .xlsx files dynamically



**2. 🔍 Filters Relevant Data**


Processes only:


Target Month & Year


OR campaigns overlapping target period





**3. ✅ Performs QA Checks**
🔹 Mandatory Field Validation


Flags missing required columns like:


Market, Channel, Campaign, Spend, Dates





**🔹 Data Type Validation**


Detects:


Invalid dates


Non-numeric values


Incorrect text formats





**🔹 Date Consistency Checks**


Ensures:


Date falls within date_start and date_end




Flags:
Date outside campaign duration



**🔹 Spend Validation**


Checks:


Net vs Gross logic


Zero spend issues




**Detects incorrect formats like:**
European format → 1.234,56 ❌Expected format → 1,234.56 ✅



**🔹 Media-Specific Checks (TV)**


Validates:


GRPs fields


Decimal format issues (e.g., 15,5 ❌)





**🔹 Numeric Integrity Checks**


Ensures fields like impressions, clicks are:


Numeric


Whole numbers (where required)





**🔹 Brand vs Product Mapping**


Validates:
Product belongs to correct Brand



🔹 Campaign Split Detection
Groups data by:
Campaign + Media Sub Channel + Brand + Product
Detects split types:


Daily


Weekly


Monthly


Partial


Irregular


Flags unexpected splits automatically.

**🧠 How Campaign Split Logic Works**
The script:


Collects all campaign dates


Calculates:


Duration (start → end)


Number of actual data points




Compares coverage:


Full → Daily


Gaps → Weekly / Monthly / Irregular





**📤 Output Generated**
1. 📋 Overview Sheet
MarketChannelMonthStatusRows ScannedFranceTVJan 2026Pass/Fail1200

2. 🚨 Issues Sheet
MarketChannelCampaignIssue TypeDetailsGermanyDigitalCampaign AMissing DataNet SpendUKTVCampaign BInvalid FormatEU Spend Format

**📁 Sample Data Included**
This repo contains:


Sample input files (multi-market structure)


Sample QA output report



**🛠️ Tech Stack**


Python


Pandas


NumPy


Excel (XlsxWriter)



**▶️ How to Run**
pip install pandas python-dateutil xlsxwriter
Update paths in script:
BASE_PATH = "your_input_folder"OUTPUT_PATH = "your_output_folder"
Run:
python qa_script.py

⚡ Performance


Processes multiple markets automatically


Handles large datasets efficiently


Reduces QA runtime significantly



****📈 Business Impact**

⏱️ 90% reduction in manual QA time


📊 Improved data accuracy for reporting


🚀 Scalable across markets & channels


🔍 Early detection of critical data issues



**🔒 Note**


Sample data is anonymized


No confidential business data included



**💡 Future Improvements**


Dashboard integration (Power BI)


Email alerts for QA failures


Real-time validation pipeline





