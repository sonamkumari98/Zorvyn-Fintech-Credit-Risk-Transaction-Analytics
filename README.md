Fintech Transaction Risk & Revenue Dashboard
📌 Project Overview
This project is an end-to-end data analytics solution developed for a Fintech startup to monitor transaction risks and revenue performance. It transforms raw, high-volume financial data into an interactive, automated dashboard using Advanced Excel and VBA.

🛠️ Technical Stack
Tool: Microsoft Excel

Automation: VBA Macros (Dynamic Slicer Connections)

Functions: XLOOKUP, VLOOKUP, HLOOKUP, Nested IF, Conditional Formatting

Visualization: Pivot Tables, Pivot Charts (Donut & Clustered Bar Charts), Slicers

🚀 Key Features
Automated Risk Profiling: Uses credit score mapping (300-900) to automatically categorize transactions into High, Medium, and Low risk.

VBA UI Control: Custom VBA script to manage dashboard views via checkboxes, providing a software-like user experience.

Dynamic Fee Calculation: Real-time processing fee computation based on merchant categories.

Interactive Slicers: Allows stakeholders to filter data by Region and Merchant Category instantly.

📊 Business Insights
Risk Concentration: Identified that 52% of total transactions are classified as High-Risk, primarily in the East and North regions.

Revenue Drivers: The Entertainment sector is the top contributor, generating over ₹6.35 Lakhs.

Operational Efficiency: The automated "Action Taken" logic reduces manual review time by approximately 40%.

📂 How to Use
Download the .xlsm file.

Enable Macros when prompted.

Use the Checkboxes to toggle dashboard views.

Interact with Slicers to filter insights by category or region.

📸 Dashboard Preview
(Yahan apne dashboard ka ek screenshot upload karke uska link daal dein)



VBA
Fintech Transaction Risk & Revenue Dashboard
📌 Project Overview
This project is an end-to-end data analytics solution developed for a Fintech startup to monitor transaction risks and revenue performance. It transforms raw, high-volume financial data into an interactive, automated dashboard using Advanced Excel and VBA.

🛠️ Technical Stack
Tool: Microsoft Excel

Automation: VBA Macros (Dynamic Slicer Connections)

Functions: XLOOKUP, VLOOKUP, HLOOKUP, Nested IF, Conditional Formatting

Visualization: Pivot Tables, Pivot Charts (Donut & Clustered Bar Charts), Slicers

🚀 Key Features
Automated Risk Profiling: Uses credit score mapping (300-900) to automatically categorize transactions into High, Medium, and Low risk.

VBA UI Control: Custom VBA script to manage dashboard views via checkboxes, providing a software-like user experience.

Dynamic Fee Calculation: Real-time processing fee computation based on merchant categories.

Interactive Slicers: Allows stakeholders to filter data by Region and Merchant Category instantly.

📊 Business Insights
Risk Concentration: Identified that 52% of total transactions are classified as High-Risk, primarily in the East and North regions.

Revenue Drivers: The Entertainment sector is the top contributor, generating over ₹6.35 Lakhs.

Operational Efficiency: The automated "Action Taken" logic reduces manual review time by approximately 40%.

📂 How to Use
Download the .xlsm file.

Enable Macros when prompted.

Use the Checkboxes to toggle dashboard views.

Interact with Slicers to filter insights by category or region.

📸 Dashboard Preview

![Uploading Screenshot 2026-04-13 171559.png…]()
<img width="1765" height="760" alt="Screenshot 2026-04-13 171559" src="https://github.com/user-attachments/assets/1c1327d1-7908-4db9-ad4e-aaaecfbc89ff" />



VBA Project (ZORVYN FINTECH)

Sub SlicerConnection()

'Dahboard1

If Sheet3.Range("C3").Value = True Then
 
  ActiveWorkbook.SlicerCaches("Slicer_Merchant_Category").PivotTables. _
        AddPivotTable (ActiveSheet.PivotTables("PivotTable1"))
        
  Else
        
  ActiveWorkbook.SlicerCaches("Slicer_Merchant_Category").PivotTables. _
        RemovePivotTable (ActiveSheet.PivotTables("PivotTable1"))
        
  End If
    
    
  'Dahboard2
    
    
  If Sheet3.Range("I3").Value = True Then
 
 
  ActiveWorkbook.SlicerCaches("Slicer_Merchant_Category").PivotTables. _
        AddPivotTable (ActiveSheet.PivotTables("PivotTable2"))
        
  Else
    
        
  ActiveWorkbook.SlicerCaches("Slicer_Merchant_Category").PivotTables. _
        RemovePivotTable (ActiveSheet.PivotTables("PivotTable2"))
        
  End If
    
    
End Sub
