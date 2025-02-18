# Excel VBA-Based Invoice & Release Order (RO) Management System

This project is an **Excel VBA-based automation system** designed to simplify **invoice and release order (RO) generation** for *Viplav Advertising Co.* It enables users to efficiently **create, edit, and manage invoices and ROs** using Excel and Word integration, making the process **faster and error-free**.

## 🚀 Features
✅ **Automated GST Invoice & RO Generation** – Eliminates manual entry, reducing errors.  
✅ **User-Friendly Forms** – Instead of manually filling cells, users enter data through a user-friendly form.  
✅ **Word Document Integration** – Automatically populates a pre-formatted Word template with the entered details.  
✅ **Multi-Item Support** – Allows multiple advertisements per invoice.  
✅ **Auto-Calculations** – Includes GST, discounts, total payable amount, and conversion to words.  
✅ **Search & Edit Functionality** – Easily retrieve invoices using Invoice No. or Client Name.  
✅ **PDF Export** – Save invoices and ROs as PDFs for easy sharing.  

## 📂 Download Files
📥 **Excel File (RO Record):** [Download rorecord.xlsm](https://github.com/viplavs2004/Ro_Automation/blob/main/rorecord.xlsm)  
📥 **Word Template (RO Template):** [Download ROtemplate.docm](https://github.com/viplavs2004/Ro_Automation/blob/main/ROtemplate.docm)  

## 🛠️ How to Use the Project

### **Step 1: Enable Macros**
Since this project uses **VBA macros**, you must enable them for automation to work.

1. Open **rorecord.xlsm** in Excel.
2. If you see a **Security Warning (Macros have been disabled)**, click **Enable Content**.
3. Ensure macros are enabled:
   - Go to **File > Options > Trust Center**.
   - Click on **Trust Center Settings**.
   - Go to **Macro Settings** and select **Enable all macros**.
   - Click **OK** and restart Excel.

### **Step 2: Generate a Release Order (RO)**
1. Open **rorecord.xlsm**.
2. Click on the **Generate RO** button.
   ![Generate RO Button](https://github.com/viplavs2004/Ro_Automation/blob/main/button_Click_generate_ro%20.png)
3. A **UserForm** will appear where you can enter advertisement details.
   ![Generate RO Form](https://github.com/viplavs2004/Ro_Automation/blob/main/genenarate_ro_form.png)
4. Fill in all required details and click **Submit** to generate the RO.
5. The data is saved in Excel, and a Word document is generated automatically.
   ![Auto-Generated RO](https://github.com/viplavs2004/Ro_Automation/blob/main/auto_generated_ro.png)

### **Step 3: Save as PDF**
1. Click the **Save as PDF** button in the Excel form.
   ![Save PDF Button](https://github.com/viplavs2004/Ro_Automation/blob/main/save_pdf_button.png)
2. The RO will be saved as a **PDF file** for easy sharing.

### **Step 4: Search and Edit an Existing RO**
1. Click on the **Search RO** button.
   ![Search RO UserForm](https://github.com/viplavs2004/Ro_Automation/blob/main/search_ro_userform.png)
2. Select an RO from the list based on **Invoice No. or Client Name**.
   ![Select RO for Edit](https://github.com/viplavs2004/Ro_Automation/blob/main/select_ro_for_edit.png)
3. The existing details will load in the form.
4. Modify the details and click **Update** to save changes.

## 🔧 Troubleshooting

### 1️⃣ **Macros Not Working?**
- Ensure macros are **enabled** in Excel (**See Step 1: Enable Macros**).
- If macros are disabled, go to **File > Options > Trust Center > Macro Settings** and enable macros.

### 2️⃣ **Word Template Not Opening?**
- Ensure **ROtemplate.docm** is in the **same folder** as **rorecord.xlsm**.
- If Excel prompts a warning about linked files, **allow access**.

### 3️⃣ **Errors in Data Entry?**
- Check that **all required fields** in the UserForm are filled.
- Ensure that **numeric values** (like price) are entered correctly.

## 📜 License
This project is developed for **Viplav Advertising Co.** and is free to use for learning purposes. If you wish to customize or extend its functionality, feel free to **fork the repository**!

## 🤝 Contributing
If you’d like to contribute:
1. **Fork** the repository.
2. Create a new branch (**git checkout -b feature-branch**).
3. Commit your changes (**git commit -m 'Add new feature'**).
4. Push to the branch (**git push origin feature-branch**).
5. Create a **Pull Request**.


