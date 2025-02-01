# Generic Excel Export Script
This script reads JSON file and will generate Excel file from that JSON file.

## 📌 Requirements
- Python3
- Dependencies installed from `./requirements.txt`

### 1️⃣ Install Dependencies
Run the following command to install all required packages:
```bash
pip3 install -r ./requirements.txt
```

## 🚀 How to Run the Script
Execute the script using the following command:
```bash
python3 ./exporter.py /path/to/input.json /path/to/output.xlsx
```

Example:
```bash
python3 ./exporter.py ./sample.json ./sample.xlsx
```

## 📂 Sample JSON Structure
The input JSON file should have the following structure:
```json
{
  "fields": {
    "id": "ID",
    "name": "Name",
    "price": "Price",
    "quantity": "Quantity"
  },
  "data": [
    {
      "sheet": "Products",
      "lists": [
        {
          "id": 1,
          "name": "Product A",
          "price": 19.99,
          "quantity": 100
        },
        {
          "id": 2,
          "name": "Product B",
          "price": 29.99,
          "quantity": 150
        }
      ]
    },
    {
      "sheet": "Sales",
      "lists": [
        {
          "id": 101,
          "name": "Sale A",
          "price": 199.99,
          "quantity": 5
        }
      ]
    }
  ],
  "settings": {
    "currency_fields": ["price"],
    "text_center_fields": ["name"],
    "text_middle_fields": ["quantity"]
  }
}
```

## 🛠 Troubleshooting
- If you get a `ModuleNotFoundError` error, ensure all dependencies are installed using:
```bash
pip3 install -r ./requirements.txt
```
- If the script fails because of permissions error, try running this command:
```bash
chmod +x ./exporter.py
```
- If you get a `No such file or directory` error, ensure the directory exists before running the script.

## ✨ Features
- ✅ Generate excel file from the specified JSON file
- ✅ Define specific fields for each sheet
- 🚧 Settings - Border Customization
- 🚧 Settings - Header Customization
- 🚧 Settings - Custom Start Position (Not only from A1)
- 🚧 Settings - Show/Hide Footer & Footer Customization
- 🚧 Settings - Merge Row/Cell