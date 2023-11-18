"""main.py"""

# Import packages
import xlwings as xw
import pandas as pd

# Open the Excel file
wb = xw.Book('config.xlsx')

# Select the sheet
sheet = wb.sheets['Sheet1']



# Read the value of cell A2
value = sheet.range('A2').value

# Print the value
print("The value in cell A2 is:", value)

column = 'A'
for i in range(10):
    value = sheet.range(f'{column}{i}').value
    print(value)



















# Define the range you want to select, for example, A1:C10
range_to_select = 'A1:C10'  # Adjust the range according to your needs

# Read the data into a DataFrame
data = sheet.range(range_to_select).options(pd.DataFrame, header=1, index=False).value

# Close the workbook
wb.close()

# Print the DataFrame
print(data)








def convert_formula_to_function(formula_str):
    def formula_func(**kwargs):
        # Définit localement les variables mentionnées dans la formule
        for key, value in kwargs.items():
            locals()[key] = value
        # Exécute la formule
        return eval(formula_str)
    return formula_func

# Exemple d'utilisation
annual_income_formula = convert_formula_to_function("monthly_income * 12")
tax_rate_formula = convert_formula_to_function("0.3 if annual_income > 50000 else 0.2")

# Exécution des fonctions avec les variables nécessaires
annual_income = annual_income_formula(monthly_income=4000)
tax_rate = tax_rate_formula(annual_income=annual_income)

print("Annual Income:", annual_income)
print("Tax Rate:", tax_rate)
