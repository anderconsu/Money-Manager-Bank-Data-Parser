import pandas as pd
from os import devnull
import xlrd
from datetime import datetime


# *: Read the XLSX template file to get the column names
template_file = "plantilla.xlsx"
template_df = pd.read_excel(template_file)
template_columns = template_df.columns.tolist()


# * Read the XLS file to get the data you want to parse
data_file = "Files/05-2024.xls"
# data_df = pd.read_excel(data_file)
wb = xlrd.open_workbook(data_file, logfile=open(devnull, "w"))
data_df = pd.read_excel(wb, dtype=str, skiprows=7, engine="xlrd")


# * Prepare the data for parsing
# Name the columns of the data_df DataFrame with the first row of the data_df DataFrame
data_df.columns = ["fecha", "concepto", "fecha valor", "importe", "saldo"]


# Try to convert the first column to a date
try:
    # Convert fecha column with format "%d/%m/%Y"
    data_df["fecha"] = pd.to_datetime(data_df["fecha"], format="%d/%m/%Y")

    # Convert fecha valor column with format "%d/%m/%Y"
    data_df["fecha valor"] = pd.to_datetime(data_df["fecha valor"], format="%d/%m/%Y")


except ValueError as e:
    # Print the error message
    print(e)
    # Exit the program
    exit()


# Try to convert the importe column to a float
try:
    # Convert importe column to float
    data_df["importe"] = data_df["importe"].str.replace(",", "").astype(float)

    # Print the first 5 rows of the DataFrame
    print(data_df.head())
except ValueError as e:
    # Print the error message
    print(e)
    # Exit the program
    exit()


# * Prepare the data with the pricing and categorizing logic
# Get different pricing data with the same structure ['Amount', 'Note', 'Category', 'Subcategory']
eroski_df = pd.read_excel("Pricing/eroski.xlsx")


# Get general pricing data for general categories ['Procedence', 'Category', 'Subcategory']
general_pricing_df = pd.read_excel("Pricing/general.xlsx")


# Create a dictionary with the different pricing dataframes
procedence_dict = {
    "EROSKI": eroski_df,
}


# Deny list to avoid periodic payments already congifured in app
deny_procedence = ["GITHUB", "DIGITEAL", "NOMINA"]


# * Create a new DataFrame to store the parsed data
parsed_data = pd.DataFrame(columns=template_columns)
account = "K26"

# * Parse the data
for index, row in data_df.iterrows():
    # Check the importe column to find matches with the pricing data
    procedence = row["concepto"]

    # Check if one of the words in procedence is in the deny_procedence list
    if any(word in procedence for word in deny_procedence):
        continue

    type = "Expense" if row["importe"] < 0 else "Income"
    #! date in format dd/mm/yyyy its important to match the format on the app
    date = row["fecha"].strftime("%d/%m/%Y")
    amount = abs(row["importe"])
    note = procedence
    category = "Otros"
    subcategory = None

    selected_pricing_df = None

    # Check the procedence of the row
    for word in procedence_dict.keys():
        if word in procedence:
            selected_pricing_df = procedence_dict[word]
            break

    if selected_pricing_df is not None and type == "Expense":
        # Check if the amount is in the pricing data
        if amount in selected_pricing_df["Amount"].values:
            # Get the row of the pricing data that matches the amount
            pricing_row = eroski_df[eroski_df["Amount"] == abs(amount)]
            # Get the note, category and subcategory of the pricing data
            note = pricing_row["Note"].values[0] if pricing_row["Note"].any() else None
            category = (
                pricing_row["Category"].values[0]
                if pricing_row["Category"].any()
                else None
            )
            subcategory = (
                pricing_row["Subcategory"].values[0]
                if pricing_row["Subcategory"].any()
                else None
            )
    else:
        # Check if the procedence is in the general pricing data
        if any(
            word.lower() in procedence.lower()
            for word in general_pricing_df["Procedence"].values
        ):
            # Get the row of the general pricing data that matches a word of the procedence
            pricing_row = general_pricing_df[
                general_pricing_df["Procedence"]
                .str.lower()
                .isin([word.lower() for word in procedence.split()])
            ]
            # Get the category and subcategory of the general pricing data
            category = (
                pricing_row["Category"].values[0]
                if pricing_row["Category"].any()
                else "Otros"
            )
            subcategory = (
                pricing_row["Subcategory"].values[0]
                if pricing_row["Subcategory"].any()
                else None
            )
            note = (
                pricing_row["Procedence"].values[0]
                if pricing_row["Procedence"].any()
                else procedence
            )

            # Insert the parsed row into the parsed_data DataFrame
    new_row = [date, account, category, subcategory, note, amount, type, None]
    parsed_data.loc[len(parsed_data)] = new_row


# * Export as tsv file

# Get the current date
current_date = datetime.now()

# Convert the current date to a string
date_string = current_date.strftime("%d-%m-%Y")

parsed_data.to_csv(f"Results/{date_string}.tsv", sep="\t", index=False)
