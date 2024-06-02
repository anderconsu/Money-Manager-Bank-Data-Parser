# Money Manager Bank Data Parser

This proyects aims to get the extractred transference data in xlsx format from your bank and parse it to a .tsv file which you can import in [MoneyManager App](https://www.realbyteapps.com/) 

For now it works with BBK data, but It shouldn't be difficult to adapt it.

## Notes

- The data format on the script is in dd/mm/yyyy, this **must** be changed to the format that you are using on the app, there is a comment with a warning (#!) to find it easier.

- All the Categories and Subcategories **must** match the ones you have on the app.

- A Category **must** be selected always, if you are unsure where to map something, use the 'Other' category, right now is mapped to 'Otros'.





## Usage

- Generate the desired Pricing data (/Pricing) in xlsx format. There are 2 examples you can use as guidance.
  
  - The *general.xlsx* file contains the data necessary to map a transference type to the desired Category and Subcategory for recurring transferences you don't know the exact price of.
  
  - The *eroski.xlsx* file is an example of a specific store prices you can have. The store selection is based on the transference type/procedence.

- You have to instantiate the Data Frame of your generated store xlsx file and add a desired word for selection.
  
  Example:
  
  You create a ikea.xlsx file in Pricing folder. Then below `# * Prepare the data with the pricing` you instantiate the data frame and you add it to the procedence dictionary with the word "IKEA" as the key. This word has to appear in the data extracted for the bank. It's what is used to select the right .xlsx file.
  
  ```python
  ikea_df = pd.read_excel("Pricing/ikea.xlsx")
  procedence_dict = {
      "EROSKI": eroski_df,
      "IKEA": ikea_df,
  }
  ```

- You can add transference types you want to skip in deny_procedence list

- Set the desired account (**you must have it in MoneyManager**)

- Run the script, it will generate a .tsv file that you can import in the app  under Results folder with todays date as a name.
  
  






