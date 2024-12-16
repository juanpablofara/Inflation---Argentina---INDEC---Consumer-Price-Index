import requests
import pandas as pd
from bs4 import BeautifulSoup
from io import BytesIO
import re

def find_and_return_xls_link(page_url, reference_text=None):
    """
    Finds and returns the most relevant .xls file link from a webpage.

    Parameters:
        page_url (str): URL of the webpage to scrape for .xls links.
        reference_text (str, optional): A keyword or phrase to prioritize a specific link.

    Returns:
        str or None: The URL of the .xls file if found, otherwise None.
    """
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/114.0.5735.110 Safari/537.36"
    }
    session = requests.Session()
    try:
        # Send a GET request to the webpage
        response = session.get(page_url, headers=headers)
        response.raise_for_status()
        soup = BeautifulSoup(response.content, "html.parser")
        
        # Find all anchor tags with .xls links
        links = soup.find_all("a", href=True)
        xls_links = [link for link in links if link["href"].endswith(".xls")]
        
        # Return None if no links are found
        if not xls_links:
            return None
        
        # Prioritize links containing the reference text if provided
        if reference_text:
            for link in xls_links:
                if reference_text.lower() in link.text.lower():
                    result = link["href"]
                    break
            else:
                result = xls_links[-1]["href"]  # Default to the last link
        else:
            result = xls_links[-1]["href"]
            
        # Ensure the link is an absolute URL
        if not result.startswith("http"):
            base_url = "/".join(page_url.split("/")[:3])
            result = base_url + result

        return result
    except requests.exceptions.RequestException as e:
        #print(f"Error accessing the website: {e}")
        return None


def is_valid_excel(content):
    """
    Validates whether the given content is a valid Excel file.

    Parameters:
        content (bytes): Binary content of the file to validate.

    Returns:
        bool: True if the content is a valid Excel file, False otherwise.
    """
    try:
        with BytesIO(content) as data:
            pd.ExcelFile(data) # Attempt to open the file as an Excel file
        return True
    except Exception:
        return False


def download_excel(url):
    """
    Downloads an Excel file from a given URL.

    Parameters:
        url (str): URL of the file to download.

    Returns:
        bytes or None: Binary content of the file if successfully downloaded, otherwise None.
    """
    try:
        # Send a GET request to the URL
        response = requests.get(url)
        response.raise_for_status()
        return response.content
    except requests.exceptions.RequestException as e:
        #print(f"Error downloading file: {e}")
        return None


def create_empty_dataframe():
    """
    Creates an empty DataFrame with predefined column names.

    Returns:
        pd.DataFrame: An empty DataFrame with columns for date, region, product, unit, and price.
    """
    return pd.DataFrame(columns=["Date", "Region", "Product", "Unit", "Price"])


def find_sheet_case_insensitive(file_content, sheet_name):
    """
    Finds a sheet in an Excel file by name, case-insensitively.

    Parameters:
        file_content (bytes): Binary content of the Excel file.
        sheet_name (str): The name of the sheet to find.

    Returns:
        str: The name of the matched sheet, or the first sheet if no match is found.
    """
    with BytesIO(file_content) as file_data:
        # Load all sheet names from the Excel file
        sheets = pd.ExcelFile(file_data).sheet_names
    for sheet in sheets:
        if sheet.lower() == sheet_name.lower():
            return sheet
        
    # Default to the first sheet if the desired one is not found
    #print(f"Sheet '{sheet_name}' not found. Using the first available sheet.")
    return sheets[0]


def populate_price_column_with_numbers(file_content, sheet_name, df):
    # Read the specified sheet from the file
    with BytesIO(file_content) as file_data:
        excel_data = pd.ExcelFile(file_data)
        actual_sheet_name = find_sheet_case_insensitive(file_content, sheet_name)
        sheet_data = pd.read_excel(excel_data, sheet_name=actual_sheet_name, header=None)
        
    #define all lists
    price_data = []
    region_data = []
    product_data = []
    unit_data = []
    date_data = []
    
    # define region list
    valid_regions = ["GBA", "Pampeana", "Noreste", "Noroeste", "Cuyo", "Patagonia"]
    
    # Dictionary to convert months to numeric format
    month_mapping = {
        "Enero": "01", "Febrero": "02", "Marzo": "03", "Abril": "04",
        "Mayo": "05", "Junio": "06", "Julio": "07", "Agosto": "08",
        "Septiembre": "09", "Octubre": "10", "Noviembre": "11", "Diciembre": "12"
    }
    
    # Keep the last year found
    last_year_found = None
    
    for col_idx, column_data in sheet_data.items():
        for row_idx, value in enumerate(column_data):
            #int and float values greater than 0
            if isinstance(value, (int, float)) and value > 0: 
                price_data.append(value)
                # Find months and years in the same column
                date_found = None
                year_found = None
                for cell_idx, cell_value in enumerate(column_data):
                    if isinstance(cell_value, str):
                        if cell_value in month_mapping:
                            date_found = month_mapping[cell_value]
                        elif re.match(r"(?i)año \d{4}", cell_value):
                            year_found = re.search(r"\d{4}", cell_value).group()
                            # Actualizar el último año encontrado
                            last_year_found = year_found   
                        if date_found and year_found:
                            break
                # Use the last year found if a current one is not found
                if date_found and not year_found:
                    year_found = last_year_found
                    
                # Combine year and month if both are found
                if date_found and year_found:
                    full_date = f"{year_found}-{date_found}-01"
                else:
                    full_date = None
                    
                # Find regions and products in the same row
                row = sheet_data.iloc[row_idx]
                region_found = None
                product_found = None
                for idx, check_value in enumerate(row):
                     if isinstance(check_value, str) and check_value in valid_regions:
                        region_found = check_value
                        
                        # Find the immediate str cell after the region for products
                        for next_value in row[idx + 1:]:
                            if isinstance(next_value, str):
                                product_found = next_value
                                break
                        product_data.append(product_found)
                        # Find the immediate str cell after the region for units
                        for next_value in row[idx + 2:]:
                            if isinstance(next_value, str):
                                product_found = next_value
                                break
                        unit_data.append(product_found) 
                        break
                        
                # If a valid region is not found, take the first cells in the row
                if not region_found:
                    region_found = row.iloc[0] if isinstance(row.iloc[0], str) else None
                    #Take the second cell for the product
                    product_found = row.iloc[1] if len(row) > 1 and isinstance(row.iloc[1], str) else None
                    product_data.append(product_found)
                    #Take the third cell for the units
                    unit_found = row.iloc[2] if len(row) > 2 and isinstance(row.iloc[2], str) else None
                    unit_data.append(product_found)   

                date_data.append(full_date)
                region_data.append(region_found)
                
    # Add the data to the columns 'Price', 'Region', 'Product','Unit', and 'Date'
    if price_data:
        df = df.reindex(range(len(price_data)))
        df["Price"] = [round(price, 2) for price in price_data]
        df["Region"] = region_data
        df["Product"] = product_data
        df["Unit"] = unit_data
        df["Date"] = date_data

    return df

def main():
    file_url = "https://www.indec.gob.ar/ftp/cuadros/economia/sh_ipc_precios_promedio.xls"
    fallback_url = "https://www.indec.gob.ar/Nivel4/Tema/3/5/31"
    reference_text = "Índice de precios al consumidor"

    content = download_excel(file_url)
    
    # Check if the downloaded file is valid
    if content is None or not is_valid_excel(content):
        # Search for an alternative dynamic link in the fallback URL
        dynamic_url = find_and_return_xls_link(fallback_url, reference_text)
        if dynamic_url:
            # If a dynamic link is found, attempt to download the file
            content = download_excel(dynamic_url)
            
        else:
            # If no alternative link is found, terminate the execution
            #print("No dynamic link to the file was found.")
            return

    if content and is_valid_excel(content):
        df = create_empty_dataframe()
        df = populate_price_column_with_numbers(content, "Nacional", df)

    else:
        print("The downloaded content is not a valid Excel file.")
        
if __name__ == "__main__":
    main()