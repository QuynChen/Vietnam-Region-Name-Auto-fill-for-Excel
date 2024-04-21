# import essential libraries
import pandas as pd
import re

excel_path = 'data/vung-khu_vuc.xlsx'
# read excel file
df = pd.read_excel(excel_path)

# read txt file
txt_path = 'data/vung.txt'
with open(txt_path, 'r', encoding='utf-16') as file:
    txt_file = file.read()

# preprocessing vung.txt file    
txt_file = re.sub('và', ',', txt_file)   # replace "và" with ","
txt_file = re.sub(' -', '', txt_file)   # remove "-"

# Extract necessary data from vung.txt file using regex
pattern = r'Vùng\s([\w\s]+)\b\sgồm\s[\w\s]+\:\s([\w\s,]+)(?=\.|\Z)'
matches = re.findall(pattern, txt_file)

# Save the results to a dictionary
province_regions = {} 
for match in matches:
    region = match[0] 
    provinces = [province.strip() for province in match[1].split(',')]
    province_regions[region] = provinces

# add special value (For provinces and cities with long names, people tend to abbreviate them)
province_regions['Đông Nam Bộ'].append('HCM')
province_regions['Đông Nam Bộ'].append('BRVT')

def check_region(row):
    """Function to check region of a given province or city
    Args:
        row (string): a province or city name
    Returns:
        string: region or empty if not found
    """
    # Check if row is int or float
    if isinstance(row, int) or isinstance(row, float):
        return '' # Return empty string if it is not a valid input
    
    # Normalize row
    normalized_row = row.lower() # convert to lower case
    normalized_row = re.sub('[-,\._;]', '', normalized_row) # remove special characters (,.-_;)

    # Iterate through dictionary 'province_regions'
    for region, locations in province_regions.items():
        # Normalize locations
        normalized_locations = [location.lower() for location in locations] # convert all lines to lower case
        
        for location in normalized_locations:
            if location in normalized_row:
                return region
                break #  Exit loop if found the region (because each row only belongs to one region)
            
    # If no region found, return empty string
    return ''

def main():
    for index, row in df.iterrows():
        region = check_region(row['Huyện, Tỉnh'])
        df.loc[index, 'Vùng'] = region # Add results to the corresponding column
    #  Save result to xlsx file
    df.to_excel('data/result.xlsx', index=False)
    
if __name__ == "__main__":
    main()
 
 