import os
import pandas as pd
from bs4 import BeautifulSoup

# Define a function to parse each HTML file and extract the required data
def parse_html(file_path):
    with open(file_path, 'r', encoding='utf-8') as file:
        soup = BeautifulSoup(file, 'html.parser')
        
        # Extract the relevant data points
        data = {}
        data['Project Name'] = soup.find(text="Project Name :").find_next('span').text.strip()
        data['Project Owner'] = soup.find(text="Project Owner :").find_next('span').text.strip()
        data['Prime or Subconsultant Role'] = soup.find(text="Role In Project :").find_next('span').text.strip()
        data['% of Work your firm completed'] = soup.find(text="Percent :").find_next('span').text.strip()
        data['Year Completed (Within Last 12 Years)'] = soup.find(text="Completed :").find_next('span').text.strip()
        data['Work Class'] = soup.find(text="Work class :").find_next('span').text.strip()
        data['Highway Complexity'] = soup.find(text="Complexity :").find_next('span').text.strip()
        data['Bridge Complexity'] = ''  # Set as empty if not present
        data['Total Highway Cost ($)'] = ''  # Set as empty if not present
        data['Total Bridge Cost ($)'] = ''  # Set as empty if not present

        # Special factors (yes/no fields)
        special_factors = {
            'L1. High Traffic Volumes (>50,000 AADT)': 'No',
            'L2. NYC Metropolitan Area': 'No',
            'L3. Large Urban Area': 'No',
            'C1. Night Work': 'No',
            'C2. Substantial Community Liaison Work Due to Controversial Nature of Project': 'No',
            'C3. Extensive Underground and Aerial Utility Relocations, Requiring Ongoing Coordination with Utilities': 'No',
            'C4. Extensive Work on and Coordination with Railroads or Urban Commuter Rail': 'No',
            'M1. Complicated M Staging, with Extensive Field Changes': 'No',
            'M2. Staging With Movable Concrete Barrier': 'No',
            'M3. Installation and Removal of Temporary Steel Bridges': 'No',
            'H1. > 2,500 Sq. Meters of New Full-Depth PCC Pavement': 'No',
            'H2. > 50,000 Metric Tons of ACC Pavement': 'No',
            'B1. Rehab or Replacement of Viaducts, Major Interchanges or Trusses': 'No',
            'B2. Rehab or Replacement of Movable Bridges': 'No',
            'B3. Painting With Class A Containment': 'No',
            'A1. 10 or More Traffic Signals': 'No',
            'A2. 10 or More Interconnected Traffic Signals': 'No',
            'A3. > .62 Miles of Highway or Interchange Lighting': 'No',
            'A4. ITS': 'No',
            'A5. 50 or More Permanent Signs': 'No',
            'A6. > 1.24 Miles of Guiderail': 'No',
            'A7. > 1.24 Miles of Fencing': 'No',
            'A8. > 1.24 Miles of New Closed Drainage Systems': 'No',
            'O1. Building Construction': 'No',
            'O2. Rest Areas': 'No',
            'O3. Hazardous Waste Remediation/Removal': 'No',
            'O4. Underwater Inspection': 'No',
            'O5. Marine Work': 'No'
        }
        
        # Check each factor by its label and mark it "Yes" if present in the HTML file
        for factor in special_factors:
            if soup.find(text=factor):
                special_factors[factor] = 'Yes'

        # Add the special factors to the data dictionary
        data.update(special_factors)

        return data

# Define a function to iterate over multiple HTML files and build the table
def build_table(file_paths):
    records = []
    
    for file_path in file_paths:
        record = parse_html(file_path)
        records.append(record)
    
    # Create a DataFrame from the records
    df = pd.DataFrame(records)
    
    # Export to Excel or allow copying the table
    df.to_excel('output_table.xlsx', index=False)
    return df

# Usage
html_files = [
    "/mnt/data/rwservlet.html",  # Add all the HTML file paths here
    "/mnt/data/rwservlet1.html",
    "/mnt/data/rwservlet5.html",
    "/mnt/data/rwservlet3.html"
]
table = build_table(html_files)