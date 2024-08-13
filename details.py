import requests
from bs4 import BeautifulSoup
import re
import json

def extract_company_website(id):
    url = 'https://232app.azurewebsites.net/Forms/ExclusionRequestItem/' + id
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/127.0.0.0 Safari/537.36 Edg/127.0.0.0'
    }

    try:
        response = requests.get(url, headers=headers)
        if response.status_code == 200:
            soup = BeautifulSoup(response.content, 'html.parser')
            
            company_website = soup.find('input', {'id': 'BIS232Request_JSONData_RequestingOrg_WebsiteAddress'})
            company_website = company_website['value'] if company_website and 'value' in company_website.attrs else None
            
            metal_class = soup.find('input', {'id': 'BIS232Request_JSONData_MetalClass'})
            metal_class = metal_class['value'] if metal_class and 'value' in metal_class.attrs else None
            
            previously_granted_er = soup.find('input', {'id': 'BIS232Request_JSONData_PreviouslyGrantedER'})
            previously_granted_er = previously_granted_er['value'] if previously_granted_er and 'value' in previously_granted_er.attrs else None
            
            total_requested_annual = soup.find('input', {'id': 'BIS232Request_JSONData_TotalRequestedAnnualExclusionQuantity'})
            total_requested_annual = total_requested_annual['value'] if total_requested_annual and 'value' in total_requested_annual.attrs else None
            
            avg_annual_consumption = soup.find('input', {'id': 'BIS232Request_JSONData_ExclusionExplanation_AvgAnnualConsumption'})
            avg_annual_consumption = avg_annual_consumption['value'] if avg_annual_consumption and 'value' in avg_annual_consumption.attrs else None
            
            product_description = soup.find('textarea', {'id': 'BIS232Request_JSONData_ProductDescription_Description'})
            product_description = product_description.text.strip() if product_description else None
            
            chemical_composition_comments = soup.find('textarea', {'id': 'BIS232Request_JSONData_ChemicalCompositionComments'})
            chemical_composition_comments = chemical_composition_comments.text.strip() if chemical_composition_comments else None
            
            product_applications = soup.find('textarea', {'id': 'BIS232Request_JSONData_AdditionalDetails_ApplicationSuitability'})
            product_applications = product_applications.text.strip() if product_applications else None
            
            commercial_names = None
            source_countries = None

            script_tags = soup.find_all('script')
            for script in script_tags:
                script_content = script.string
                if script_content:
                    commercial_names_match = re.search(r'function createCommercialNamesTable\(\) \{.*?var arrValues = \[(.*?)\];', script_content, re.DOTALL)
                    if commercial_names_match:
                        commercial_names = commercial_names_match.group(1).replace('"', '').split(', ')
                    
                    source_countries_match = re.search(r'function createSourceCountriesTable\(\) \{.*?var arrValues = \[(.*?)\];', script_content, re.DOTALL)
                    if source_countries_match:
                        arr_values_str = source_countries_match.group(1)
                        source_countries = json.loads(f"[{arr_values_str}]")
                        
            return {
                'company_website': company_website,
                'metal_class': metal_class,
                'previously_granted_er': previously_granted_er,
                'total_requested_annual': total_requested_annual,
                'avg_annual_consumption': avg_annual_consumption,
                'product_description': product_description,
                'chemical_composition_comments': chemical_composition_comments,
                'product_applications': product_applications,
                'commercial_names': commercial_names,
                'source_countries': source_countries
            }
        else:
            print(f"Failed to retrieve the webpage. Status code: {response.status_code}")
            return None
    except requests.exceptions.RequestException as e:
        print(f"An error occurred: {e}")
        return None
