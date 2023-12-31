import requests
from bs4 import BeautifulSoup
import pandas as pd

# The URL of the organization's project page - e.g "https://www.archaeologydataservice.ac.uk/library/browse/organisationDetails.xhtml?organisationId=1004881"
organizations = []

# Create lists to store data
data = {"Organization": [], "Project": []}

for url in organizations:
    response = requests.get(url)
    soup = BeautifulSoup(response.content, "html.parser")

    # Find the organization name within the td tag, handle NoneType
    organization_td = soup.find("td", string=True)
    organization_name = organization_td.get_text(strip=True) if organization_td else "Unknown"

    # Find all <a> elements for project titles
    project_links = soup.find_all("a", title=True)

    # Append organization name and project titles to the data lists
    for link in project_links:
        project_name = link.get_text()
        data["Organization"].append(organization_name)
        data["Project"].append(project_name)

# Create a pandas DataFrame from the data
df = pd.DataFrame(data)

# Export the DataFrame to an Excel file
excel_file = "projects.xlsx"
df.to_excel(excel_file, index=False)
print(f"Data exported to {excel_file}")

