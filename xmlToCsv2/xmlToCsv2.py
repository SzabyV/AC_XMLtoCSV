
# Importing the required libraries
import xml.etree.ElementTree as ET
import pandas as pd

def extract_xml_structure(node, parent_path=''):
    """Recursively extract the structure of the XML file."""
    structure = {}

    if parent_path:
        parent_path += '.'

    if isinstance(node, ET.Element):
        if len(node) == 0:
            # Leaf node (element without children)
            structure[node.tag] = parent_path + node.tag
            if node.attrib:
                structure[node.tag] = {**structure[node.tag], '@': node.attrib}
        else:
            # Non-leaf node (element with children)
            children = {}

            for child in node:
                child_structure = extract_xml_structure(child, parent_path + node.tag)
                children.update(child_structure)

            structure[node.tag] = children
            if node.attrib:
                structure[node.tag] = {**structure[node.tag], '@': node.attrib}
    else:
        # Text node or other non-element node
        structure = parent_path + '<text_node>'

    return structure

# Usage:
inputXML = "Z:\\Mitarbeiter\\Szabolcs\\ArchicadVorlage\\AttributenStudie\\345_Linien.xml"
outputCSV = 'C:\\Users\\veres\\OneDrive\\Desktop\\output.csv'
outputXML = 'C:\\Users\\veres\\OneDrive\\Desktop\\output.xml'
# Load the XML file
tree = ET.parse(inputXML)
root = tree.getroot()

# Extract the XML structure
xml_structure = extract_xml_structure(root)


  
cols = ["name", "phone", "email", "date", "country"]
rows = []
  
# Parsing the XML file
xmlparse = ET.parse('sample.xml')
root = xmlparse.getroot()
for i in root:
    name = i.find("name").text
    phone = i.find("phone").text
    email = i.find("email").text
    date = i.find("date").text
    country = i.find("country").text
  
    rows.append({"name": name,
                 "phone": phone,
                 "email": email,
                 "date": date,
                 "country": country})
  
df = pd.DataFrame(rows, columns=cols)
  
# Writing dataframe to csv
df.to_csv('output.csv')
