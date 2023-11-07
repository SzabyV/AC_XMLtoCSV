import csv
import xml.etree.ElementTree as ET
import pandas as pd
import json
import os


def open_csv_in_excel(csv_file):
    # Read csv file using pandas and specifying the encoding
    df = pd.read_csv(csv_file, encoding='ISO-8859-1')  # or 'latin1'

    # Save as excel
    excel_file = csv_file.replace('.csv', '.xlsx')
    df.to_excel(excel_file, index=False)

    # Open the file in Excel
    os.system(f'start excel "{excel_file}"')

def extract_xml_structure(node):
    if isinstance(node, ET.Element):
        children = [extract_xml_structure(child) for child in node]
        has_text = True if node.text and not node.text.isspace() else False
        return (node.tag, node.attrib if node.attrib else None, children if children else None, has_text)
    else:
        return None



    return data

def extract_xml_data(node):
    """Extract the data from the XML elements on the upper level only."""
    data = {}
    
    if isinstance(node, ET.Element):
        if node.text: #and node.text.strip():
            if(node.text.strip()):
                data["#text"] = node.text.strip()
            else:
                data["#text"] = "(Empty)"

        if node.attrib:
            data.update(node.attrib)

        for child in node:
            if len(child) == 0:
                child_data = child.text.strip() if child.text else ''
                data[child.tag] = child_data

    return data

def get_level(element, parent_map):
    """Calculate the level of an XML element using parent mapping."""
    level = 0
    while element is not None:
        level += 1
        element = parent_map.get(element)
    return level

def flatten_list(data_list):
    """Flatten a list of dictionaries into a single dictionary."""
    flattened = {}
    for item in data_list:
        if isinstance(item, dict):
            flattened.update(item)
    return flattened

def is_header_in_data(header, data):
    if isinstance(data, dict):
        if header in data:
            return True
        for value in data.values():
            if is_header_in_data(header, value):
                return True
    return False

def xml_to_csv(xml_file, csv_file):
    tree = ET.parse(xml_file)
    root = tree.getroot()

    parent_map = {c: p for p in tree.iter() for c in p}

    with open(csv_file, 'w', newline='') as f:
        writer = csv.writer(f, delimiter= ",")

        headers = []
        header_set = set()

        for item in root.iter():
            item_tag = item.tag
            item_level = get_level(item, parent_map)  

            header = (item_tag, item_level)

            if header not in header_set:
                headers.append(header)
                header_set.add(header)

            for attr in item.attrib.keys():
                attribute = (attr, item_level)
                if attribute not in header_set:
                    headers.append(attribute)
                    header_set.add(attribute)

        headers.sort(key=lambda h: h[1])

        header_tags = [header[0] for header in headers]

        writer.writerow(header_tags)

        attributes = list(root.attrib.keys())

        row = [""] * len(headers)
        for i, (header, header_level) in enumerate(headers):  # Manual correction to export also attributes of root
            if header in attributes and header_level == 1:
                row[i] = (root.attrib[header])

        writer.writerow(row)

        write_to_single_line = False
        keywordLevel = -1
        currentKeyword = None

        row = [""] * len(headers)

        for item in root.findall('.//'):
            itemTag = item.tag
            itemLevel = get_level(item, parent_map)
            data = extract_xml_data(item)

            

            if (itemTag,itemLevel) in keywordsTuple:
                if write_to_single_line and (itemLevel <= keywordLevel or itemTag != currentKeyword):
                    writer.writerow(row)
                    row = [""] * len(headers)
                    
                write_to_single_line = True
                keywordLevel = itemLevel
                currentKeyword = itemTag

            elif itemLevel <= keywordLevel:
                if write_to_single_line:
                    writer.writerow(row)
                    row = [""] * len(headers)

                write_to_single_line = False
                keywordLevel = -1
                currentKeyword = None
                
            for i, (header, header_level) in enumerate(headers):
                if header == itemTag:
                    shortCategory = True
                else:
                    shortCategory = False

                if ((header in data) or shortCategory) and (itemLevel == header_level):
                    if(header != itemTag):
                        if(header == "RefID" and itemTag != "TypeRoot" and header_level == 3):
                            continue
                        else:
                            if isinstance(data[header], dict):
                                row[i] = (json.dumps(data[header]['#text']))  
                            elif isinstance(data[header], list):
                                flattened_dict = flatten_list(data[header]['#text'])
                                row[i] = (json.dumps(flattened_dict)) 
                            else:
                                row[i] = (data[header])
                    else:
                        try:
                            data['#text']
                            if isinstance(data['#text'], dict):
                                row[i] = (json.dumps(data['#text']))  
                            elif isinstance(data['#text'], list):
                                flattened_dict = flatten_list(data['#text'])
                                row[i] = (json.dumps(flattened_dict)) 
                            else:
                                row[i] = (data['#text'])
                        except:
                            continue

            rowNotEmpty = any(element != "" for element in row)

            if rowNotEmpty and not write_to_single_line:
                writer.writerow(row)
                row = [""] * len(headers)


                    
                    
                       


def create_xml_skeleton(structure):
    tag, attributes, children_structure, has_text = structure

    element = ET.Element(tag)

    if attributes:
        for attr, value in attributes.items():
            element.set(attr, value)

    if children_structure:
        for child_structure in children_structure:
            child_element = create_xml_skeleton(child_structure)
            element.append(child_element)

    return element


def populate_xml_element(element, row, headers):
    tag = element.tag
    text = row[headers.index(tag)] if tag in headers else None

    # Replace '(Empty)' with an empty string
    if text == '(Empty)':
        text = ''

    if text and element.text is None:  # Only set text if the tag matches and no text is set yet
        element.text = text

    for child in element:
        populate_xml_element(child, row, headers)


def csv_to_xml(xml_structure, csv_file, xml_file):
    root_element = create_xml_skeleton(xml_structure)

    with open(csv_file, 'r') as f:
        reader = csv.reader(f)
        headers = next(reader)

        for row in reader:
            populate_xml_element(root_element, row, headers)

        tree = ET.ElementTree(root_element)
        tree.write(xml_file)






# Usage:
inputXML = "Z:\\Mitarbeiter\\Szabolcs\\ArchicadVorlage\\Test.xml" #"Z:\\Mitarbeiter\\Szabolcs\\ArchicadVorlage\\AttributenStudie\\345_Linien.xml"
outputCSV = 'C:\\Users\\s.veress\\Desktop\\output.csv'
outputXML = 'C:\\Users\\s.veress\\Desktop\\output.xml'

keywordsTuple = [("Layer",4), ("LineType",4), ("BuildingMaterial",4), ("CompositeWall",4), ("Material",4), ("Fill",4), ("Profile",4),("Attributes",3)]#
# Load the XML file
tree = ET.parse(inputXML)
root = tree.getroot()






# Extract the XML structure
xml_structure = extract_xml_structure(root)
#xml_structure = merge_duplicate_keys(xml_structure)

#allXMLLeaves = get_all_values(xml_structure)

# Print the extracted structure

#print(json.dumps(xml_structure, indent=4))
print(xml_structure)


# Parsing the XML file

xml_to_csv(inputXML, outputCSV)

#open_csv_in_excel(outputCSV)

input("Press Enter to continue...")

csv_to_xml(xml_structure, outputCSV, outputXML)
