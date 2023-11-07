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

def build_header_dict(element, parent_map, parent_path="", header_dict=None):
    if header_dict is None:
        header_dict = {}

    level = get_level(element, parent_map)
    tag = element.tag
    full_tag_path = f"{parent_path}/{tag}" if parent_path else tag

    header_dict[(tag, level)] = full_tag_path

    # Add attributes to the dictionary
    for attr in element.attrib:
        attr_path = f"{full_tag_path}@{attr}"
        header_dict[(attr, level)] = attr_path

    # Recursively add children to the dictionary
    for child in element:
        build_header_dict(child, parent_map, full_tag_path, header_dict)

    return header_dict



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

def get_full_tag_path(element, parent_path=""):
    tag = element.tag
    full_tag_path = f"{parent_path}/{tag}" if parent_path else tag
    return full_tag_path

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
        cell_path_row = [""] * len(headers)
        for i, (header, header_level) in enumerate(headers):  # Manual correction to export also attributes of root
            if header in attributes and header_level == 1:
                row[i] = (root.attrib[header])
                full_tag_path = header_dict[header,header_level]
                cell_path_row[i] = full_tag_path

        writer.writerow(row)
        cell_path_matrix.append(cell_path_row)

        write_to_single_line = False
        keywordLevel = -1
        currentKeyword = None

        row = [""] * len(headers)
        cell_path_row = [""] * len(headers)

        for item in root.findall('.//'):
            itemTag = item.tag
            itemLevel = get_level(item, parent_map)
            data = extract_xml_data(item)

            

            if (itemTag,itemLevel) in keywordsTuple:
                if write_to_single_line and (itemLevel <= keywordLevel or itemTag != currentKeyword):
                    writer.writerow(row)
                    cell_path_matrix.append(cell_path_row)
                    row = [""] * len(headers)
                    cell_path_row = [""] * len(headers)
                    
                write_to_single_line = True
                keywordLevel = itemLevel
                currentKeyword = itemTag

            elif itemLevel <= keywordLevel:
                if write_to_single_line:
                    writer.writerow(row)
                    cell_path_matrix.append(cell_path_row)
                    row = [""] * len(headers)
                    cell_path_row = [""] * len(headers)

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
                                full_tag_path = header_dict[header,header_level]
                                cell_path_row[i] = full_tag_path
                            elif isinstance(data[header], list):
                                flattened_dict = flatten_list(data[header]['#text'])
                                row[i] = (json.dumps(flattened_dict))
                                full_tag_path = header_dict[header,header_level]
                                cell_path_row[i] = full_tag_path
                            else:
                                row[i] = (data[header])
                                full_tag_path = header_dict[header,header_level]
                                cell_path_row[i] = full_tag_path
                    else:
                        try:
                            data['#text']
                            if isinstance(data['#text'], dict):
                                row[i] = (json.dumps(data['#text']))
                                full_tag_path = header_dict[header,header_level]
                                cell_path_row[i] = full_tag_path
                            elif isinstance(data['#text'], list):
                                flattened_dict = flatten_list(data['#text'])
                                row[i] = (json.dumps(flattened_dict))
                                full_tag_path = header_dict[header,header_level]
                                cell_path_row[i] = full_tag_path
                            else:
                                row[i] = (data['#text'])
                                full_tag_path = header_dict[header,header_level]
                                cell_path_row[i] = full_tag_path
                        except:
                            continue

            rowNotEmpty = any(element != "" for element in row)

            if rowNotEmpty and not write_to_single_line:
                writer.writerow(row)
                cell_path_matrix.append(cell_path_row)
                row = [""] * len(headers)
                cell_path_row = [""] * len(headers)


                    
                    
                       


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


def csv_to_xml(xml_structure, csv_file, xml_file):
    with open(csv_file, 'r') as f:
        reader = csv.reader(f, delimiter=",") #### kind of stupid, but if you save the csv file in excel, then delimiter is ";" and not "," as Python exports it
        headers = next(reader)

        # Create the XML structure using xml_structure
        root = create_xml_skeleton(xml_structure)

        # Create the header-path list
        header_path_list = create_header_path_list(xml_structure)

        # Go through each row in the csv file and populate the XML with data
        for csv_row_index, row in enumerate(reader):
            for csv_cell_index, string in enumerate(row):
                populate_xml_element(root, csv_row_index,csv_cell_index, row, cell_path_matrix)


        tree = ET.ElementTree(root)
        tree.write(xml_file)

def navigate_to_element(element, tag_path_parts):
    current_element = element
    for tag in tag_path_parts:
        if current_element.tag == tag:
            continue
        for child in current_element:
            if child.tag == tag:
                current_element = child
                break
    return current_element

def populate_xml_element(element, csv_row_index, csv_cell_index, row, cell_path_matrix):    ###### works quite nicely, but since there are multiple tags called "Attributes" all the data is written into the first one. Should think of a method to differentiate all "Attributes"
    full_tag_path = cell_path_matrix[csv_row_index][csv_cell_index]

    # Break full path into tag and attribute parts
    path_parts = full_tag_path.split("@")
    tag_path_parts = path_parts[0].split("/")
    attr = path_parts[1] if len(path_parts) > 1 else None

    # Navigate to the corresponding tag in XML
    current_element = navigate_to_element(element, tag_path_parts)

    # Set attribute value or text
    if attr:
        # Set attribute value
        current_element.set(attr, row[csv_cell_index])
    else:
        # Set text
        if(row[csv_cell_index] == '(Empty)'):
            current_element.text = ""
        else:
            current_element.text = row[csv_cell_index]










# Usage:
inputXML = "C:\\Users\\s.veress\\Desktop\\382 LP2+3.xml" #"Z:\\Mitarbeiter\\Szabolcs\\ArchicadVorlage\\AttributenStudie\\345_Linien.xml" Z:\\Mitarbeiter\\Szabolcs\\ArchicadVorlage\\Test.xml
outputCSV = 'C:\\Users\\s.veress\\Desktop\\output.csv'
outputXML = 'C:\\Users\\s.veress\\Desktop\\output.xml'

keywordsTuple = [("Layer",4), ("LineType",4), ("BuildingMaterial",4), ("CompositeWall",4), ("Material",4), ("Fill",4), ("Profile",4),("Attributes",3), ("PenTable", 4), ("Pens",5), ("Pen", 6)]#
# Load the XML file
tree = ET.parse(inputXML)
root = tree.getroot()


# Create a map from child elements to their parent
parent_map = {c: p for p in tree.iter() for c in p}

# Build the dictionary
header_dict = build_header_dict(root, parent_map)



# Now you can use header_dict to get the full path of any element,
# given its name and depth.



# Extract the XML structure
xml_structure = extract_xml_structure(root)
#xml_structure = merge_duplicate_keys(xml_structure)

#allXMLLeaves = get_all_values(xml_structure)

# Print the extracted structure

#print(json.dumps(xml_structure, indent=4))
print(xml_structure)

global cell_path_matrix
cell_path_matrix = []
# Parsing the XML file

xml_to_csv(inputXML, outputCSV)

#open_csv_in_excel(outputCSV)

input("Press Enter to continue...")

csv_to_xml(xml_structure, outputCSV, outputXML)






######################################################################

def create_header_path_list(xml_structure, path=""):
    header_path_list = []

    tag, attributes, children_structure, has_text = xml_structure

    full_tag_path = f"{path}/{tag}" if path else tag

    # add tag to the list
    header_path_list.append(full_tag_path)

    # add attributes to the list
    if attributes is not None:
        for attr in attributes:
            attr_path = f"{full_tag_path}@{attr}"
            header_path_list.append(attr_path)

    # recursively add children to the list
    if children_structure is not None:
        for child_structure in children_structure:
            header_path_list.extend(create_header_path_list(child_structure, full_tag_path))

    return header_path_list



header_path_list = []





