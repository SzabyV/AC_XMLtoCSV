
import csv
import xml.etree.ElementTree as ET
import pandas as pd
import json




def extract_xml_structure(node, parent_path='', depth=0, occurrences=None):
    nodeTag = node.tag
    structure = {}

    if parent_path:
        parent_path += '.'

    if isinstance(node, ET.Element):
        if len(node) == 0:
            key = (node.tag, depth)
            structure['#text', depth] = parent_path + node.tag
            if node.attrib:
                structure['@', depth] = node.attrib
        else:
            children = {}
            if occurrences is None:
                occurrences = {}

            for child in node:
                childTag = child.tag
                child_occurrences = occurrences.get(childTag, 0) + 1
                occurrences[childTag] = child_occurrences

                child_structure = extract_xml_structure(child, parent_path + node.tag, depth + 1, occurrences)
                child_key = (child.tag, depth + 1, child_occurrences)
                if child_key in children:
                    if isinstance(children[child_key], list):
                        children[child_key].append(child_structure)
                    else:
                        children[child_key] = [children[child_key], child_structure]
                else:
                    children[child_key] = child_structure

            if node.attrib:
                children['@', depth] = node.attrib
            parent_key = (node.tag, depth)
            if parent_key in children.keys():
                structure = {**structure, **children}
            else:
                structure[parent_key] = children
    else:
        structure['#text', depth] = parent_path + '<text_node>'

    return structure



    return data

def extract_xml_data(node):
    """Extract the data from the XML elements on the upper level only."""
    data = {}
    
    if isinstance(node, ET.Element):
        if node.text and node.text.strip():
            data["#text"] = node.text.strip()

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
    rootTag = root.tag

    parent_map = {c: p for p in tree.iter() for c in p}

    with open(csv_file, 'w', newline='') as f:
        writer = csv.writer(f)

        headers = []
        header_set = set()

        for item in root.iter():

            
            item_tag = item.tag
            item_level = get_level(item, parent_map)  # Replace `get_level` with your own function to determine the level of the XML element

            header = (item_tag, item_level)

            

            if header not in header_set:
                headers.append(header)
                header_set.add(header)

            
            for attr in item.attrib.keys():
                attribute = (attr, item_level)
                if attribute not in header_set:
                    headers.append(attribute)
                    header_set.add(attribute)


        # Sort the headers based on the level, so headers at higher levels come first
        headers.sort(key=lambda h: h[1])

        # Extract the tag names from the sorted headers
        header_tags = [header[0] for header in headers]

        writer.writerow(header_tags)

        attributes = list(root.attrib.keys())

        
        #data = extract_xml_data(attr)
        row = []
        for header, header_level in headers:
            if header in attributes and header_level == 1:
                row.append(root.attrib[header])
            else:
                row.append('')
        writer.writerow(row)

        write_to_single_line = False  # Flag to determine whether to write to a single line
        #single_line_data = []  # Store the data for the single line
        keywordLevel = 0

        for item in root.findall('.//'):
            
            itemTag = item.tag
            itemLevel = get_level(item, parent_map)
            data = extract_xml_data(item)

            if write_to_single_line and itemLevel <= keywordLevel: # if new item is a new keyword then exportt the new line, before saving the next one
                writer.writerow(row)  # Write the single line data as a complete line
                write_to_single_line = False  # Reset the flag
                keywordLevel = 0

            if not write_to_single_line: # reset row if a new line is needed
                row=[]
                for header, header_level in headers: #fill row with empty cells
                    if(not write_to_single_line):
                        row.append("")

            if itemTag in keywordsTuple and not write_to_single_line:
                write_to_single_line = True  # Start writing to a single line
                #single_line_data = []  # Clear the single line data
                keywordLevel = itemLevel

            

            #row = []
            for i, (header, header_level) in enumerate(headers):
                
                #if is_header_in_data (header, data):
                if(header == itemTag):
                    shortCategory = True
                else:
                    shortCategory = False

                


                if ((header in data) or shortCategory) and (itemLevel == header_level): #
                #if ((header in data) or (header in itemTag)):
                    #if(itemLevel == header_level):
                        if(header != itemTag):
                            if(header == "RefID" and itemTag != "TypeRoot" and header_level == 3):
                                continue
                            else:
                                if isinstance(data[header], dict):
                                    row[i] = (json.dumps(data[header]['#text']))  # Convert dictionary to JSON string
                                elif isinstance(data[header], list):
                                    flattened_dict = flatten_list(data[header]['#text'])
                                    row[i] = (json.dumps(flattened_dict))  # Convert flattened dictionary to JSON string
                                else:
                                    row[i] = (data[header])
                        else:
                            try:
                                data['#text']
                                if isinstance(data['#text'], dict):
                                    row[i] = (json.dumps(data['#text']))  # Convert dictionary to JSON string
                                elif isinstance(data['#text'], list):
                                    flattened_dict = flatten_list(data['#text'])
                                    row[i] = (json.dumps(flattened_dict))  # Convert flattened dictionary to JSON string
                                else:
                                    row[i] = (data['#text'])
                            except:
                                #row[i] = ('')
                                continue
                else:
                    #row[i] = ('')
                    continue



            rowNotEmpty = False
            for i, element in enumerate(row):
                if element != "": #and i != 11: # added exception for RefId cell that is a doubled property and are on two consequent rows
                    rowNotEmpty = True
                    break
            if rowNotEmpty:
                if not write_to_single_line:
                        writer.writerow(row)  
                #else:
                    
                    
                        

def csv_to_xml(xml_structure, csv_file, xml_file):
    root_element = list(xml_structure.keys())[0]

    with open(csv_file, 'r') as f:
        reader = csv.reader(f)
        headers = next(reader)

        root = ET.Element(root_element)

        for row in reader:
            element = create_xml_element(xml_structure, headers, row)
            root.append(element)

        tree = ET.ElementTree(root)
        tree.write(xml_file)



def create_xml_element(xml_structure, headers, row):
    element = ET.Element(list(xml_structure.keys())[0])

    for index, value in enumerate(row):
        header = headers[index]

        if header in xml_structure:
            if header == '@':
                # Handle attributes
                for attr, attr_value in xml_structure[header].items():
                    element.set(attr, attr_value)
            elif isinstance(xml_structure[header], dict):
                child_element = create_xml_element(xml_structure[header], headers, row)
                element.append(child_element)
            elif isinstance(xml_structure[header], list):
                for item in xml_structure[header]:
                    child_element = create_xml_element(item, headers, row)
                    element.append(child_element)
            else:
                child_element = ET.Element(header)
                child_element.text = value
                element.append(child_element)

    return element




# Usage:
inputXML = "Z:\\Mitarbeiter\\Szabolcs\\ArchicadVorlage\\Test.xml" #"Z:\\Mitarbeiter\\Szabolcs\\ArchicadVorlage\\AttributenStudie\\345_Linien.xml"
outputCSV = 'C:\\Users\\s.veress\\Desktop\\output.csv'
outputXML = 'C:\\Users\\s.veress\\Desktop\\output.xml'

keywordsTuple = ["Layer", "LineType", "BuildingMaterial", "CompositeWall", "Material", "Fill", "Profile"]#
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

csv_to_xml(xml_structure, outputCSV, outputXML)








###############################################################################################################################################################################










def xml_to_csv2(xml_file, csv_file):
    tree = ET.parse(xml_file)
    root = tree.getroot()

    cols = []
    rows = []

    for item in root.iter():
        cols.append(item.tag)
        
    cols = list(set(cols))
    for i in root:
        
        contents = []
        for name in cols:
            
            content = i.find(name)
            if not isinstance(content, str):
                contents.append("") 
            else:
                contents.append(content)
        rows.append(contents)

    df = pd.DataFrame(rows, columns=cols)
    df.to_csv(csv_file, index=False)
    
def xml_to_csv3(xml_file, csv_file):

    rows = []
    cols = allXMLLeaves
    cols = list(set(cols))

    for i in root:
        for name in cols:
            content = i.find(name)
        
  
            rows.append({name: content,
                         })
  
    df = pd.DataFrame(rows, columns=cols)

    # Writing dataframe to csv
    df.to_csv(csv_file)

def get_all_values(d):
            values = []
            if(isinstance(d, dict)):
                for value in d.values():
                    if isinstance(value, dict):
                        values.extend(get_all_values(value))
                    else:
                        values.append(value)
                return values
            else:
                values.append(d)
                return values

def extract_xml_data99(node):
    """Recursively extract the data from XML based on the provided structure."""
    data = {}

    if node.attrib:
        data['@attributes'] = node.attrib

    if len(node) == 0:
        # Leaf node (element without children)
        data[node.tag] = node.text
    else:
        # Non-leaf node (element with children)
        for child in node:
            child_data = extract_xml_data(child)

            if child.tag in data:
                if isinstance(data[child.tag], list):
                    data[child.tag].append(child_data)
                else:
                    data[child.tag] = [data[child.tag], child_data]
            else:
                data[child.tag] = child_data

    return data


def xml_to_csv99(xml_file, csv_file):
    tree = ET.parse(xml_file)
    root = tree.getroot()

    with open(csv_file, 'w', newline='') as f:
        writer = csv.writer(f)

        headers = set()
        for item in root.iter():
            headers.add(item.tag)
            for attr in item.attrib.keys():
                headers.add(attr)

        writer.writerow(headers)
        
        for item in root.findall('.//'):
            data = extract_xml_data(item)
            row = []
            for header in headers:
                if header in data:
                    if isinstance(data[header], dict):
                        row.append(json.dumps(data[header]))  # Convert dictionary to JSON string
                    elif isinstance(data[header], list):
                        row.append(', '.join(data[header]))
                    else:
                        row.append(data[header])
                else:
                    row.append('')
            writer.writerow(row)


def create_xml_element99(xml_structure, headers, row):
    element = ET.Element(list(xml_structure.keys())[0])
    elementTag = element.tag

    for index, value in enumerate(row):
        header = headers[index]

        if header in xml_structure:
            if isinstance(xml_structure[header], dict):
                child_element = create_xml_element(xml_structure[header], headers, row)
                element.append(child_element)
            elif header != '#text':
                element.set(header, value)
            else:
                element.text = value

    return element

def create_xml_element99(xml_structure, headers, row):
    element = ET.Element(list(xml_structure.keys())[0])

    for index, value in enumerate(row):
        header = headers[index]

        if header in xml_structure:
            if isinstance(xml_structure[header], dict):
                child_element = create_xml_element(xml_structure[header], headers, row)
                element.append(child_element)
            else:
                attribute = header
                if isinstance(xml_structure, dict) and '@' in xml_structure:
                    attribute_value = xml_structure.get('@')
                    if attribute_value is not None:
                        element.set(attribute, attribute_value)
                else:
                    element.text = value

    return element

def merge_duplicate_keys(dictionary):
    merged_dict = {}
    for key, value in dictionary.items():
        if isinstance(value, dict):
            # Recursively merge keys of nested dictionaries
            value = merge_duplicate_keys(value)
        if isinstance(key, tuple):
            sub_dict = merged_dict
            for sub_key in key[:-1]:
                if sub_key not in sub_dict:
                    sub_dict[sub_key] = {}
                sub_dict = sub_dict[sub_key]
            sub_dict[key[-1]] = value
        else:
            merged_dict[key] = value
    return merged_dict

def create_xml_element97(xml_structure, headers, row):
    element = ET.Element(list(xml_structure.keys())[0])

    for header, value in zip(headers, row):
        if header in xml_structure:
            if isinstance(xml_structure[header], dict):
                child_element = create_xml_element(xml_structure[header], headers, row)
                element.append(child_element)
            else:
                child_element = ET.Element(header)
                child_element.text = value
                element.append(child_element)

    return element


def extract_xml_structure97(node, parent_path='', depth=0):
    nodeTag = node.tag
    """Recursively extract the structure of the XML file."""
    structure = {}

    if parent_path:
        parent_path += '.'

    if isinstance(node, ET.Element):
        if len(node) == 0:
            structure['#text', depth] = parent_path + node.tag
            if node.attrib:
                structure['@', depth] = node.attrib
        else:
            children = {}

            for child in node:
                childTag = child.tag
                child_structure = extract_xml_structure(child, parent_path + node.tag, depth + 1)
                child_key = (child.tag, depth + 1)
                if child_key in children:
                    if isinstance(children[child_key], list):
                        children[child_key].append(child_structure)
                    else:
                        children[child_key] = [children[child_key], child_structure]
                else:
                    children[child_key] = child_structure

            if node.attrib:
                children['@', depth] = node.attrib
            parent_key = (node.tag, depth)
            if(parent_key in children.keys()):
                structure = {**structure, **children}
            else:
                structure[parent_key] = children
    else:
        structure['#text', depth] = parent_path + '<text_node>'

    return structure



def extract_xml_structure98(node, parent_path=''):
    nodeTag = node.tag
    """Recursively extract the structure of the XML file."""
    structure = {}

    if parent_path:
        parent_path += '.'

    if isinstance(node, ET.Element):
        if len(node) == 0:
            # Leaf node (element without children)
            structure['#text'] = parent_path + node.tag
            if node.attrib:
                structure['@'] = node.attrib
        else:
            # Non-leaf node (element with children)
            children = {}

            for child in node:
                childTag = child.tag
                child_structure = extract_xml_structure(child, parent_path + node.tag)
                if child.tag in children:
                    if isinstance(children[child.tag], list):
                        children[child.tag].append(child_structure)
                    else:
                        children[child.tag] = [children[child.tag], child_structure]
                else:
                    children[child.tag] = child_structure

            if node.attrib:
                children['@'] = node.attrib
            structure[node.tag] = children
    else:
        # Text node or other non-element node
        structure['#text'] = parent_path + '<text_node>'

    return structure

def extract_xml_structure99(node, parent_path=''):
    nodeTag = node.tag
    """Recursively extract the structure of the XML file."""
    structure = {}

    if parent_path:
        parent_path += '.'

    if isinstance(node, ET.Element):
        if len(node) == 0:
            # Leaf node (element without children)
            structure[node.tag] = parent_path + node.tag
            if node.attrib:
                if isinstance(structure[node.tag], str):
                    structure[node.tag] = {'#text': structure[node.tag], '@': node.attrib} # if node has "Type" attribute
                else:
                    structure[node.tag] = {**structure[node.tag], '@': node.attrib}
            else:
                structure[node.tag] = {'#text': structure[node.tag]}
        else:
            # Non-leaf node (element with children)
            children = {}

            for child in node:
                childTag = child.tag
                child_structure = extract_xml_structure(child, parent_path + node.tag)
                if child.tag in children:
                    if isinstance(children[child.tag], list):
                        children[child.tag].append(child_structure)
                    else:
                        children[child.tag] = [children[child.tag], child_structure]
                else:
                    children[child.tag] = child_structure

            structure[node.tag] = children
            if node.attrib:
                structure[node.tag] = {**structure[node.tag], '@': node.attrib}
    else:
        # Text node or other non-element node
        structure = {'#text': parent_path + '<text_node>'}

    return structure




def extract_xml_data_recursive(node):
    """Recursively extract the data from the XML elements."""
    data = {}
    nodeTag = node.tag

    if isinstance(node, ET.Element):
        if node.text and node.text.strip():
            nodeText = node.text
            data[nodeTag] = node.text.strip()
            #data['#text'] = node.text.strip() ## old version from chatgpt

        if node.attrib:
            nodeAttrib = node.attrib
            data.update(node.attrib)

        for child in node:
            child_data = extract_xml_data(child)
            if child.tag in data:
                if isinstance(data[child.tag], list):
                    data[child.tag].append(child_data)
                else:
                    data[child.tag] = [data[child.tag], child_data]
            else:
                data[child.tag] = child_data