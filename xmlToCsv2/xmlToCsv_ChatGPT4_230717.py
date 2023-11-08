from ast import Try
import csv
import xml.etree.ElementTree as ET
import pandas as pd
import json
import os
from collections import defaultdict

#from symbol import try_stmt


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

def extract_xml_data(node, parent_map):
    """Extract the data from the XML elements on the upper level only."""
    data = {}
    
    if isinstance(node, ET.Element):
        parent = parent_map.get(node)
        parentTag = parent.tag if parent is not None else None

        if node.text: #and node.text.strip():
            if(node.text.strip()):
                data["#text"] = node.text.strip()
            else:
                data["#text"] = "(Empty)"

        if node.attrib:
            for key, value in node.attrib.items():
                if value.strip():
                    data[(key, node.tag)] = value
                else:
                    data[(key, node.tag)] = "(Empty)"

        for child in node:
            if len(child) == 0:
                child_data = child.text.strip() if child.text else ''
                child_tag = child.tag #if "Idx" not in child.attrib else f"{child.tag}_{child.attrib['Idx']}"
                data[(child_tag, node.tag)] = child_data

    return data


def get_level(element, parent_map):
    """Calculate the level of an XML element using parent mapping."""
    level = 0
    while element is not None:
        level += 1
        element = parent_map.get(element)
    return level



def build_header_dict(element, parent_map, parent_path="", header_dict=None, keyword_indexes=None, last_keyword_level=-1, currentKeyword = None , duplicate_indexes= None):
    if header_dict is None:
        header_dict = {}

    if keyword_indexes is None:
        keyword_indexes = {keyword: 0 for keyword in keywords}

    if duplicate_indexes is None:
        duplicate_indexes = {}

    level = get_level(element, parent_map)

    parent_tag = parent_path.split("/")[-1].split("_idx")[0] if parent_path else None
    try:
        grandParent_tag = parent_path.split("/")[-2].split("_idx")[0] if parent_path else None
    except:
        grandParent_tag = None

    try:
        gGrandParent_tag = parent_path.split("/")[-3].split("_idx")[0] if parent_path else None
    except:
        gGrandParent_tag = None

    elementTag = element.tag
    
    ancestors,ancestor_tags,ancestor_tags_tuple,ancestor_tags_tuple_withoutParent, ancestor_tags_tuple_withoutGrandParent, parent_tag,grandParent_tag = get_ancestors(element,parent_map)
    """
    ancestor_tags = [tag for (tag, level) in ancestors]

    if(len(ancestor_tags)>2):
        ancestor_tags_tuple_withoutParent = tuple(ancestor_tags[1:])
    else:
        ancestor_tags_tuple_withoutParent = None

    if(len(ancestor_tags)>1):
            ancestor_tags_tuple = tuple(ancestor_tags)
    else:
        ancestor_tags_tuple = None
    """
    

    # For duplicate elements, we keep track of the count for each tag at each level
    # and each parent. This count serves as the index for the element.
    if (elementTag, level, ancestor_tags_tuple) in duplicate_indexes:  #if (elementTag, level, parent_tag, grandParent_tag) in duplicate_indexes:
        duplicate_indexes[(elementTag, level, ancestor_tags_tuple)] += 1 # improvement idea1: duplicate_indexes[elementTag, level, [*list with all ancestor's tags*]] /// duplicate_indexes[(elementTag, level, parent_tag,grandParent_tag,gGrandParent_tag)] += 1
    else:
        duplicate_indexes[(elementTag, level, ancestor_tags_tuple)] = 0 #improvement idea1: duplicate_indexes[elementTag, level, [*list with all ancestor's tags*]] duplicate_indexes[(elementTag, level, parent_tag,grandParent_tag,gGrandParent_tag)]
    index = duplicate_indexes[(elementTag, level, ancestor_tags_tuple)] # index = duplicate_indexes[(elementTag, level, parent_tag,grandParent_tag,gGrandParent_tag)]

    if (elementTag, level) in keywordsTuple:
        currentKeyword = elementTag
        keyword_indexes[elementTag] += 1
        tag = f"{elementTag}_{keyword_indexes[elementTag]}"
        last_keyword_level = level
    else:
        if level <= last_keyword_level: #reset counter
            tag = elementTag
            currentKeyword = None
            last_keyword_level = -1
        else:
            if currentKeyword:
                tag = f"{elementTag}_{keyword_indexes[currentKeyword]}"
            else:
                tag = elementTag

    full_tag_path = f"{parent_path}/{tag}_idx{index}" if parent_path else f"{tag}_idx{index}"

    # Add index to the key in the dictionary.
    header_dict[(tag, level, ancestors[0][0], index,ancestor_tags_tuple_withoutParent)] = full_tag_path # improvement idea1: header_dict[(tag, level, parent_tag, index, [*list with all ancestor's tags*])] -> use get_ancestors function /// header_dict[(tag, level, parent_tag, index,grandParent_tag,gGrandParent_tag)] = full_tag_path

    attr_duplicate_indexes = {}

    for attr in element.attrib:
        attr_path = f"{full_tag_path}@{attr}"
        attr_duplicate_indexes[attr] = attr_duplicate_indexes.get(attr, 0) + 1
        attr_index = attr_duplicate_indexes[attr] - 1
        if currentKeyword:
            attrName = f"{attr}_{keyword_indexes[currentKeyword]}"
            header_dict[(attrName, level, tag, attr_index,ancestor_tags_tuple)] = attr_path # improvement idea1: ///header_dict[(attrName, level, tag, attr_index,parent_tag,grandParent_tag)] = attr_path
        else:
            header_dict[(attr, level, tag, attr_index,ancestor_tags_tuple)] = attr_path # improvement idea1: //// header_dict[(attr, level, tag, attr_index,parent_tag,grandParent_tag)] = attr_path

    for child in element:
        build_header_dict(child, parent_map, full_tag_path, header_dict, keyword_indexes, last_keyword_level, currentKeyword, duplicate_indexes) # improvement idea1:

    return header_dict


def get_full_tag_path(element, parent_path=""):
    tag = element.tag
    full_tag_path = f"{parent_path}/{tag}" if parent_path else tag
    return full_tag_path

def process_children_dict(children_dict):
    # Create a list of keywords
    keywords = list(children_dict.keys())

    # Compare each list of elements with every other list of elements
    for i in range(len(keywords)):
        for j in range(len(keywords)):
            if i == j:
                continue  # Don't compare a list with itself

            # If the elements of one list are a subset of the other, remove the elements from the larger set
            if set(children_dict[keywords[i]]) <= set(children_dict[keywords[j]]):
                children_dict[keywords[j]] = list(set(children_dict[keywords[j]]) - set(children_dict[keywords[i]]))

    return children_dict

def get_headers(root, parent_map): ####still some problems with saving all the headers - some branches have significantly more sub-branches, which have an impact on the headers amount too. The error might be here or with csv exporting. More troubleshooting necessary...
    children_dict = {}
    occurrence_counts = {}

    for item in root.iter():
        itemTag = item.tag
        item_level = get_level(item, parent_map)

        ancestors,ancestor_tags,ancestor_tags_tuple,ancestor_tags_tuple_withoutParent,ancestor_tags_tuple_withoutGrandParent,parentTag,grandParentTag  = get_ancestors(item,parent_map)
        """
        ancestor_tags = [tag for (tag, level) in ancestors]
        x= len(ancestor_tags)
        if(len(ancestor_tags)>2):
            ancestor_tags_tuple_withoutParent = tuple(ancestor_tags[1:])
        else:
            ancestor_tags_tuple_withoutParent = None

        if(len(ancestor_tags)>1):
                ancestor_tags_tuple = tuple(ancestor_tags)
        else:
            ancestor_tags_tuple = None
        

        parent = parent_map.get(item)
        parentTag = parent.tag if parent is not None else None
        grandParent = parent_map.get(parent)
        grandParentTag = grandParent.tag if grandParent is not None else None
        gGrandParent = parent_map.get(grandParent)
        gGrandParentTag = gGrandParent.tag if gGrandParent is not None else None

        ggGrandParent = parent_map.get(gGrandParent)
        ggGrandParentTag = ggGrandParent.tag if ggGrandParent is not None else None

        gggGrandParent = parent_map.get(ggGrandParent)
        gggGrandParentTag = gggGrandParent.tag if gggGrandParent is not None else None
        """
        
        for tag,level,ancestor_tags_tuple_temp in occurrence_counts.keys(): #for tag,level,pTag,GpTag,gGpTag,ggGpTag,gggGpTag in occurrence_counts.keys():
            if level > item_level:
                occurrence_counts[tag,level,ancestor_tags_tuple_temp] = 0 #occurrence_counts[tag,level,pTag,GpTag,gGpTag,ggGpTag,gggGpTag] = 0
               

        # Increment the occurrence count for this tag, level, parent, and grandparent, and get the current count.
        occurrence_counts[(itemTag, item_level, ancestor_tags_tuple)] = occurrence_counts.get((itemTag, item_level, ancestor_tags_tuple), 0) + 1 #occurrence_counts[(itemTag, item_level, parentTag,grandParentTag, gGrandParentTag,ggGrandParentTag,gggGrandParentTag)] = occurrence_counts.get((itemTag, item_level, parentTag, grandParentTag,gGrandParentTag,ggGrandParentTag,gggGrandParentTag), 0) + 1
        index = occurrence_counts[(itemTag, item_level, ancestor_tags_tuple)] - 1 #index = occurrence_counts[(itemTag, item_level, parentTag, grandParentTag, gGrandParentTag,ggGrandParentTag,gggGrandParentTag)] - 1

        # Add the item to its parent's list of children in children_dict.
        children_dict.setdefault((parentTag, grandParentTag, ancestor_tags_tuple_withoutGrandParent),[]).append((itemTag, item_level, parentTag, index, grandParentTag, ancestor_tags_tuple_withoutGrandParent)) #children_dict.setdefault((parentTag,grandParentTag,gGrandParentTag,ggGrandParentTag,gggGrandParentTag),[]).append((itemTag, item_level, parentTag, index, grandParentTag,gGrandParentTag,ggGrandParentTag,gggGrandParentTag))

        # Reset the index for each attribute as they are local to the item.
        attr_occurrence_counts = {}

        for attr in item.attrib.keys():
            attr_occurrence_counts[attr] = attr_occurrence_counts.get(attr, 0) + 1
            attr_index = attr_occurrence_counts[attr] - 1
            children_dict.setdefault((itemTag, parentTag, ancestor_tags_tuple_withoutParent), []).append((attr, item_level, itemTag, attr_index, parentTag,ancestor_tags_tuple_withoutParent)) # children_dict.setdefault((itemTag,parentTag,grandParentTag,gGrandParentTag,ggGrandParentTag), []).append((attr, item_level, itemTag, attr_index, parentTag,grandParentTag,gGrandParentTag,ggGrandParentTag))

    #children_dict = process_children_dict(children_dict)

    headers = []
    processed_set = set()
    
    def add_header(header):
        if header in processed_set:
            return
        
        processed_set.add(header)
        headers.append(header)
        z = header[0]
        y = header[4]
        w = header[5]
        x = children_dict.get((header[0], header[4], header[5]), [])
        for child in children_dict.get((header[0], header[4], header[5]), []): #for child in children_dict.get((header[0],header[2], header[4],header[5],header[6]), []):
            #if((child[0], child[1]) not in keywordsTuple):
                add_header(child)
    
    for header in children_dict.get((None,None, None), []): #for header in children_dict.get((None,None,None,None,None), []):
        add_header(header)

    #for header in children_dict.get(None, []):
        #add_header(header)

    finalHeaders = []
    seen_keywords = set()  # set to keep track of the keywords you have seen

    for header in headers:
        if (header[0], header[1]) in keywordsTuple:
            if (header[0], header[1]) not in seen_keywords:
                # If this keyword pair hasn't been seen before, add it to the headers and the seen set
                finalHeaders.append(header)
                seen_keywords.add((header[0], header[1]))
        else:
            # If it's not a keyword, just add it to the headers
            finalHeaders.append(header)
    
    headers = finalHeaders
    header_tags = [header[0] for header in headers]

    return headers, header_tags


def get_ancestors(element, parent_map):
    ancestors = []

    def find_ancestors(item):
        parent = parent_map.get(item)
        if parent is not None:
            ancestors.append((parent.tag, get_level(parent, parent_map)))
            find_ancestors(parent)
        else:
            ancestors.append((None, 0))

    find_ancestors(element)

    ancestor_tags = [tag for (tag, level) in ancestors]
    
    if(len(ancestor_tags)>3):
        ancestor_tags_tuple_withoutGrandParent = tuple(ancestor_tags[2:])
        
    else:
        ancestor_tags_tuple_withoutGrandParent = None
        

    if(len(ancestor_tags)>2):
        ancestor_tags_tuple_withoutParent = tuple(ancestor_tags[1:])
        
    else:
        ancestor_tags_tuple_withoutParent = None
        

    if(len(ancestor_tags)>1):
            ancestor_tags_tuple = tuple(ancestor_tags)
    else:
        ancestor_tags_tuple = None

    try: 
        parentTag = ancestor_tags[0]
    except:
        parentTag = None
    try:
        grandParentTag = ancestor_tags[1]
    except:
        grandParentTag = None

    return ancestors, ancestor_tags, ancestor_tags_tuple, ancestor_tags_tuple_withoutParent,ancestor_tags_tuple_withoutGrandParent, parentTag, grandParentTag




def get_first_row(headers, row, root,cell_path_matrix):

    attributes = list(root.attrib.keys())
    full_tag_path = []

    row = [""] * len(headers)
    cell_path_row = [""] * len(headers)
    for i, (header, header_level, header_parent, index, header_ancestors_withoutParent) in enumerate(headers):  # Manual correction to export also attributes of root /// for i, (header, header_level, header_parent, index, header_grandParent, header_gGrandParent,header_ggGrandParent, header_gggGrandParent) in enumerate(headers):
        if header in attributes and header_level == 1:
            row[i] = (root.attrib[header])
            full_tag_path = header_dict[header,header_level, header_parent, index,header_ancestors_withoutParent] # full_tag_path = header_dict[header,header_level, header_parent, index,header_grandParent,header_gGrandParent]
            cell_path_row[i] = full_tag_path

    
    cell_path_matrix.append(cell_path_row)

    return row

def write_lines (headers, root, cell_path_matrix, writer, parent_map): ##### trying to correct write_lines by introducing greatgrandparents into header_dict, but now having problems with accessing the dictionary when writing stuff out. Logic might need some deep rethinking. I created keyword_ancestor_dict to simplify the process

    write_to_single_line = False
    keywordLevel = -1
    currentKeyword = None
    full_tag_path = []

    keyword_indexes = {keyword: 0 for keyword in keywords}

    row = [""] * len(headers)
    cell_path_row = [""] * len(headers)
    
    # New: keep track of the occurrence counts for each tag at the same level with the same parent.
    occurrence_counts = {}

    for item in root.findall('.//'):
        itemTag = item.tag
        itemLevel = get_level(item, parent_map)
        data = extract_xml_data(item,parent_map)

        ancestors,ancestor_tags,ancestor_tags_tuple,ancestor_tags_tuple_withoutParent, ancestor_tags_tuple_withoutGrandParent,parentTag,grandParentTag  = get_ancestors(item,parent_map)
        """
        ancestor_tags = [tag for (tag, level) in ancestors]

        if(len(ancestor_tags)>2):
            ancestor_tags_tuple_withoutParent = tuple(ancestor_tags[1:])
        else:
            ancestor_tags_tuple_withoutParent = None

        if(len(ancestor_tags)>1):
               ancestor_tags_tuple = tuple(ancestor_tags)
        else:
            ancestor_tags_tuple = None
        
        parent = parent_map.get(item)
        parentTag = parent.tag if parent is not None else None
        parentLevel = get_level(parent, parent_map) if parent is not None else None

        grandParent = parent_map.get(parent)
        grandParentTag = grandParent.tag if grandParent is not None else None
        grandParentLevel = get_level(grandParent, parent_map) if grandParent is not None else None

        gGrandParent= parent_map.get(grandParent)
        gGrandParentTag = gGrandParent.tag if gGrandParent is not None else None

        ggGrandParent = parent_map.get(gGrandParent)
        ggGrandParentTag = ggGrandParent.tag if ggGrandParent is not None else None
        """
        # Increment the occurrence count for this tag, level, and parent, and get the current count.
        occurrence_counts[(itemTag, itemLevel, ancestor_tags_tuple)] = occurrence_counts.get((itemTag, itemLevel, ancestor_tags_tuple), 0) + 1
        index = occurrence_counts[(itemTag, itemLevel, ancestor_tags_tuple)] - 1
        #indexP = occurrence_counts[(parentTag, parentLevel, grandParentTag,gGrandParentTag)]
        #indexGP = occurrence_counts[(grandParentTag, grandParentLevel, gGrandParentTag,ggGrandParentTag)]
            
        if (itemTag, itemLevel) in keywordsTuple:
            if write_to_single_line and (itemLevel <= keywordLevel or itemTag != currentKeyword):
                writer.writerow(row)
                cell_path_matrix.append(cell_path_row)
                row = [""] * len(headers)
                cell_path_row = [""] * len(headers)
                    
            write_to_single_line = True
            keywordLevel = itemLevel
            currentKeyword = itemTag
            keyword_indexes[currentKeyword] += 1

        elif itemLevel <= keywordLevel:
            if write_to_single_line:
                writer.writerow(row)
                cell_path_matrix.append(cell_path_row)
                row = [""] * len(headers)
                cell_path_row = [""] * len(headers)

            write_to_single_line = False
            keywordLevel = -1
            currentKeyword = None
                
        for i, (header, header_level, header_parent, header_index, header_ancestors_withoutParent) in enumerate(headers): # for i, (header, header_level, header_parent, header_index, header_grandParent, header_gGrandParent, header_ggGrandParent,header_gggGrandParent) in enumerate(headers):
            if header == itemTag and header_parent == parentTag: #if data only has the value itself
                shortCategory = True
            else:
                shortCategory = False

            parentLevel = get_level(header_parent, parent_map)
            if (((header,header_parent) in data) or shortCategory) and (itemLevel == header_level):
            #if ((header,header_parent) in data) and (itemLevel == header_level):
            #if ((header in data) and (itemLevel == header_level) and (parentTag == header_parent)):
            
                if(header != itemTag and header_parent == itemTag and header_ancestors_withoutParent == ancestor_tags_tuple): #if we found attributes of item (which means level is the same as parent) // if(header != itemTag and header_parent == itemTag and header_grandParent == parentTag and header_gGrandParent == grandParentTag):
                    """
                    try:
                        parentAncestor = keyword_ancestor_dict[(header_parent,header_level, header_grandParent,header_gGrandParent)][0]
                    except:
                        parentAncestor = None
                    try:
                        grandParentAncestor = keyword_ancestor_dict[(header_grandParent, header_level-1, header_gGrandParent, header_ggGrandParent)][0]
                    except:
                        grandParentAncestor = None
                    try:
                        gGrandParentAncestor = keyword_ancestor_dict[(header_gGrandParent, header_level-2, header_ggGrandParent, header_gggGrandParent)][0]
                    except:
                        gGrandParentAncestor = None

                    parentAncestorLevel = get_level(parentAncestor,parent_map) if parentAncestor is not None else None
                    grandParentAncestorLevel = get_level(grandParentAncestor,parent_map) if grandParentAncestor is not None else None
                    """
                    if(header == "RefID" and itemTag != "TypeRoot" and header_level == 3):
                        continue
                    else:
                        if(currentKeyword is not None): #if there is a keyword
                            if(parentAncestor is not None):#if parent also has  keyword
                                if(grandParentAncestor is not None):#if grandparent also has a keyword
                                    if(gGrandParentAncestor is not None):
                                        full_tag_path = header_dict[(f"{header}_{keyword_indexes[currentKeyword]}", header_level, f"{header_parent}_{keyword_indexes[parentAncestor]}", header_index,f"{header_grandParent}_{keyword_indexes[grandParentAncestor]}",f"{header_gGrandParent}_{keyword_indexes[gGrandParentAncestor]}")]
                                    else:
                                        full_tag_path = header_dict[(f"{header}_{keyword_indexes[currentKeyword]}", header_level, f"{header_parent}_{keyword_indexes[parentAncestor]}", header_index,f"{header_grandParent}_{keyword_indexes[grandParentAncestor]}",header_gGrandParent)]
                                else: # if grandparent does not have a keyword
                                    full_tag_path = header_dict[(f"{header}_{keyword_indexes[currentKeyword]}", header_level, f"{header_parent}_{keyword_indexes[parentAncestor]}", header_index,header_grandParent,header_gGrandParent)]
                            else: # if neither parent, nor grandparent has a keyword
                                full_tag_path = header_dict[(f"{header}_{keyword_indexes[currentKeyword]}", header_level, header_parent, header_index,header_grandParent,header_gGrandParent)]
                            
                        else: # if there are no keywords
                            
                                try:
                                    full_tag_path = header_dict[(header, header_level, header_parent, index, header_ancestors_withoutParent)] # full_tag_path = header_dict[(header, header_level, header_parent, index, header_grandParent,header_gGrandParent)]
                                except:
                                    continue
                            

                        if((header,header_parent) in data):
                            if isinstance(data[header,header_parent], dict):
                                row[i] = (json.dumps(data[header]['#text']))
                                cell_path_row[i] = full_tag_path
                            elif isinstance(data[header,header_parent], list):
                                flattened_dict = flatten_list(data[header,header_parent]['#text'])
                                row[i] = (json.dumps(flattened_dict))
                                cell_path_row[i] = full_tag_path
                            else:
                                row[i] = (data[header,header_parent])
                                cell_path_row[i] = full_tag_path
                else: #if we found the element itself
                    if(header == itemTag and header_parent == parentTag and header_grandParent == grandParentTag and header_gGrandParent == gGrandParentTag):
                        try:
                            parentAncestor = keyword_ancestor_dict[(header_parent,header_level-1, header_grandParent,header_gGrandParent)][0]
                        except:
                            parentAncestor = None
                        try:
                            grandParentAncestor = keyword_ancestor_dict[(header_grandParent, header_level-2, header_gGrandParent,header_ggGrandParent)][0]
                        except:
                            grandParentAncestor = None
                        try:
                            gGrandParentAncestor = keyword_ancestor_dict[(header_gGrandParent, header_level-3, header_ggGrandParent, header_gggGrandParent)][0]
                        except:
                            gGrandParentAncestor = None


                        try:
                            data['#text']
                        
                            if(currentKeyword is not None):
                                if(parentAncestor is not None):#if parent also has  keyword
                                    if(grandParentAncestor is not None):#if grandparent also has a keyword
                                        if(gGrandParentAncestor is not None):#if great grandparent also has a keyword
                                            if(currentKeyword == header):
                                                full_tag_path = header_dict[(f"{header}_{keyword_indexes[currentKeyword]}", header_level, f"{header_parent}_{keyword_indexes[parentAncestor]}", index,f"{header_grandParent}_{keyword_indexes[grandParentAncestor]}",f"{header_gGrandParent}_{keyword_indexes[gGrandParentAncestor]}")]
                                            else:
                                                full_tag_path = header_dict[(f"{header}_{keyword_indexes[currentKeyword]}", header_level, f"{header_parent}_{keyword_indexes[parentAncestor]}", header_index,f"{header_grandParent}_{keyword_indexes[grandParentAncestor]}",f"{header_gGrandParent}_{keyword_indexes[gGrandParentAncestor]}")]
                                        else: # if great grandparent does not have a keyword
                                            full_tag_path = header_dict[(f"{header}_{keyword_indexes[currentKeyword]}", header_level, f"{header_parent}_{keyword_indexes[parentAncestor]}", header_index,f"{header_grandParent}_{keyword_indexes[grandParentAncestor]}",header_gGrandParent)]
                                    else:
                                        full_tag_path = header_dict[(f"{header}_{keyword_indexes[currentKeyword]}", header_level, f"{header_parent}_{keyword_indexes[parentAncestor]}", header_index,header_grandParent,header_gGrandParent)] 
                                else: # if neither parent, nor grandparent has a keyword
                                    full_tag_path = header_dict[(f"{header}_{keyword_indexes[currentKeyword]}", header_level, header_parent, header_index,header_grandParent,header_gGrandParent)]
 
                               
                            else:
                                full_tag_path = header_dict[(header, header_level, header_parent, index, header_grandParent,header_gGrandParent)]

                            if isinstance(data['#text'], dict):
                                row[i] = (json.dumps(data['#text']))
                                cell_path_row[i] = full_tag_path
                            elif isinstance(data['#text'], list):
                                flattened_dict = flatten_list(data['#text'])
                                row[i] = (json.dumps(flattened_dict))
                                cell_path_row[i] = full_tag_path
                            else:
                                row[i] = (data['#text'])
                                cell_path_row[i] = full_tag_path
                        except:
                            continue

        rowNotEmpty = any(element != "" for element in row)

        if rowNotEmpty and not write_to_single_line:
            writer.writerow(row)
            cell_path_matrix.append(cell_path_row)
            row = [""] * len(headers)
            cell_path_row = [""] * len(headers)


def xml_to_csv(xml_file, csv_file):
    tree = ET.parse(xml_file)
    root = tree.getroot()

    parent_map = {c: p for p in tree.iter() for c in p}


    

    with open(csv_file, 'w', newline='', encoding= "UTF-8") as f:
        writer = csv.writer(f, delimiter= ";")

        headers  = []
        header_tags = []
        #header_parents = []
        headers, header_tags = get_headers(root, parent_map)
        writer.writerow(header_tags)

        row = []
        
        row = get_first_row(headers, row, root, cell_path_matrix)
        writer.writerow(row)

        write_lines(headers, root, cell_path_matrix, writer, parent_map)


                    
                    
                       


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





header_path_list = []


def csv_to_xml(xml_structure, csv_file, xml_file):
    with open(csv_file, 'r',encoding='utf-8') as f:
        reader = csv.reader(f, delimiter=";") #### kind of stupid, but if you save the csv file in excel, then delimiter is ";" and not "," as Python exports it
        headers = next(reader)

        # Create the XML structure using xml_structure
        root = create_xml_skeleton(xml_structure)

        # Create the header-path list
        #header_path_list = create_header_path_list(xml_structure)

        # Go through each row in the csv file and populate the XML with data
        for csv_row_index, row in enumerate(reader):
            for csv_cell_index, string in enumerate(row):
                populate_xml_element(root, csv_row_index,csv_cell_index, row, cell_path_matrix)


        tree = ET.ElementTree(root)
        tree.write(xml_file)

def navigate_to_element(element, tag_path_parts):  ##### the keyword and elemnt index logic does not function properly yet - this is probably the last bug we need to correct
    current_element = element
    keyword_counts = {keyword: 0 for keyword in keywords}

    for level, tag in enumerate(tag_path_parts,1):
        tag_parts = tag.split("_")
        actual_tag = tag_parts[0]
        keyword_idx = None
        element_idx = None

        # Identifying keyword and element indices
        for part in tag_parts[1:]:
            if part.startswith('idx'):
                element_idx = int(part[3:])
            else:
                keyword_idx = int(part)

        # Counting keywords
        if (actual_tag, level) in keywordsTuple:
            for child in current_element:
                if child.tag == actual_tag:
                    keyword_counts[actual_tag] += 1
                    if keyword_counts[actual_tag] == keyword_idx:
                        current_element = child
                        break
        else:# Counting elements
            count = -1
            for child in current_element:
                if child.tag == actual_tag:
                    count += 1
                    if count == element_idx:
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
        if(row[csv_cell_index] == '(Empty)'):
            current_element.text = ""
            current_element.set(attr, "")
        else:
            
            current_element.set(attr, row[csv_cell_index])
    else:
        # Set text
        if(row[csv_cell_index] == '(Empty)'):
            current_element.text = ""
        else:
            current_element.text = row[csv_cell_index]


def get_keyword_ancestors(root, parent_map, keywords_tuple):
    keyword_ancestors = {}

    def find_ancestor(item, keywords_tuple):
        current_keyword = (item.tag, get_level(item, parent_map))
        if current_keyword in keywords_tuple:
            return current_keyword
        parent = parent_map.get(item)
        if parent is not None:
            return find_ancestor(parent, keywords_tuple)
        return None

    def explore_node(node):
        has_keyword_ancestor = find_ancestor(node, keywords_tuple)
        parent = parent_map.get(node)
        grandparent = parent_map.get(parent) if parent is not None else None
        node_key = (
            node.tag,
            get_level(node, parent_map),
            parent.tag if parent is not None else None,
            grandparent.tag if grandparent is not None else None
        )
        a = node_key[0]
        b = node_key[1]
        if (a, b) in keywords_tuple:
            keyword_ancestors[node_key] = (a, b)
        else:
            keyword_ancestors[node_key] = has_keyword_ancestor
        for child in node:
            explore_node(child)

    explore_node(root)

    return keyword_ancestors











# Usage:
inputXML = "C:\\Users\\s.veress\\Desktop\\123.xml" #"Z:\\Mitarbeiter\\Szabolcs\\ArchicadVorlage\\AttributenStudie\\345_Linien.xml" Z:\\Mitarbeiter\\Szabolcs\\ArchicadVorlage\\Test.xml
outputCSV = 'C:\\Users\\s.veress\\Desktop\\output.csv'
outputXML = 'C:\\Users\\s.veress\\Desktop\\output.xml'

keywordsTuple = [("Layer",4), ("LineType",4), ("BuildingMaterial",4), ("CompositeWall",4), ("Material",4), ("Fill",4), ("Profile",4),("Attributes",3), ("PenTable", 4), ("Pens",5), ("Pen", 6), ("LayerCombination",4), ("LayerStatus",5), ("LayerIndex",5), ("Skin", 5),("SkineFaceLine",5)] 
#keywords = ["Layer", "LineType", "BuildingMaterial", "CompositeWall", "Material", "Fill", "Profile","Attributes", "PenTable", "Pens", "Pen"]# 
keywords = [p for (p,q) in keywordsTuple]
# Load the XML file
tree = ET.parse(inputXML)
root = tree.getroot()


# Create a map from child elements to their parent
parent_map = {c: p for p in tree.iter() for c in p}

# Build the dictionary
#header_dict key structure [itemTag, itemLevel, parentTag, itemIndex (if there are duplicates), grandParentTag]
header_dict = build_header_dict(root, parent_map)
keyword_ancestor_dict = get_keyword_ancestors(root, parent_map, keywordsTuple)



# Now you can use header_dict to get the full path of any element,
# given its name and depth.



# Extract the XML structure
xml_structure = extract_xml_structure(root)
#xml_structure = merge_duplicate_keys(xml_structure)

#allXMLLeaves = get_all_values(xml_structure)

# Print the extracted structure

#print(json.dumps(xml_structure, indent=4))
#print(xml_structure)

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


def is_header_in_data(header, data):
    if isinstance(data, dict):
        if header in data:
            return True
        for value in data.values():
            if is_header_in_data(header, value):
                return True
    return False

def flatten_list(data_list):
    """Flatten a list of dictionaries into a single dictionary."""
    flattened = {}
    for item in data_list:
        if isinstance(item, dict):
            flattened.update(item)
    return flattened