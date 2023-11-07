import csv
import xml.etree.ElementTree as ET


def save_xml_structure(root, library):
    """
    Recursively saves the XML structure in nested libraries.
    Each element tag is a key in the library, and its value is a dictionary containing the element's attributes and children.
    """
    library[root.tag] = {
        'attributes': root.attrib,
        'children': {}
    }

    for child in root:
        save_xml_structure(child, library[root.tag]['children'])


def construct_xml_element(tag, attributes, children):
    """
    Constructs an XML element with the given tag, attributes, and children.
    """
    element = ET.Element(tag, attributes)

    for child_tag, child_data in children.items():
        child_element = construct_xml_element(child_tag, child_data['attributes'], child_data['children'])
        element.append(child_element)

    return element


def reconstruct_xml(library, root_tag):
    """
    Reconstructs the XML structure based on the nested library and the root tag.
    """
    root_element = construct_xml_element(root_tag, library[root_tag]['attributes'], library[root_tag]['children'])
    return ET.ElementTree(root_element)


def main():
    # Read the XML file
    xml_file_path = 'path/to/input.xml'
    tree = ET.parse(xml_file_path)
    root = tree.getroot()

    # Create the library to store the XML structure
    library = {}

    # Save the XML structure in the library
    save_xml_structure(root, library)

    # Export the library to a CSV file
    csv_file_path = 'path/to/output.csv'
    with open(csv_file_path, 'w', newline='') as csv_file:
        writer = csv.writer(csv_file)
        writer.writerow(['Tag', 'Attributes', 'Children'])
        write_csv_row(root.tag, library[root.tag]['attributes'], library[root.tag]['children'], writer)

    # Read the CSV file and reconstruct the XML structure
    reconstructed_library = {}
    with open(csv_file_path, 'r') as csv_file:
        reader = csv.reader(csv_file)
        next(reader)  # Skip the header row
        for row in reader:
            tag = row[0]
            attributes = eval(row[1])  # Convert attributes string to a dictionary
            children = eval(row[2])  # Convert children string to a dictionary
            reconstructed_library[tag] = {
                'attributes': attributes,
                'children': children
            }

    # Reconstruct the XML structure from the library
    reconstructed_tree = reconstruct_xml(reconstructed_library, root.tag)

    # Save the reconstructed XML tree to a file
    reconstructed_xml_file_path = 'path/to/reconstructed.xml'
    reconstructed_tree.write(reconstructed_xml_file_path)


def write_csv_row(tag, attributes, children, writer):
    """
    Recursively writes a CSV row for each XML element in the library.
    """
    writer.writerow([tag, repr(attributes), repr(children)])

    for child_tag, child_data in children.items():
        write_csv_row(child_tag, child_data['attributes'], child_data['children'], writer)


if __name__ == '__main__':
    main()

