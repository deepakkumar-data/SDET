import xml.etree.ElementTree as ET
import pandas as pd
import time


class XML2DataFrame:

    def __init__(self, xml_data):
        self.root = ET.XML(xml_data)

    def parse_root(self, root):
        return [self.parse_element(child) for child in iter(root)]

    def parse_element(self, element, parsed=None):
        if parsed is None:
            parsed = dict()
        for key in element.keys():
            parsed[key] = element.attrib.get(key)
        if element.text:
            parsed[element.tag] = element.text
        for child in list(element):
            self.parse_element(child, parsed)
        return parsed

    def process_data(self):
        structure_data = self.parse_root(self.root)
        return pd.DataFrame(structure_data)
# xml_data = ET.parse('Sample.xml', ET.XMLParser(encoding='utf-8')) 
print 'XML to Excel Converter using Python'


while True:
    filename = raw_input("Enter the filename of the XML file (without extension): ")
    try:
        tree = ET.parse(filename+'.xml')
        break
    except:
        print 'File not found. Please try again...!!!!'


    # tree = ET.parse('Sample.xml')
root = tree.getroot()
xml_data=ET.tostring(root, encoding='utf8', method='xml')
xml2df = XML2DataFrame(xml_data)
xml_dataframe = xml2df.process_data()
# print xml_dataframe

timestr = time.strftime("%Y%m%d-%H%M%S")
writer = pd.ExcelWriter('XML_TO_EXCEL_Output_'+time.strftime("%Y%m%d-%H%M%S")+'.xlsx')
xml_dataframe.to_excel(writer,'XMLDATA',index=False)
worksheet= writer.sheets['XMLDATA']
worksheet.set_column('A:Z',50)
writer.save()

raw_input('*********************Hit Enter to Exit***********************')
print 'Exiting Tool....!!!!'
