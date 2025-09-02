# -*- coding: utf-8 -*-
"""
Created on Wed Aug 27 10:55:44 2025

@author: Administrator
Simatic SD to .xlsx 
"""

# Read Default Table xml 
import xml.etree.ElementTree as ET


tree = ET.parse('Enter the file Path of Your XML file, / is acceptable')
root = tree.getroot()


# Example to access the child node of xml, we interested in Name, Logical Address and Data Type, and V

Tag_name_elements = root.findall(".//SW.Tags.PlcTag/AttributeList/Name")
logical_addr_elements = root.findall(".//SW.Tags.PlcTag/AttributeList/LogicalAddress")
Data_type_elements = root.findall(".//SW.Tags.PlcTag/AttributeList/DataTypeName")

# Only tag in English is considered

Comment_elements= root.findall(".//SW.Tags.PlcTag/ObjectList/MultilingualText[@CompositionName='Comment']/ObjectList/MultilingualTextItem/AttributeList")

Eng_Only = []
for i in Comment_elements:
    # print(i.find("./Culture").text)
    if i.find("./Culture").text=='en-US':
        Eng_Only.append(i.find("./Text"))
        #print(i.find("./Text").text)
Comment_elements=Eng_Only

# The following code is used to validate the xml child path
# for item_element in cveList:
#       print(f"Name content: {item_element.text}")
