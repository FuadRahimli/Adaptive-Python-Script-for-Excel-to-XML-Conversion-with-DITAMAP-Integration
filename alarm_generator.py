from xml.dom.minidom import Document, DOMImplementation
import pandas as pd
import os
import sys
import codecs
from xml.dom import minidom
import numpy as np

if not os.path.exists("topics/reference_alarms"):
    os.makedirs("topics/reference_alarms")
 

def prepare_xml (group):
    alarmIdNetAct = group ["alarmId NetAct"].iloc[0]
    imp = DOMImplementation ()
    doctype = imp.createDocumentType ("concept", "-//OASIS//DTD DITA Concept//EN", "concept.dtd")
    doc = Document ()
    doc.appendChild (doctype)
    concept = doc.createElement ("concept")
    concept.setAttribute ("id", f"_{alarmIdNetAct}")
    title = doc.createElement ("title")
    title_text = doc.createTextNode (f"{alarmIdNetAct}")
    title.appendChild(title_text)
    concept.appendChild(title)

    conbody = doc.createElement ("conbody")
    
    for key in group.columns:
        p = doc.createElement ("p")
        p_text = doc.createTextNode (key)
        p.appendChild (p_text)
        conbody.appendChild(p)
        if len (group[key].unique()) > 1:
            for value in group[key].unique ():
                codeblock = doc.createElement ("codeblock") 
                if pd.isna(value):
                    codeblock_text = doc.createTextNode ("")
                    codeblock.appendChild (codeblock_text) 
                elif isinstance (value, (float, np.floating)) and value.is_integer():
                    codeblock_text = doc.createTextNode (str(int(value)))
                    codeblock.appendChild (codeblock_text)  
                else:
                    codeblock_text = doc.createTextNode (str(value))
                    codeblock.appendChild (codeblock_text)
                conbody.appendChild (codeblock)
        else:
            codeblock = doc.createElement ("codeblock")
            if pd.isna (group[key].iloc[0]):
                codeblock_text = doc.createTextNode ("")
                codeblock.appendChild(codeblock_text)
            elif isinstance (group[key].iloc[0], (float, np.floating)) and group[key].iloc[0].is_integer():
                codeblock_text = doc.createTextNode(str(int(group[key].iloc[0])))
            else:
                codeblock_text = doc.createTextNode(str(group[key].iloc[0]))
                codeblock.appendChild(codeblock_text)
            conbody.appendChild(codeblock)

        concept.appendChild(conbody)
        doc.appendChild(concept)

        
        with codecs.open(f"topics/reference_alarms/alarms_{alarmIdNetAct}.xml", 'w', encoding='UTF-8') as f:
             doc.writexml (f,indent = "\t", addindent = '\t', newl = '\n', encoding = 'UTF-8')
        f.close()

def read_excel_data (file_name):
    try:
        df = pd.read_excel (file_name, sheet_name = "Alarm_list")
    except ValueError:
        raise ValueError ("'Alarm_list' sheet not found in file. Please check the sheet name or check if it exists!")
        df = None

    return df

if __name__ == '__main__':
    try:
        file_name = sys.argv[1]
        df = read_excel_data (file_name)
    except IndexError:
        print("Please provide a filename as a command-line argument!")
        sys.exit(1)
    except FileNotFoundError:
        print(f"File '{file_name}' does not exist. Please recheck for any mistypes and if the file exists at all!")
        sys.exit(1)
    except ValueError as e:
        print (str(e))
        sys.exit(1)

    if df is not None:
        for alarmIdNetAct, group in df.groupby ("alarmId NetAct"):
            prepare_xml (group)


#How to create a DITA Map

def prepare_ditamap(alarmIdNetAct):
    imp = DOMImplementation()
    doctype = imp.createDocumentType ("map", "-//OASIS//DTD DITA Map//EN", "map.dtd")
    doc = minidom.Document()
    doc.appendChild(doctype)

    map = doc.createElement ("map")
    map.setAttribute ('\n'"base", 
    """ pdm-category(Core Network) pdm-subcategory(Mobile Circuit Core) pdm-product(Open Mobile Softswitch)  """)    
    map.setAttribute ('\n'"id","troubleshoot_mss_alarms")
    map.setAttribute ("outputclass", "sans-numbering")
    doc.appendChild(map)

    title = doc.createElement ("title")
    title_text = doc.createTextNode ("Alarm List for CNCS")
    title.appendChild (title_text)
    map.appendChild(title)

    topicmeta = doc.createElement ("topicmeta")

    othermeta = doc.createElement ("othermeta")
    othermeta.setAttribute ("content", "reference_alarms_DN1000111289")
    othermeta.setAttribute ("name", "Filename")
    topicmeta.appendChild (othermeta)
    
    othermeta = doc.createElement ("othermeta")
    othermeta.setAttribute ("content", "document")
    othermeta.setAttribute ("name", "maptype")
    topicmeta.appendChild (othermeta)

    othermeta = doc.createElement ("othermeta")
    othermeta.setAttribute ("content", "DN1000111289")
    othermeta.setAttribute ("name", "docnumber")
    topicmeta.appendChild (othermeta)

    othermeta = doc.createElement ("othermeta")
    othermeta.setAttribute ("content", "270291757")
    othermeta.setAttribute ("name", "chronicleid")
    topicmeta.appendChild (othermeta)

    othermeta = doc.createElement ("othermeta")
    othermeta.setAttribute ("content", "1-0-0")
    othermeta.setAttribute ("name", "issue")
    topicmeta.appendChild (othermeta)

    map.appendChild(topicmeta)

    
    topicref1 = doc.createElement ("topicref")
    topicref1.setAttribute ("type", "concept")
    topicref1.setAttribute ("href", "topics/reference_alarms/alarm_introduction.xml")
    topicref1.setAttribute ("keys", "alarm_intro")
    

    topicref2 = doc.createElement ("topicref")
    topicref2.setAttribute ("type", "concept")
    topicref2.setAttribute ("href", "topics/reference_alarms/alarm_description.xml")
    topicref2.setAttribute ("keys", "alarm_desc")
    topicref1.appendChild(topicref2)
    
    map.appendChild(topicref1)

    topicref3 = doc.createElement ("topicref")
    topicref3.setAttribute ("type", "concept")
    topicref3.setAttribute ("href", "topics/reference_alarms/alarm_list.xml")
    topicref3.setAttribute ("keys", "alarm_list")

    map.appendChild(topicref3)

    for alarmIdNetAct, group in df.groupby("alarmId NetAct"):
        topicref = doc.createElement("topicref")
        topicref.setAttribute("type", "concept")
        topicref.setAttribute("keys", f"alarms_{alarmIdNetAct}")
        topicref.setAttribute("href", f"topics/reference_alarms/alarms_{alarmIdNetAct}.xml")
        topicref3.appendChild(topicref)

    map.appendChild(topicref3)
    doc.appendChild(map)
    
    
    with codecs.open(f"reference_alarms.ditamap", 'w', encoding='UTF-8') as f:
        doc.writexml(f, indent="\t", addindent='\t', newl='\n', encoding='UTF-8')
    f.close()
         
prepare_ditamap(alarmIdNetAct="alarmId NetAct")
        

