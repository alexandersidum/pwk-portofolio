import xml.etree.ElementTree as ET

class XMLReader:
    def __init__(self, xmlFile):
        self.xmlFile = xmlFile

    # xmlFile = string path file
    def readXml(self):
        it = ET.iterparse(self.xmlFile)
        # Menghilangkan namespace
        for _, el in it:
            prefix, has_namespace, postfix = el.tag.partition('}')
            if has_namespace:
                el.tag = postfix
        root = it.root
        return root


    def getRows(self,parsedXml):
        worksheets = parsedXml.find('Worksheet')
        tables = worksheets.find('Table')
        rows = tables.findall('Row')
        return rows


    # Outputnya multidimension array
    # index 0 = ID, 1 = Nrp - nama, 2 = Ujian tulis? , 3 = Critical Review? , 4 = Praktikum, 5 = Tugas Besar
    def getFinalData(self,parsedXml):
        rows = self.getRows(parsedXml)
        output = []
        for row in rows:
            pass
            temp = []
            for cell in row.findall('Cell'):
                for data in cell.findall('Data'):
                    temp.append(data.text)
            output.append(temp)
        del output[:2]
        return output

    def getData(self):
        parsedXml = self.readXml()
        finalData = self.getFinalData(parsedXml)
        return finalData