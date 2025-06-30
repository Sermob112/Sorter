from abc import ABC, abstractmethod
from typing import List
import os
import re
import xml.etree.ElementTree as ET
from xml.dom import minidom

class FileSystemWalker(ABC):
    @abstractmethod
    def collect_files(self, directory: str) -> List[str]:
        pass
class LocalFileSystemWalker(FileSystemWalker):
    def collect_files(self, directory: str) -> List[str]:
        file_list = []
        for root, _, files in os.walk(directory):
            for file in files:
                full_path = os.path.join(root, file)
                file_list.append(full_path)
        return file_list
class FileListProcessor:
    def __init__(self, file_list: List[str]):
        self.file_list = file_list

    def print_files(self):
        for file in self.file_list:
            print(file)

    def save_to_file(self, output_file: str):
        with open(output_file, "w", encoding="utf-8") as f:
            for file in self.file_list:
                f.write(file + "\n")
    
    def filter_by_extension(self, extension: str) -> List[str]:
        return [file for file in self.file_list if file.endswith(extension)]

    def save_only_filenames(self) -> List[str]:
        """
        Возвращает список имен файлов (без полного пути).
        """
        filenames = []
        for file_path in self.file_list:
            filename = os.path.basename(file_path)
            filenames.append(filename)
        return filenames
            
    def create_assembly_xml(self, filenames: List[str], output_xml: str):
        """
        Создает XML-файл с элементами <Assembly> на основе списка файлов.
        Добавляет стандартную шапку XML.
        """
    
        root = ET.Element("MBOM")

      
        spr = ET.SubElement(root, "Spr")


        docs = ET.SubElement(spr, "Docs")
        ET.SubElement(docs, "TechDocument", 
                    ident="PR600.362111.208", 
                    descr="Объемная секция коробчатого флора в р-не 65 - 70 шп.", 
                    code="СП", 
                    doctype="РКД02", 
                    filename="PR600.362111.208.xlsx", 
                    format="А3")
        ET.SubElement(docs, "TechDocument", 
                    ident="PR600.362111.208", 
                    descr="Объемная секция коробчатого флора в р-не 65 - 70 шп.", 
                    code="СБ", 
                    doctype="РКД02", 
                    filename="PR600.362111.208_Объемная секция коробчатого флора в районе 65---70 шп_2016.12.09.dwg", 
                    format="А3")

        doc_codes = ET.SubElement(spr, "DocCodes")
        ET.SubElement(doc_codes, "DocCode", code="СП", descr="Спецификация")
        ET.SubElement(doc_codes, "DocCode", code="СБ", descr="Сборочный чертёж")

        doc_formats = ET.SubElement(spr, "DocFormats")
        ET.SubElement(doc_formats, "DocFormat", code="А3")

        units = ET.SubElement(spr, "Units")
        ET.SubElement(units, "Unit", code="796", descr="штука", short_name="шт.")

        standards = ET.SubElement(spr, "Standards")
        ET.SubElement(standards, "Standard", ident="ГОСТ 19903-2015", descr="Прокат листовой горячекатаный. Сортамент")
        ET.SubElement(standards, "Standard", ident="ГОСТ Р 52927-2015", descr="Прокат для судостроения из стали нормальной, повышенной и высокой прочности. Технические условия")
        ET.SubElement(standards, "Standard", ident="ГОСТ 21937-76", descr="Полособульб горячекатаный несимметричный для судостроения. Сортамент")

        mat_marks = ET.SubElement(spr, "MatMarks")
        ET.SubElement(mat_marks, "MatMark", ident="РСА", descr="Сталь нормальной прочности, изготавливаемая под надзором Регистра", ntd="ГОСТ Р 52927-2015")

    
        file_groups = {}
        for filename in filenames:
            name, ext = os.path.splitext(filename)
            if name not in file_groups:
                file_groups[name] = []
            file_groups[name].append(ext.lower())  
      
        for name, extensions in file_groups.items():
           
            parts = re.split(r"[_ -]", name, maxsplit=1)  
            if len(parts) == 2:
                ident, descr = parts  
            else:
                ident = name  
                descr = name 

            assembly = ET.SubElement(root, "Assembly")
            assembly.set("ident", name)
            assembly.set("descr", descr)  

            # Добавляем элементы <RefDoc>
            for ext in extensions:
                if ext in [".dwg", ".frw"]:
                    code = "СБ"  # Для .dwg и .FRW используем код "СБ"
                else:
                    code = "СП"  # Для остальных расширений используем код "СП"

                ref_doc = ET.SubElement(assembly, "RefDoc")
                ref_doc.set("ident", ident)
                ref_doc.set("code", code)

     
        xml_str = ET.tostring(root, encoding="utf-8", xml_declaration=True)
        dom = minidom.parseString(xml_str)
        pretty_xml = dom.toprettyxml(indent="  ", encoding="utf-8")

        with open(output_xml, "wb") as f:
            f.write(pretty_xml)

# Используем конкретную реализацию для обхода файловой системы
# walker: FileSystemWalker = LocalFileSystemWalker()
# root_directory = "XML test/HB600"
# # Собираем список файлов
# files = walker.collect_files(root_directory)

# # Обрабатываем список файлов
# processor = FileListProcessor(files)
# processor.print_files()  # Выводим файлы на экран
# # processor.save_to_file("output.txt")  
# files = processor.save_only_filenames("output_filenames.txt")  
# processor.create_assembly_xml(files, "XML_test.xml")