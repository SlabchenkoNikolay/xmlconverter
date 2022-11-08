import xml.etree.ElementTree as ET  # подключаем ElementTree для работы со структурой
from xml.dom import minidom  # подключаем минидом для работы с парсингом документа
import uuid  # универсальный уникальный идентификатор генерация( он же ГУид, не все гуиды - ииды)
import sys  # sys нужен для передачи argv в QApplication
from PyQt5 import QtCore, QtGui, QtWidgets
import openpyxl

# Парсинг файла дизайна для взятия значений переменных в программу
root_design = ET.parse('design_xml.ui').getroot()
# ET.dump(root_design) # Проверка корректности открытия коренного тега
mass = []  # Список текстовых данных, содержит по порядку значения из UI документа дизайна
for tag in root_design.findall('widget/widget/widget/property/string'):
    mass.append(tag.text)
# for i in range(24):
#   print(i, mass[i])

# Строка даты создания документа из дизайн файла
date_when_created = []
for tag in root_design.findall('widget/widget/widget/property/datetime/year'):
    date_when_created.append(tag.text)
for tag in root_design.findall('widget/widget/widget/property/datetime/month'):
    date_when_created.append(tag.text)
for tag in root_design.findall('widget/widget/widget/property/datetime/day'):
    date_when_created.append(tag.text)

data_sozd = date_when_created[0] + "-" + date_when_created[1] + "-" + date_when_created[2]
print(data_sozd)

backslash = "\\"
# Взятие данных их EXCEL документа
book = openpyxl.open('XML/Каталоги 21/Сочи/Дуб голубой 2 дерева П.xlsx')
sheet = book.active
AreaText = '28'
InaccuracyText = '18'
# Генерация гуида
# guid_num = uuid.uuid4()

guid_num = uuid.uuid4()
guid_num_str = str(guid_num)

if float(sheet[2][2].value) < 2000000:
    msk_text = 'зона 1'
    CsId = 'Id1e83be3b-46c2-429c-a0c1-184bdbf6b7f9'
else:
    msk_text = 'зона 2'
    CsId = 'Id2e87c032-c1c7-4ffc-ba82-c76280ad18fc'
print(guid_num_str)



# Переменные
oopt_name = 'TerritoryToGKN_' + guid_num_str

AppliedFileName1 = 'План охранной зоны'
AppliedFileName2 = 'Охранная зона'
AppliedFileName3 = 'Доверенность'

# Создание скелета XML файла
territorytogkn = ET.Element('TerritoryToGKN xmlns:Simple2="urn://x-artefacts-rosreestr-ru/commons/simple-types/2.0.1" '
                            'xmlns:Simple1="urn://x-artefacts-rosreestr-ru/commons/simple-types/1.0" '
                            'xmlns:tns="urn://x-artefacts-smev-gov-ru/supplementary/commons/1.0.1" '
                            'xmlns:Simple4="urn://x-artefacts-rosreestr-ru/commons/simple-types/4.1.1" '
                            'xmlns:Simple7="urn://x-artefacts-rosreestr-ru/commons/simple-types/7.0.1" '
                            'xmlns:CadEng4="urn://x-artefacts-rosreestr-ru/commons/complex-types/cadastral-engineer/4'
                            '.1.1" xmlns:Spa1="urn://x-artefacts-rosreestr-ru/commons/complex-types/entity-spatial/2'
                            '.0.1" xmlns:dAl3="urn://x-artefacts-rosreestr-ru/commons/directories/all-documents/3.0.2'
                            '" xmlns:DocI5="urn://x-artefacts-rosreestr-ru/commons/complex-types/document-info/5.0.1" '
                            'xmlns="urn://x-artefacts-rosreestr-ru/incoming/territory-to-gkn/1.0.4" '
                            'xmlns:dUn1="urn://x-artefacts-rosreestr-ru/commons/directories/unit/1.0.1" '
                            'xmlns:Zon4="urn://x-artefacts-rosreestr-ru/commons/complex-types/zone/4.2.2" '
                            'NameSoftware="Полигон Про" VersionSoftware="5.1.1.21" GUID="' + guid_num_str + '"',
                            )  # Главный элемент TerritoryToGKN,кучу свойств дописать к нему
title = ET.SubElement(territorytogkn, 'Title')  # Без свойств

clients = ET.SubElement(title, 'Clients')  # Без свойств
client = ET.SubElement(clients, 'Client', Date=data_sozd)  # Дата создания документа в свойство дописать
governance = ET.SubElement(client, 'Governance')
gover_name = ET.SubElement(governance, 'Name').text = mass[7]

agent = ET.SubElement(governance, 'Agent')
appointment = ET.SubElement(agent, 'Appointment').text = mass[1]
tnsfamilyname = ET.SubElement(agent, 'tns:FamilyName').text = mass[2]
tnsfirstname = ET.SubElement(agent, 'tns:FirstName').text = mass[3]
tnspatronymic = ET.SubElement(agent, 'tns:Patronymic').text = mass[4]

contractor = ET.SubElement(title, 'Contractor')
organisationcontractor = ET.SubElement(contractor, 'OrganisationContractor')
contractorname = ET.SubElement(organisationcontractor, 'Name').text = mass[5]
codeogrn = ET.SubElement(organisationcontractor, 'CodeOGRN').text = mass[6]
telephone = ET.SubElement(organisationcontractor, 'Telephone').text = mass[8]
address = ET.SubElement(organisationcontractor, 'Address').text = mass[9]
agentOC = ET.SubElement(organisationcontractor, 'Agent')

AgentOCAppointment = ET.SubElement(agentOC, 'Appointment').text = mass[10]
AgentOCtnsfamilyname = ET.SubElement(agentOC, 'tns:FamilyName').text = mass[11]
AgentOCtnsfirstname = ET.SubElement(agentOC, 'tns:FirstName').text = mass[12]
AgentOCtnspatronymic = ET.SubElement(agentOC, 'tns:Patronymic').text = mass[13]
AgentOCAttorneyDocument = ET.SubElement(agentOC, 'AttorneyDocument')
AgentOCAttorneyDocumentDocI5CodeDocument = ET.SubElement(AgentOCAttorneyDocument, 'DocI5:CodeDocument').text = mass[14]
AgentOCAttorneyDocumentDocI5Name = ET.SubElement(AgentOCAttorneyDocument, 'DocI5:Name').text = mass[15]
AgentOCAttorneyDocumentDocI5Number = ET.SubElement(AgentOCAttorneyDocument, 'DocI5:Number').text = mass[16]
AgentOCAttorneyDocumentDocI5Date = ET.SubElement(AgentOCAttorneyDocument, 'DocI5:Date').text = mass[17]
AgentOCAttorneyDocumentDocI5IssueOrgan = ET.SubElement(AgentOCAttorneyDocument, 'DocI5:IssueOrgan').text = mass[18]
# Ссылка на изображение документа прописывается в свойстве элемента
AgentOCAttorneyDocumentDocI5AppliedFile = ET.SubElement(AgentOCAttorneyDocument,
                                                        'DocI5:AppliedFile Name="Images\\' + AppliedFileName3 + '.pdf" Kind="01" ')

Coordinations = ET.SubElement(title, 'Coordinations')
Coordination = ET.SubElement(Coordinations, 'Coordination')
CoordinationName = ET.SubElement(Coordination, 'Name').text = mass[19]
CoordinationOfficial = ET.SubElement(Coordination, 'Official')
CoordinationOfficialAppointemnt = ET.SubElement(CoordinationOfficial, 'Appointment').text = mass[20]
CoordinationOfficialtnsFamilyName = ET.SubElement(CoordinationOfficial, 'tns:FamilyName').text = mass[21]
CoordinationOfficialtnsFirstName = ET.SubElement(CoordinationOfficial, 'tns:FirstName').text = mass[22]
CoordinationOfficialtnsPatronymic = ET.SubElement(CoordinationOfficial, 'tns:Patronymic').text = mass[23]

# ЦИКЛИЧНО ЗАБИВАЕМ ТОЧКИ ИЗ ФАЙЛА EXCEL
EntitySpatialEntSys = ET.SubElement(territorytogkn, 'EntitySpatial EntSys=' + '"' + CsId + '"')

# Считаем количество SpatialElement и добавляем в массив точку начала, точку окончания айдишника и точку конечную
border_point = [1]
border_id = 1
SpatialElement_counter = 1
for row in range(2, sheet.max_row + 1):
    sheet_border_id = int(sheet[row][3].value)
    if border_id != sheet_border_id:
        SpatialElement_counter += 1
        border_id = sheet_border_id
        border_point.append(row-1)
border_point.append(sheet.max_row)
ot = 0
do = 1
# print(sheet[2][0].value, sheet[2][1].value, sheet[2][2].value, sheet[2][3].value)
for prohod in range(len(border_point)-1):
    SpaSpatialElement = ET.SubElement(EntitySpatialEntSys, 'Spa1:SpatialElement')
    SuNmb = 0

    if do < len(border_point):

        for row in range(border_point[ot]+1, border_point[do]+1):
            SuNmb += 1
            # print(sheet[row][0].value, sheet[row][1].value, sheet[row][2].value, sheet[row][3].value)
            x = '%.2f' % float(sheet[row][1].value)
            y = '%.2f' % float(sheet[row][2].value)
            # print(sheet[row][0].value, str(x), str(y))
            SpaElementUnit = ET.SubElement(SpaSpatialElement,
                                           'Spa1:SpelementUnit TypeUnit="Точка" SuNmb="' + str(SuNmb) + '"')
            SpaOrdinate = ET.SubElement(SpaElementUnit,
                                        'Spa1:Ordinate X="' + str(x) + '" Y="' + str(y) + '" NumGeopoint="' + str(sheet[row][0].value) + '"' + ' GeopointOpred="692003000000" DeltaGeopoint="1.00"')
        SuNmb += 1
        x = '%.2f' % float(sheet[border_point[ot]+1][1].value)
        y = '%.2f' % float(sheet[border_point[ot]+1][2].value)
        SpaElementUnit = ET.SubElement(SpaSpatialElement,
                                       'Spa1:SpelementUnit TypeUnit="Точка" SuNmb="' + str(SuNmb) + '"')
        SpaOrdinate = ET.SubElement(SpaElementUnit,
                                    'Spa1:Ordinate X="' + str(x) + '" Y="' + str(
                                        y) + '" NumGeopoint="' + str(sheet[border_point[ot]+1][
                                                                                           0].value) + '"' + ' GeopointOpred="692003000000" DeltaGeopoint="1.00"')

        ot += 1
        do += 1









# Area
Area = ET.SubElement(territorytogkn, 'Area')
AreaMeter = ET.SubElement(Area, 'AreaMeter')
AreaInAreaMeter = ET.SubElement(AreaMeter, "Area").text = AreaText
UnitInAreaMeter = ET.SubElement(AreaMeter, "Unit").text = '055'
InaccuracyInAreaMeter = ET.SubElement(AreaMeter, "Inaccuracy").text = InaccuracyText

# CoordSystem
CoordSystem = ET.SubElement(territorytogkn, 'CoordSystems')
SpaCoordSystem = ET.SubElement(CoordSystem, 'Spa1:CoordSystem Name=' + '"МСК-23, '+ msk_text + '"' + ' CsId="' + CsId + '"')

# Diagram
Diagram = ET.SubElement(territorytogkn, 'Diagram')
AppliedFile1 = ET.SubElement(Diagram, 'AppliedFile Kind="01" Name="Images\\' + AppliedFileName1 + '.pdf"')
AppliedFile2 = ET.SubElement(Diagram, 'AppliedFile Kind="01" Name="Images\\' + AppliedFileName2 + '.pdf"')
AppliedFile3 = ET.SubElement(Diagram, 'AppliedFile Kind="01" Name="Images\\' + AppliedFileName3 + '.pdf"')

tree = ET.ElementTree(territorytogkn)
tree.write(oopt_name + ".xml", encoding="utf-8", xml_declaration=True)

# ET.dump(territorytogkn) #Вывод дерева в ком строке
