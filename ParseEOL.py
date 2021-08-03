from openpyxl import Workbook

myFile = r'\\vwoachsfile01\assembly\Departments\Assembly Launch\Battery Plant Status\Pilot Hall\IOL_EOL_Leak Docs\PVS\288 TX6PVS0340_2021_06_16_180429.ini'
input_file = open(myFile,'r')

#Variable Setup
#Master dictionary
Master = {}

#Loop through each line in file
while input_file:
    # Read next line
    currentLine = input_file.readline()
    #Check if EOF
    if currentLine == "":
        break
    #Identify header (will be separate dictionary entry in master dictionary
    if currentLine[0] == "[":
        header = currentLine.strip()
        #Create a blank dictionary to temporarily save to before adding to master
        myDict = {}
        dictName = currentLine

        #Index to next line
        currentLine = input_file.readline()

        while currentLine.strip() != "":

            #If there is a ";" in the string
            if ";" in currentLine.strip():
                data = currentLine[currentLine.find('"') + 1:].strip()
                data = data[:-1]
                # Create list of data/value pair
                myData = data.split(";")
                myDict[myData[0]] = myData[1]
            # If there is NOT a ";" in the string
            else:
                #      0 = "IO"
                data = currentLine
                myData = currentLine.split("=")
                myDict[myData[0].strip(" ")] = myData[1].strip().strip(" ")

            #Index to the next line at end of loop
            currentLine = input_file.readline()

        Master[header] = myDict

#Print each dictionary
#for (key,value) in Master.items():
    #print(key, value)

###############################################
#Print to Excel

wb = Workbook()
sheet = wb.active
count = 0

#General Data
battery = Master["[Pruefling]"]["Seriennummer"][7:15]
partNum = Master["[Pruefling]"]["Teilenummer"]
batType = Master["[Pruefling]"]["Batteriesystemtyp"]

#PreCheck
i_DTC_Count = Master["[SW_Steuergeraete_Allgemeine Informationen]"]["uPruefungszaehlerIst"]

#Software Setpoint Check
initBCMe = Master["[SW_Steuergeraete_Pruefergebnisse]"]["SW_BMCe_Get"]
initCMC1 = Master["[SW_Steuergeraete_Pruefergebnisse]"]["SW_CMC1_Get"]
initCMC2 = Master["[SW_Steuergeraete_Pruefergebnisse]"]["SW_CMC2_Get"]
initCMC3 = Master["[SW_Steuergeraete_Pruefergebnisse]"]["SW_CMC3_Get"]
initBL_BCMe = Master["[SW_Steuergeraete_Pruefergebnisse]"]["BL_BMCe_Get"]
initBL_CMC1 = Master["[SW_Steuergeraete_Pruefergebnisse]"]["BL_CMC1_Get"]
initBL_CMC2 = Master["[SW_Steuergeraete_Pruefergebnisse]"]["BL_CMC2_Get"]
initBL_CMC3 = Master["[SW_Steuergeraete_Pruefergebnisse]"]["BL_CMC3_Get"]

#Default parameters set
batConfigPset = Master["[SW_Steuergeraete_Pruefergebnisse]"]["DatensatzVersion_Batteriekonfig_Get"]
batVehiclePset = Master["[SW_Steuergeraete_Pruefergebnisse]"]["DatensatzVersion_Fahrzeug_Get"]
batVehiclePset = Master["[SW_Steuergeraete_Pruefergebnisse]"]["DatensatzVersion_Fahrzeug_Get"]
batTargetMarketPset = Master["[SW_Steuergeraete_Pruefergebnisse]"]["DatensatzVersion_Zielmarkt_Get"]
batConfigNamePset = Master["[SW_Steuergeraete_Pruefergebnisse]"]["DatensatzName_Batteriekonfig_Get"]
batThermoPset = Master["[SW_Steuergeraete_Pruefergebnisse]"]["DatensatzName_Thermo_Get"]
batNameVehiclePset = Master["[SW_Steuergeraete_Pruefergebnisse]"]["DatensatzName_Fahrzeug_Get"]
batNameTargetMarket = Master["[SW_Steuergeraete_Pruefergebnisse]"]["DatensatzName_Zielmarkt_Get"]

#Hardware PN
HW_NumCMC1 = Master["[SW_Steuergeraete_Pruefergebnisse]"]["F1A3_CMC1_Get"]
HW_NumCMC2 = Master["[SW_Steuergeraete_Pruefergebnisse]"]["F1A3_CMC2_Get"]
HW_NumCMC3 = Master["[SW_Steuergeraete_Pruefergebnisse]"]["F1A3_CMC3_Get"]
serial_CMC1 = Master["[SW_Steuergeraete_Pruefergebnisse]"]["SerialNr_CMC1_Get"]
serial_CMC2 = Master["[SW_Steuergeraete_Pruefergebnisse]"]["SerialNr_CMC2_Get"]
serial_CMC3 = Master["[SW_Steuergeraete_Pruefergebnisse]"]["SerialNr_CMC3_Get"]
HW_NumBMCe = Master["[SW_Steuergeraete_Pruefergebnisse]"]["F1A3_BMCe_Get"]
HW_serial_BMCe = Master["[SW_Steuergeraete_Pruefergebnisse]"]["F191_BMCe_Get"]
SW_serial_BMCe = Master["[SW_Steuergeraete_Pruefergebnisse]"]["F187_BMCe_Get"]
HW_serial_CMC1 = Master["[SW_Steuergeraete_Pruefergebnisse]"]["F191_CMC1_Get"]
HW_serial_CMC2 = Master["[SW_Steuergeraete_Pruefergebnisse]"]["F191_CMC2_Get"]
HW_serial_CMC3 = Master["[SW_Steuergeraete_Pruefergebnisse]"]["F191_CMC3_Get"]
SW_serial_CMC1 = Master["[SW_Steuergeraete_Pruefergebnisse]"]["F187_CMC1_Get"]
SW_serial_CMC2 = Master["[SW_Steuergeraete_Pruefergebnisse]"]["F187_CMC2_Get"]
SW_serial_CMC3 = Master["[SW_Steuergeraete_Pruefergebnisse]"]["F187_CMC3_Get"]

#Time/Date
timeStamp = Master["[Prueflauf]"]["Startzeit"]
startDate = timeStamp[:timeStamp.find("_")]
startDate = startDate[4:6] + "/" + startDate[6:] + "/" + startDate[:4]
startTime = timeStamp[timeStamp.find("_")+1:].replace("-",":")

#SW Flash Report
preFlashBCMeSW_Check = Master["[SW_Steuergeraete_Ablaufkriterien]"]["PreFlash_BMCSwVers_CheckOK?"]
PostFlashBCMeSW_Check = Master["[SW_Steuergeraete_Ablaufkriterien]"]["PostFlash_BMCSwVers_CheckOK?"]
dataSet_download_Check = Master["[SW_Steuergeraete_Ablaufkriterien]"]["DatensatzdownloadOK?"]
PostDSDL_BCMeSW_Check = Master["[SW_Steuergeraete_Ablaufkriterien]"]["PostDSDL_DSDLSwVers_CheckOK?"]
preFlashCMCSW_Check = Master["[SW_Steuergeraete_Ablaufkriterien]"]["PreFlash_CMCSwVers_CheckOK?"]
PostFlashCMCSW_Check = Master["[SW_Steuergeraete_Ablaufkriterien]"]["PostFlash_CMCSwVers_CheckOK?"]
SW_Check = Master["[SW_Steuergeraete_Ergebnis]"]["0"]

#Battery Status Setpoints Specs
CellTempMin = "{:.2f}".format(float(Master["[Batteriestatus_Sollwertvorgaben]"]["T_ZellMin_Set_Batteriestatus"]))
CellTempMax = "{:.2f}".format(float(Master["[Batteriestatus_Sollwertvorgaben]"]["T_ZellMax_Set_Batteriestatus"]))
CellTempRange = "{:.2f}".format(float(Master["[Batteriestatus_Sollwertvorgaben]"]["dT_ZellMax_Set_Batteriestatus"]))
CellVoltageMin = "{:.2f}".format(float(Master["[Batteriestatus_Sollwertvorgaben]"]["U_ZellMin_Set_Batteriestatus"]))
CellVoltageMax = "{:.2f}".format(float(Master["[Batteriestatus_Sollwertvorgaben]"]["U_ZellMax_Set_Batteriestatus"]))
CellVoltageRange = "{:.2f}".format(float(Master["[Batteriestatus_Sollwertvorgaben]"]["dU_ZellMax_Set_Batteriestatus"]))

#Temp/Voltage Check
CellTempCheck = Master["[Batteriestatus_Pruefergebnisse]"]["Temperaturwerte iO/niO?"]
CellTemp1 = "{:.2f}".format(float(Master["[Batteriestatus_Pruefergebnisse]"]["T_Zell_Get[n] [0]"]))
CellTemp2 = "{:.2f}".format(float(Master["[Batteriestatus_Pruefergebnisse]"]["T_Zell_Get[n] [1]"]))
CellTemp3 = "{:.2f}".format(float(Master["[Batteriestatus_Pruefergebnisse]"]["T_Zell_Get[n] [2]"]))
CellTemp4 = "{:.2f}".format(float(Master["[Batteriestatus_Pruefergebnisse]"]["T_Zell_Get[n] [3]"]))
CellTemp5 = "{:.2f}".format(float(Master["[Batteriestatus_Pruefergebnisse]"]["T_Zell_Get[n] [4]"]))
CellTemp6 = "{:.2f}".format(float(Master["[Batteriestatus_Pruefergebnisse]"]["T_Zell_Get[n] [5]"]))
CellTemp7 = "{:.2f}".format(float(Master["[Batteriestatus_Pruefergebnisse]"]["T_Zell_Get[n] [6]"]))
CellTemp8 = "{:.2f}".format(float(Master["[Batteriestatus_Pruefergebnisse]"]["T_Zell_Get[n] [7]"]))
CellTemp9 = "{:.2f}".format(float(Master["[Batteriestatus_Pruefergebnisse]"]["T_Zell_Get[n] [8]"]))
CellTemp10 = "{:.2f}".format(float(Master["[Batteriestatus_Pruefergebnisse]"]["T_Zell_Get[n] [9]"]))
CellTemp11 = "{:.2f}".format(float(Master["[Batteriestatus_Pruefergebnisse]"]["T_Zell_Get[n] [10]"]))
CellTemp12 = "{:.2f}".format(float(Master["[Batteriestatus_Pruefergebnisse]"]["T_Zell_Get[n] [11]"]))
CellTemp13 = "{:.2f}".format(float(Master["[Batteriestatus_Pruefergebnisse]"]["T_Zell_Get[n] [12]"]))
CellTemp14 = "{:.2f}".format(float(Master["[Batteriestatus_Pruefergebnisse]"]["T_Zell_Get[n] [13]"]))
CellTemp15 = "{:.2f}".format(float(Master["[Batteriestatus_Pruefergebnisse]"]["T_Zell_Get[n] [14]"]))
CellTemp16 = "{:.2f}".format(float(Master["[Batteriestatus_Pruefergebnisse]"]["T_Zell_Get[n] [15]"]))
CellTemp17 = "{:.2f}".format(float(Master["[Batteriestatus_Pruefergebnisse]"]["T_Zell_Get[n] [16]"]))
CellTemp18 = "{:.2f}".format(float(Master["[Batteriestatus_Pruefergebnisse]"]["T_Zell_Get[n] [17]"]))
CellTemp19 = "{:.2f}".format(float(Master["[Batteriestatus_Pruefergebnisse]"]["T_Zell_Get[n] [18]"]))
CellTemp20 = "{:.2f}".format(float(Master["[Batteriestatus_Pruefergebnisse]"]["T_Zell_Get[n] [19]"]))
CellTemp21 = "{:.2f}".format(float(Master["[Batteriestatus_Pruefergebnisse]"]["T_Zell_Get[n] [20]"]))
CellTemp22 = "{:.2f}".format(float(Master["[Batteriestatus_Pruefergebnisse]"]["T_Zell_Get[n] [21]"]))
CellTemp23 = "{:.2f}".format(float(Master["[Batteriestatus_Pruefergebnisse]"]["T_Zell_Get[n] [22]"]))
CellTemp24 = "{:.2f}".format(float(Master["[Batteriestatus_Pruefergebnisse]"]["T_Zell_Get[n] [23]"]))

#List of values for output file
myExcelData = [battery, partNum, batType, startDate, startTime, i_DTC_Count, initBCMe, initCMC1, initCMC2, initCMC3, initBL_BCMe, initBL_CMC1, initBL_CMC2, initBL_CMC3, batConfigPset, batVehiclePset, batTargetMarketPset, batConfigNamePset, batThermoPset, batNameVehiclePset, batNameTargetMarket, HW_NumCMC1, HW_NumCMC2, HW_NumCMC3,serial_CMC1,serial_CMC2,serial_CMC3,HW_NumBMCe,HW_serial_BMCe,SW_serial_BMCe,HW_serial_CMC1,HW_serial_CMC2,HW_serial_CMC3,SW_serial_CMC1,SW_serial_CMC2,SW_serial_CMC3,\
               preFlashBCMeSW_Check,PostFlashBCMeSW_Check,dataSet_download_Check,PostDSDL_BCMeSW_Check,preFlashCMCSW_Check,PostFlashCMCSW_Check,SW_Check,CellTempMin,CellTempMax,CellTempRange,CellVoltageMin,CellVoltageMax,CellVoltageRange,CellTempCheck,CellTemp1,CellTemp2,CellTemp3,CellTemp4,CellTemp5,CellTemp6,CellTemp7,CellTemp8,CellTemp9,CellTemp10,CellTemp11,CellTemp12,CellTemp13,CellTemp14,CellTemp15,CellTemp16,CellTemp17,CellTemp18,CellTemp19,CellTemp20,CellTemp21,CellTemp22,CellTemp23,CellTemp24]
myHeaders = ["Battery","Part Number", "Type","Date","Start Time","DTC Count","BCMe Begin SW","CMC1 Begin SW","CMC2 Begin SW","CMC3 Begin SW","BCMe Begin BL","CMC1 Begin BL","CMC2 Begin BL","CMC3 Begin BL","BatConfigPSet","VehiclePSet","TargetMarket","ConfigNamePSet","ThermoPSet","NameVehiclePSet","NameTargetMarket","CMC1 HW","CMC2 HW","CMC3 HW","CMC1 Serial","CMC2 Serial","CMC3 Serial","BCMe HW", "BCMe HW Serial", "BCMe SW Serial", "CMC1 HW Serial","CMC2 HW Serial","CMC3 HW Serial","CMC1 SW Serial","CMC2 SW Serial","CMC3 SW Serial",\
             "preFlashBCMeSW_Check","PostFlashBCMeSW_Check","dataSet_download_Check","PostDSDL_BCMeSW_Check","preFlashCMCSW_Check","PostFlashCMCSW_Check","SW_Check","CellTempMin","CellTempMax","CellTempRange","CellVoltageMin","CellVoltageMax","CellVoltageRange","CellTempCheck","CellTemp1","CellTemp2","CellTemp3","CellTemp4","CellTemp5","CellTemp6","CellTemp7","CellTemp8","CellTemp9","CellTemp10","CellTemp11","CellTemp12","CellTemp13","CellTemp14","CellTemp15","CellTemp16","CellTemp17","CellTemp18","CellTemp19","CellTemp20","CellTemp21","CellTemp22","CellTemp23","CellTemp24"]
count = 0

for header in myHeaders:
    count = count + 1
    sheet.cell(1, count).value = header

count = 0
for myVals in myExcelData:
    count = count + 1
    sheet.cell(2,count).value =myVals

#Format for time
sheet.cell(2,5).number_format = "h:mm:ss AM/PM"

for i in range(len(myExcelData)):
    print(myHeaders[i] + " : " +myExcelData[i])

wb.save(filename="hello_world.xlsx")