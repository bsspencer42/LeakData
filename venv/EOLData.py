from openpyxl import Workbook
import os

#Function for getting data from file
def getEOL(myFile):
    input_file = open(myFile,'r')
    #Variable Setup
    #Master dictionary
    Master = {}

    #Initial Values for array
    battery, partNum, batType,testType, startDate, startTime, i_DTC_Count,SW_BMCe, SW_CMC1, SW_CMC2, SW_CMC3, SW_CMC4, BatConfigPSet, ThermoPSet, VehiclePSet, TargetMarket, NameTargetMarket, NameVehiclePSet, NameThermoPSet, NameBatConfig, CellTempMin, CellTempMax, CellTempRange, CellVoltageMin, CellVoltageMax, CellVoltageRange, CellTempCheck, CellTemp1, CellTemp2, CellTemp3, CellTemp4, CellTemp5, CellTemp6, CellTemp7, CellTemp8, CellTemp9, CellTemp10, CellTemp11, CellTemp12, CellTemp13, CellTemp14, CellTemp15, CellTemp16, CellTemp17, CellTemp18, CellTemp19, CellTemp20, CellTemp21, CellTemp22, CellTemp23, CellTemp24, CellTempRange, CellVolt1, CellVolt2, CellVolt3, CellVolt4, CellVolt5, CellVolt6, CellVolt7, CellVolt8, CellVolt9, CellVolt10, CellVolt11, CellVolt12, CellVolt13, CellVolt14, CellVolt15, CellVolt16, CellVolt17, CellVolt18, CellVolt19, CellVolt20, CellVolt21, CellVolt22, CellVolt23, CellVolt24, CellVolt25, CellVolt26, CellVolt27, CellVolt28, CellVolt29, CellVolt30, CellVolt31, CellVolt32, CellVolt33, CellVolt34, CellVolt35, CellVolt36, CellVolt37, CellVolt38, CellVolt39, CellVolt40, CellVolt41, CellVolt42, CellVolt43, CellVolt44, CellVolt45, CellVolt46, CellVolt47, CellVolt48, CellVolt49, CellVolt50, CellVolt51, CellVolt52, CellVolt53, CellVolt54, CellVolt55, CellVolt56, CellVolt57, CellVolt58, CellVolt59, CellVolt60, CellVolt61, CellVolt62, CellVolt63, CellVolt64, CellVolt65, CellVolt66, CellVolt67, CellVolt68, CellVolt69, CellVolt70, CellVolt71, CellVolt72, CellVolt73, CellVolt74, CellVolt75, CellVolt76, CellVolt77, CellVolt78, CellVolt79, CellVolt80, CellVolt81, CellVolt82, CellVolt83, CellVolt84, CellVolt85, CellVolt86, CellVolt87, CellVolt88, CellVolt89, CellVolt90, CellVolt91, CellVolt92, CellVolt93, CellVolt94, CellVolt95, CellVolt96, CellVoltRange, CellVoltageMinActual, CellVoltageMaxActual, CellTempMinActual, CellVoltageMaxActual, TempVoltageResult, FinalResult = "", "", "", "", "","", "", "", "", "", "", "", "", "", "","","", "", "", "", "", "", "", "", "", "","", "", "", "", "", "", "", "", "", "","", "", "", "", "", "", "", "", "", "","", "", "", "", "", "", "", "", "", "","", "", "", "", "", "", "", "", "", "","", "", "", "", "", "", "", "", "", "","", "", "", "", "", "", "", "", "", "","", "", "", "", "", "", "", "", "", "","", "", "", "", "", "", "", "", "", "","", "", "", "", "", "", "", "", "", "","", "", "", "", "", "", "", "", "", "","", "", "", "", "", "", "", "", "", "","", "", "", "", "", "", "", "", "", "","", "","","","","","","",""

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

   #Check to see if valid EOL file
    if "[Isolationswiderstand_Allgemeine Informationen]" not in Master.keys():
        return False

    try:
        #General Data
        battery = Master["[Pruefling]"].get("Seriennummer")[7:15]
        partNum = Master["[Pruefling]"].get("Teilenummer")
        batType = Master["[Pruefling]"].get("Batteriesystemtyp")
        testType = Master["[Pruefstand]"].get("Pruefstandstyp")

        #PreCheck
        i_DTC_Count = Master["[Isolationswiderstand_Allgemeine Informationen]"].get("uPruefungszaehlerIst")

        #Time/Date
        timeStamp = Master["[Prueflauf]"].get("Startzeit")
        startDate = timeStamp[:timeStamp.find("_")]
        startDate = startDate[4:6] + "/" + startDate[6:] + "/" + startDate[:4]
        startTime = timeStamp[timeStamp.find("_")+1:].replace("-",":")

        #Battery Status Setpoints Specs
        CellTempMin = "{:.2f}".format(float(Master.get("[Abschluss_EOL_Sollwertvorgaben]").get("T_ZellMin_Set_AbschlEOL")))
        CellTempMax = "{:.2f}".format(float(Master.get("[Abschluss_EOL_Sollwertvorgaben]").get("T_ZellMax_Set_AbschlEOL")))
        CellTempRange = "{:.2f}".format(float(Master.get("[Abschluss_EOL_Sollwertvorgaben]").get("dT_ZellMax_Set_AbschlEOL")))
        CellVoltageMin = "{:.2f}".format(float(Master["[Abschluss_EOL_Sollwertvorgaben]"].get("U_ZellMin_Set_AbschlEOL")))
        CellVoltageMax = "{:.2f}".format(float(Master["[Abschluss_EOL_Sollwertvorgaben]"].get("U_ZellMax_Set_AbschlEOL")))
        CellVoltageRange = "{:.2f}".format(float(Master["[Abschluss_EOL_Sollwertvorgaben]"].get("dU_ZellMax_Set_AbschlEOL")))
    except:
        pass

    try:
        # Temp Check
        CellTempCheck = Master["[Abschluss_EOL_Pruefergebnisse]"].get("Temperaturwerte IO?")
        CellTemp1 = "{:.2f}".format(float(Master["[Abschluss_EOL_Pruefergebnisse]"].get("T_Zell_Get[n] [0]")))
        CellTemp2 = "{:.2f}".format(float(Master["[Abschluss_EOL_Pruefergebnisse]"].get("T_Zell_Get[n] [1]")))
        CellTemp3 = "{:.2f}".format(float(Master["[Abschluss_EOL_Pruefergebnisse]"].get("T_Zell_Get[n] [2]")))
        CellTemp4 = "{:.2f}".format(float(Master["[Abschluss_EOL_Pruefergebnisse]"].get("T_Zell_Get[n] [3]")))
        CellTemp5 = "{:.2f}".format(float(Master["[Abschluss_EOL_Pruefergebnisse]"].get("T_Zell_Get[n] [4]")))
        CellTemp6 = "{:.2f}".format(float(Master["[Abschluss_EOL_Pruefergebnisse]"].get("T_Zell_Get[n] [5]")))
        CellTemp7 = "{:.2f}".format(float(Master["[Abschluss_EOL_Pruefergebnisse]"].get("T_Zell_Get[n] [6]")))
        CellTemp8 = "{:.2f}".format(float(Master["[Abschluss_EOL_Pruefergebnisse]"].get("T_Zell_Get[n] [7]")))
        CellTemp9 = "{:.2f}".format(float(Master["[Abschluss_EOL_Pruefergebnisse]"].get("T_Zell_Get[n] [8]")))
        CellTemp10 = "{:.2f}".format(float(Master["[Abschluss_EOL_Pruefergebnisse]"].get("T_Zell_Get[n] [9]")))
        CellTemp11 = "{:.2f}".format(float(Master["[Abschluss_EOL_Pruefergebnisse]"].get("T_Zell_Get[n] [10]")))
        CellTemp12 = "{:.2f}".format(float(Master["[Abschluss_EOL_Pruefergebnisse]"].get("T_Zell_Get[n] [11]")))
        CellTemp13 = "{:.2f}".format(float(Master["[Abschluss_EOL_Pruefergebnisse]"].get("T_Zell_Get[n] [12]")))
        CellTemp14 = "{:.2f}".format(float(Master["[Abschluss_EOL_Pruefergebnisse]"].get("T_Zell_Get[n] [13]")))
        CellTemp15 = "{:.2f}".format(float(Master["[Abschluss_EOL_Pruefergebnisse]"].get("T_Zell_Get[n] [14]")))
        CellTemp16 = "{:.2f}".format(float(Master["[Abschluss_EOL_Pruefergebnisse]"].get("T_Zell_Get[n] [15]")))
        CellTemp17 = "{:.2f}".format(float(Master["[Abschluss_EOL_Pruefergebnisse]"].get("T_Zell_Get[n] [16]")))
        CellTemp18 = "{:.2f}".format(float(Master["[Abschluss_EOL_Pruefergebnisse]"].get("T_Zell_Get[n] [17]")))
        CellTemp19 = "{:.2f}".format(float(Master["[Abschluss_EOL_Pruefergebnisse]"].get("T_Zell_Get[n] [18]")))
        CellTemp20 = "{:.2f}".format(float(Master["[Abschluss_EOL_Pruefergebnisse]"].get("T_Zell_Get[n] [19]")))
        CellTemp21 = "{:.2f}".format(float(Master["[Abschluss_EOL_Pruefergebnisse]"].get("T_Zell_Get[n] [20]")))
        CellTemp22 = "{:.2f}".format(float(Master["[Abschluss_EOL_Pruefergebnisse]"].get("T_Zell_Get[n] [21]")))
        CellTemp23 = "{:.2f}".format(float(Master["[Abschluss_EOL_Pruefergebnisse]"].get("T_Zell_Get[n] [22]")))
        CellTemp24 = "{:.2f}".format(float(Master["[Abschluss_EOL_Pruefergebnisse]"].get("T_Zell_Get[n] [23]")))
        CellTempRange = "{:.2f}".format(float(Master["[Abschluss_EOL_Pruefergebnisse]"].get("dT_Zell_Calc")))

    except:
        pass

    try:
        # Voltage Check
        CellVoltCheck = Master["[Abschluss_EOL_Pruefergebnisse]"].get("Spannungswerte IO?")
        CellVolt1 = "{:.3f}".format(float(Master["[Abschluss_EOL_Pruefergebnisse]"].get("U_Zell_Get[n] [0]")))
        CellVolt2 = "{:.3f}".format(float(Master["[Abschluss_EOL_Pruefergebnisse]"].get("U_Zell_Get[n] [1]")))
        CellVolt3 = "{:.3f}".format(float(Master["[Abschluss_EOL_Pruefergebnisse]"].get("U_Zell_Get[n] [2]")))
        CellVolt4 = "{:.3f}".format(float(Master["[Abschluss_EOL_Pruefergebnisse]"].get("U_Zell_Get[n] [3]")))
        CellVolt5 = "{:.3f}".format(float(Master["[Abschluss_EOL_Pruefergebnisse]"].get("U_Zell_Get[n] [4]")))
        CellVolt6 = "{:.3f}".format(float(Master["[Abschluss_EOL_Pruefergebnisse]"].get("U_Zell_Get[n] [5]")))
        CellVolt7 = "{:.3f}".format(float(Master["[Abschluss_EOL_Pruefergebnisse]"].get("U_Zell_Get[n] [6]")))
        CellVolt8 = "{:.3f}".format(float(Master["[Abschluss_EOL_Pruefergebnisse]"].get("U_Zell_Get[n] [7]")))
        CellVolt9 = "{:.3f}".format(float(Master["[Abschluss_EOL_Pruefergebnisse]"].get("U_Zell_Get[n] [8]")))
        CellVolt10 = "{:.3f}".format(float(Master["[Abschluss_EOL_Pruefergebnisse]"].get("U_Zell_Get[n] [9]")))
        CellVolt11 = "{:.3f}".format(float(Master["[Abschluss_EOL_Pruefergebnisse]"].get("U_Zell_Get[n] [10]")))
        CellVolt12 = "{:.3f}".format(float(Master["[Abschluss_EOL_Pruefergebnisse]"].get("U_Zell_Get[n] [11]")))
        CellVolt13 = "{:.3f}".format(float(Master["[Abschluss_EOL_Pruefergebnisse]"].get("U_Zell_Get[n] [12]")))
        CellVolt14 = "{:.3f}".format(float(Master["[Abschluss_EOL_Pruefergebnisse]"].get("U_Zell_Get[n] [13]")))
        CellVolt15 = "{:.3f}".format(float(Master["[Abschluss_EOL_Pruefergebnisse]"].get("U_Zell_Get[n] [14]")))
        CellVolt16 = "{:.3f}".format(float(Master["[Abschluss_EOL_Pruefergebnisse]"].get("U_Zell_Get[n] [15]")))
        CellVolt17 = "{:.3f}".format(float(Master["[Abschluss_EOL_Pruefergebnisse]"].get("U_Zell_Get[n] [16]")))
        CellVolt18 = "{:.3f}".format(float(Master["[Abschluss_EOL_Pruefergebnisse]"].get("U_Zell_Get[n] [17]")))
        CellVolt19 = "{:.3f}".format(float(Master["[Abschluss_EOL_Pruefergebnisse]"].get("U_Zell_Get[n] [18]")))
        CellVolt20 = "{:.3f}".format(float(Master["[Abschluss_EOL_Pruefergebnisse]"].get("U_Zell_Get[n] [19]")))
        CellVolt21 = "{:.3f}".format(float(Master["[Abschluss_EOL_Pruefergebnisse]"].get("U_Zell_Get[n] [20]")))
        CellVolt22 = "{:.3f}".format(float(Master["[Abschluss_EOL_Pruefergebnisse]"].get("U_Zell_Get[n] [21]")))
        CellVolt23 = "{:.3f}".format(float(Master["[Abschluss_EOL_Pruefergebnisse]"].get("U_Zell_Get[n] [22]")))
        CellVolt24 = "{:.3f}".format(float(Master["[Abschluss_EOL_Pruefergebnisse]"].get("U_Zell_Get[n] [23]")))
        CellVolt25 = "{:.3f}".format(float(Master["[Abschluss_EOL_Pruefergebnisse]"].get("U_Zell_Get[n] [24]")))
        CellVolt26 = "{:.3f}".format(float(Master["[Abschluss_EOL_Pruefergebnisse]"].get("U_Zell_Get[n] [25]")))
        CellVolt27 = "{:.3f}".format(float(Master["[Abschluss_EOL_Pruefergebnisse]"].get("U_Zell_Get[n] [26]")))
        CellVolt28 = "{:.3f}".format(float(Master["[Abschluss_EOL_Pruefergebnisse]"].get("U_Zell_Get[n] [27]")))
        CellVolt29 = "{:.3f}".format(float(Master["[Abschluss_EOL_Pruefergebnisse]"].get("U_Zell_Get[n] [28]")))
        CellVolt30 = "{:.3f}".format(float(Master["[Abschluss_EOL_Pruefergebnisse]"].get("U_Zell_Get[n] [29]")))
        CellVolt31 = "{:.3f}".format(float(Master["[Abschluss_EOL_Pruefergebnisse]"].get("U_Zell_Get[n] [30]")))
        CellVolt32 = "{:.3f}".format(float(Master["[Abschluss_EOL_Pruefergebnisse]"].get("U_Zell_Get[n] [31]")))
        CellVolt33 = "{:.3f}".format(float(Master["[Abschluss_EOL_Pruefergebnisse]"].get("U_Zell_Get[n] [32]")))
        CellVolt34 = "{:.3f}".format(float(Master["[Abschluss_EOL_Pruefergebnisse]"].get("U_Zell_Get[n] [33]")))
        CellVolt35 = "{:.3f}".format(float(Master["[Abschluss_EOL_Pruefergebnisse]"].get("U_Zell_Get[n] [34]")))
        CellVolt36 = "{:.3f}".format(float(Master["[Abschluss_EOL_Pruefergebnisse]"].get("U_Zell_Get[n] [35]")))
        CellVolt37 = "{:.3f}".format(float(Master["[Abschluss_EOL_Pruefergebnisse]"].get("U_Zell_Get[n] [36]")))
        CellVolt38 = "{:.3f}".format(float(Master["[Abschluss_EOL_Pruefergebnisse]"].get("U_Zell_Get[n] [37]")))
        CellVolt39 = "{:.3f}".format(float(Master["[Abschluss_EOL_Pruefergebnisse]"].get("U_Zell_Get[n] [38]")))
        CellVolt40 = "{:.3f}".format(float(Master["[Abschluss_EOL_Pruefergebnisse]"].get("U_Zell_Get[n] [39]")))
        CellVolt41 = "{:.3f}".format(float(Master["[Abschluss_EOL_Pruefergebnisse]"].get("U_Zell_Get[n] [40]")))
        CellVolt42 = "{:.3f}".format(float(Master["[Abschluss_EOL_Pruefergebnisse]"].get("U_Zell_Get[n] [41]")))
        CellVolt43 = "{:.3f}".format(float(Master["[Abschluss_EOL_Pruefergebnisse]"].get("U_Zell_Get[n] [42]")))
        CellVolt44 = "{:.3f}".format(float(Master["[Abschluss_EOL_Pruefergebnisse]"].get("U_Zell_Get[n] [43]")))
        CellVolt45 = "{:.3f}".format(float(Master["[Abschluss_EOL_Pruefergebnisse]"].get("U_Zell_Get[n] [44]")))
        CellVolt46 = "{:.3f}".format(float(Master["[Abschluss_EOL_Pruefergebnisse]"].get("U_Zell_Get[n] [45]")))
        CellVolt47 = "{:.3f}".format(float(Master["[Abschluss_EOL_Pruefergebnisse]"].get("U_Zell_Get[n] [46]")))
        CellVolt48 = "{:.3f}".format(float(Master["[Abschluss_EOL_Pruefergebnisse]"].get("U_Zell_Get[n] [47]")))
        CellVolt49 = "{:.3f}".format(float(Master["[Abschluss_EOL_Pruefergebnisse]"].get("U_Zell_Get[n] [48]")))
        CellVolt50 = "{:.3f}".format(float(Master["[Abschluss_EOL_Pruefergebnisse]"].get("U_Zell_Get[n] [49]")))
        CellVolt51 = "{:.3f}".format(float(Master["[Abschluss_EOL_Pruefergebnisse]"].get("U_Zell_Get[n] [50]")))
        CellVolt52 = "{:.3f}".format(float(Master["[Abschluss_EOL_Pruefergebnisse]"].get("U_Zell_Get[n] [51]")))
        CellVolt53 = "{:.3f}".format(float(Master["[Abschluss_EOL_Pruefergebnisse]"].get("U_Zell_Get[n] [52]")))
        CellVolt54 = "{:.3f}".format(float(Master["[Abschluss_EOL_Pruefergebnisse]"].get("U_Zell_Get[n] [53]")))
        CellVolt55 = "{:.3f}".format(float(Master["[Abschluss_EOL_Pruefergebnisse]"].get("U_Zell_Get[n] [54]")))
        CellVolt56 = "{:.3f}".format(float(Master["[Abschluss_EOL_Pruefergebnisse]"].get("U_Zell_Get[n] [55]")))
        CellVolt57 = "{:.3f}".format(float(Master["[Abschluss_EOL_Pruefergebnisse]"].get("U_Zell_Get[n] [56]")))
        CellVolt58 = "{:.3f}".format(float(Master["[Abschluss_EOL_Pruefergebnisse]"].get("U_Zell_Get[n] [57]")))
        CellVolt59 = "{:.3f}".format(float(Master["[Abschluss_EOL_Pruefergebnisse]"].get("U_Zell_Get[n] [58]")))
        CellVolt60 = "{:.3f}".format(float(Master["[Abschluss_EOL_Pruefergebnisse]"].get("U_Zell_Get[n] [59]")))
        CellVolt61 = "{:.3f}".format(float(Master["[Abschluss_EOL_Pruefergebnisse]"].get("U_Zell_Get[n] [60]")))
        CellVolt62 = "{:.3f}".format(float(Master["[Abschluss_EOL_Pruefergebnisse]"].get("U_Zell_Get[n] [61]")))
        CellVolt63 = "{:.3f}".format(float(Master["[Abschluss_EOL_Pruefergebnisse]"].get("U_Zell_Get[n] [62]")))
        CellVolt64 = "{:.3f}".format(float(Master["[Abschluss_EOL_Pruefergebnisse]"].get("U_Zell_Get[n] [63]")))
        CellVolt65 = "{:.3f}".format(float(Master["[Abschluss_EOL_Pruefergebnisse]"].get("U_Zell_Get[n] [64]")))
        CellVolt66 = "{:.3f}".format(float(Master["[Abschluss_EOL_Pruefergebnisse]"].get("U_Zell_Get[n] [65]")))
        CellVolt67 = "{:.3f}".format(float(Master["[Abschluss_EOL_Pruefergebnisse]"].get("U_Zell_Get[n] [66]")))
        CellVolt68 = "{:.3f}".format(float(Master["[Abschluss_EOL_Pruefergebnisse]"].get("U_Zell_Get[n] [67]")))
        CellVolt69 = "{:.3f}".format(float(Master["[Abschluss_EOL_Pruefergebnisse]"].get("U_Zell_Get[n] [68]")))
        CellVolt70 = "{:.3f}".format(float(Master["[Abschluss_EOL_Pruefergebnisse]"].get("U_Zell_Get[n] [69]")))
        CellVolt71 = "{:.3f}".format(float(Master["[Abschluss_EOL_Pruefergebnisse]"].get("U_Zell_Get[n] [70]")))
        CellVolt72 = "{:.3f}".format(float(Master["[Abschluss_EOL_Pruefergebnisse]"].get("U_Zell_Get[n] [71]")))
        CellVolt73 = "{:.3f}".format(float(Master["[Abschluss_EOL_Pruefergebnisse]"].get("U_Zell_Get[n] [72]")))
        CellVolt74 = "{:.3f}".format(float(Master["[Abschluss_EOL_Pruefergebnisse]"].get("U_Zell_Get[n] [73]")))
        CellVolt75 = "{:.3f}".format(float(Master["[Abschluss_EOL_Pruefergebnisse]"].get("U_Zell_Get[n] [74]")))
        CellVolt76 = "{:.3f}".format(float(Master["[Abschluss_EOL_Pruefergebnisse]"].get("U_Zell_Get[n] [75]")))
        CellVolt77 = "{:.3f}".format(float(Master["[Abschluss_EOL_Pruefergebnisse]"].get("U_Zell_Get[n] [76]")))
        CellVolt78 = "{:.3f}".format(float(Master["[Abschluss_EOL_Pruefergebnisse]"].get("U_Zell_Get[n] [77]")))
        CellVolt79 = "{:.3f}".format(float(Master["[Abschluss_EOL_Pruefergebnisse]"].get("U_Zell_Get[n] [78]")))
        CellVolt80 = "{:.3f}".format(float(Master["[Abschluss_EOL_Pruefergebnisse]"].get("U_Zell_Get[n] [79]")))
        CellVolt81 = "{:.3f}".format(float(Master["[Abschluss_EOL_Pruefergebnisse]"].get("U_Zell_Get[n] [80]")))
        CellVolt82 = "{:.3f}".format(float(Master["[Abschluss_EOL_Pruefergebnisse]"].get("U_Zell_Get[n] [81]")))
        CellVolt83 = "{:.3f}".format(float(Master["[Abschluss_EOL_Pruefergebnisse]"].get("U_Zell_Get[n] [82]")))
        CellVolt84 = "{:.3f}".format(float(Master["[Abschluss_EOL_Pruefergebnisse]"].get("U_Zell_Get[n] [83]")))
        CellVolt85 = "{:.3f}".format(float(Master["[Abschluss_EOL_Pruefergebnisse]"].get("U_Zell_Get[n] [84]")))
        CellVolt86 = "{:.3f}".format(float(Master["[Abschluss_EOL_Pruefergebnisse]"].get("U_Zell_Get[n] [85]")))
        CellVolt87 = "{:.3f}".format(float(Master["[Abschluss_EOL_Pruefergebnisse]"].get("U_Zell_Get[n] [86]")))
        CellVolt88 = "{:.3f}".format(float(Master["[Abschluss_EOL_Pruefergebnisse]"].get("U_Zell_Get[n] [87]")))
        CellVolt89 = "{:.3f}".format(float(Master["[Abschluss_EOL_Pruefergebnisse]"].get("U_Zell_Get[n] [88]")))
        CellVolt90 = "{:.3f}".format(float(Master["[Abschluss_EOL_Pruefergebnisse]"].get("U_Zell_Get[n] [89]")))
        CellVolt91 = "{:.3f}".format(float(Master["[Abschluss_EOL_Pruefergebnisse]"].get("U_Zell_Get[n] [90]")))
        CellVolt92 = "{:.3f}".format(float(Master["[Abschluss_EOL_Pruefergebnisse]"].get("U_Zell_Get[n] [91]")))
        CellVolt93 = "{:.3f}".format(float(Master["[Abschluss_EOL_Pruefergebnisse]"].get("U_Zell_Get[n] [92]")))
        CellVolt94 = "{:.3f}".format(float(Master["[Abschluss_EOL_Pruefergebnisse]"].get("U_Zell_Get[n] [93]")))
        CellVolt95 = "{:.3f}".format(float(Master["[Abschluss_EOL_Pruefergebnisse]"].get("U_Zell_Get[n] [94]")))
        CellVolt96 = "{:.3f}".format(float(Master["[Abschluss_EOL_Pruefergebnisse]"].get("U_Zell_Get[n] [95]")))

        SW_BMCe = Master["[Abschluss_EOL_Pruefergebnisse]"].get("SW_BMCe_Get")
        SW_CMC1 = Master["[Abschluss_EOL_Pruefergebnisse]"].get("SW_CMC1_Get")
        SW_CMC2 = Master["[Abschluss_EOL_Pruefergebnisse]"].get("SW_CMC2_Get")
        SW_CMC3 = Master["[Abschluss_EOL_Pruefergebnisse]"].get("SW_CMC3_Get")
        SW_CMC4 = Master["[Abschluss_EOL_Pruefergebnisse]"].get("SW_CMC4_Get")
        BatConfigPSet = Master["[Abschluss_EOL_Pruefergebnisse]"].get("DatensatzVersion_BatteriekonfigAbschlEOL_Get")
        ThermoPSet = Master["[Abschluss_EOL_Pruefergebnisse]"].get("DatensatzVersion_ThermoAbschlEOL_Get")
        VehiclePSet = Master["[Abschluss_EOL_Pruefergebnisse]"].get("DatensatzVersion_FahrzeugAbschlEOL_Get")
        TargetMarket = Master["[Abschluss_EOL_Pruefergebnisse]"].get("DatensatzVersion_ZielmarktAbschlEOL_Get")
        NameTargetMarket = Master["[Abschluss_EOL_Pruefergebnisse]"].get("DatensatzName_ZielmarktAbschlEOL_Get")
        NameVehiclePSet = Master["[Abschluss_EOL_Pruefergebnisse]"].get("DatensatzName_FahrzeugAbschlEOL_Get")
        NameThermoPSet = Master["[Abschluss_EOL_Pruefergebnisse]"].get("DatensatzName_ThermoAbschlEOL_Get")
        NameBatConfig = Master["[Abschluss_EOL_Pruefergebnisse]"].get("DatensatzName_BatteriekonfigAbschlEOL_Get")
    except:
        pass

    try:
        # Cell Voltage and Temp Results
        CellVoltRange = "{:.1f}".format(float(Master["[Abschluss_EOL_Pruefergebnisse]"].get("dU_ZellMax_Calc")))
        CellVoltageMinActual = "{:.3f}".format(float(Master["[Abschluss_EOL_Pruefergebnisse]"].get("U_ZellMin_Get"))/1000)
        CellVoltageMaxActual = "{:.3f}".format(float(Master["[Abschluss_EOL_Pruefergebnisse]"].get("U_ZellMax_Get"))/1000)
        CellTempMinActual = "{:.3f}".format(float(Master["[Abschluss_EOL_Pruefergebnisse]"].get("T_ZellMin_Get")))
        CellTempMaxActual = "{:.3f}".format(float(Master["[Abschluss_EOL_Pruefergebnisse]"].get("T_ZellMax_Get")))
        TempVoltageResult = Master["[Abschluss_EOL_Ergebnis]"].get("0")
    except:
        pass

    #Final Endline Result
    FinalResult = Master["[Ergebnis]"].get("0")

    #List of values for output file
    myExcelData = [battery, partNum, batType,testType, startDate, startTime, i_DTC_Count, \
                   CellTempMin, CellTempMax, CellTempRange, SW_BMCe, SW_CMC1, SW_CMC2, SW_CMC3, SW_CMC4, BatConfigPSet, ThermoPSet, VehiclePSet, TargetMarket, NameTargetMarket, NameVehiclePSet, NameThermoPSet, NameBatConfig, CellVoltageMin, CellVoltageMax, CellVoltageRange, CellTempCheck, CellTemp1, CellTemp2, CellTemp3, CellTemp4, CellTemp5, CellTemp6, CellTemp7, CellTemp8, CellTemp9, CellTemp10, CellTemp11, CellTemp12, CellTemp13, CellTemp14, CellTemp15, CellTemp16, CellTemp17, CellTemp18, CellTemp19, CellTemp20, CellTemp21, CellTemp22, CellTemp23, CellTemp24, \
                   CellTempRange, CellVolt1, CellVolt2, CellVolt3, CellVolt4, CellVolt5, CellVolt6, CellVolt7, CellVolt8, CellVolt9, CellVolt10, CellVolt11, CellVolt12, CellVolt13, CellVolt14, CellVolt15, CellVolt16, CellVolt17, CellVolt18, CellVolt19, CellVolt20, CellVolt21, CellVolt22, CellVolt23, CellVolt24, CellVolt25, CellVolt26, CellVolt27, CellVolt28, CellVolt29, CellVolt30, CellVolt31, CellVolt32, CellVolt33, CellVolt34, CellVolt35, CellVolt36, CellVolt37, CellVolt38, CellVolt39, CellVolt40, CellVolt41, CellVolt42, CellVolt43, CellVolt44, CellVolt45, CellVolt46, CellVolt47, CellVolt48, CellVolt49, CellVolt50, CellVolt51, CellVolt52, CellVolt53, CellVolt54, CellVolt55, CellVolt56, CellVolt57, CellVolt58, CellVolt59, CellVolt60, CellVolt61, CellVolt62, CellVolt63, CellVolt64, CellVolt65, CellVolt66, CellVolt67, CellVolt68, CellVolt69, CellVolt70, CellVolt71, CellVolt72, CellVolt73, CellVolt74, CellVolt75, CellVolt76, CellVolt77, CellVolt78, CellVolt79, CellVolt80, CellVolt81, CellVolt82, CellVolt83, CellVolt84, CellVolt85, CellVolt86, CellVolt87, CellVolt88, CellVolt89, CellVolt90, CellVolt91, CellVolt92, CellVolt93, CellVolt94, CellVolt95, CellVolt96, \
                   CellVoltRange,CellVoltageMinActual, CellVoltageMaxActual, CellTempMinActual, CellTempMaxActual, TempVoltageResult, FinalResult]
    return myExcelData

#Directory for leak test data
myEOLDir = r'\\vwoachsfile01\assembly\Departments\Assembly Launch\Battery Plant Status\Pilot Hall\IOL_EOL_Leak Docs\EOL'
#myEOLDir = r'\\vwoachsfile01\assembly\Departments\Assembly Launch\Battery Plant Status\Pilot Hall\IOL_EOL_Leak Docs\Test'


#Setup Excel file
wb = Workbook()
sheet = wb.active

#Setup Headers
myHeaders = ["Battery","Part Number", "Type","Test Type","Date","Start Time","DTC Count",\
             "CellTempMin","CellTempMax","CellTempRange","SW_BMCe", "SW_CMC1", "SW_CMC2", "SW_CMC3", "SW_CMC4", "BatConfigPSet", "ThermoPSet", "VehiclePSet", "TargetMarket", "NameTargetMarket", "NameVehiclePSet", "NameThermoPSet", "NameBatConfig","CellVoltageMin","CellVoltageMax","CellVoltageRange","CellTempCheck","CellTemp1","CellTemp2","CellTemp3","CellTemp4","CellTemp5","CellTemp6","CellTemp7","CellTemp8","CellTemp9","CellTemp10","CellTemp11","CellTemp12","CellTemp13","CellTemp14","CellTemp15","CellTemp16","CellTemp17","CellTemp18","CellTemp19","CellTemp20","CellTemp21","CellTemp22","CellTemp23","CellTemp24",\
             "CellTempRange","CellVolt1","CellVolt2","CellVolt3","CellVolt4","CellVolt5","CellVolt6","CellVolt7","CellVolt8","CellVolt9","CellVolt10","CellVolt11","CellVolt12","CellVolt13","CellVolt14","CellVolt15","CellVolt16","CellVolt17","CellVolt18","CellVolt19","CellVolt20","CellVolt21","CellVolt22","CellVolt23","CellVolt24","CellVolt25","CellVolt26","CellVolt27","CellVolt28","CellVolt29","CellVolt30","CellVolt31","CellVolt32","CellVolt33","CellVolt34","CellVolt35","CellVolt36","CellVolt37","CellVolt38","CellVolt39","CellVolt40","CellVolt41","CellVolt42","CellVolt43","CellVolt44","CellVolt45","CellVolt46","CellVolt47","CellVolt48","CellVolt49","CellVolt50","CellVolt51","CellVolt52","CellVolt53","CellVolt54","CellVolt55","CellVolt56","CellVolt57","CellVolt58","CellVolt59","CellVolt60","CellVolt61","CellVolt62","CellVolt63","CellVolt64","CellVolt65","CellVolt66","CellVolt67","CellVolt68","CellVolt69","CellVolt70","CellVolt71","CellVolt72","CellVolt73","CellVolt74","CellVolt75","CellVolt76","CellVolt77","CellVolt78","CellVolt79","CellVolt80","CellVolt81","CellVolt82","CellVolt83","CellVolt84","CellVolt85","CellVolt86","CellVolt87","CellVolt88","CellVolt89","CellVolt90","CellVolt91","CellVolt92","CellVolt93","CellVolt94","CellVolt95","CellVolt96",\
             "CellVoltRange","CellVoltageMinActual","CellVoltageMaxActual","CellTempMinActual","CellVoltageMaxActual","TempVoltageResult", "Voltage Set","Final Result"]

for i in range(len(myHeaders)):
    sheet.cell(1,i+1).value = myHeaders[i]

#Populate excel sheet w/ test data
lineNum = 2
for testData in os.listdir(myEOLDir):
    #Get next filename
    myFile = myEOLDir + "\\" + testData
    #Call file parser function
    myExcelData = getEOL(myFile)
    if myExcelData == False or myExcelData[0] == "":
        continue
    count = 0
    for myVals in myExcelData:
        count = count + 1
        sheet.cell(lineNum, count).value = myVals
    sheet.cell(lineNum,count+1).value = testData
    # Format for time
    sheet.cell(lineNum, 5).number_format = "h:mm:ss AM/PM"
    lineNum += 1
    wb.save(filename="EOL_data.xlsx")

#Adjust cell columns
for column_cells in sheet.columns:
    length = max(len(str(cell.value)) for cell in column_cells)
    sheet.column_dimensions[column_cells[0].column_letter].width = length

wb.save(filename="EOL_data.xlsx")