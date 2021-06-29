myFile = r'\\vwoachsfile01\assembly\Departments\Assembly Launch\Battery Plant Status\Pilot Hall\IOL_EOL_Leak Docs\PVS\288 TX6PVS0340_2021_06_16_180429.ini'
input_file = open(myFile,'r')
line = input_file.readline()

#Find Software setpoint specs header
while input_file:
    line = input_file.readline().strip()
    if line == "[SW_Steuergeraete_Sollwertvorgaben]":
        break
#Set master dictionary
Master = {}

#Dictionary to collect current data
myDict = {}

#Loop to collect current data
while line != "":
    #String parsing
    line = input_file.readline()
    data = line[line.find('"')+1:].strip()
    data = data[:-1]
    #Create list of data/value pair
    myData = data.split(";")
    #Create list of lists
    if line.strip() == "":
        break
    myDict[myData[0]] = myData[1]

#Add to master dictionary
Master["PartInfo"] = myDict

print(Master["PartInfo"])

