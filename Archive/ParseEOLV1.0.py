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
for (key,value) in Master.items():
    print(key, value)