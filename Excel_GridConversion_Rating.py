#Format text EDW file into grid for pasting into the excel model
#needs to have the | delimiter for defining columns
import csv


#need to enter the path of the text file below
textfileDirectory = ('c:\\ExcelGridConversion\\Detail_2015_SP.txt')

#open the CSV file, import into a list, each line is an index with additional indexes for each cell
opendetailFile = open(textfileDirectory)
readdataFile = csv.reader(opendetailFile, delimiter='|')
delimitedData = list(readdataFile)
opendetailFile.close()

#find out what indexes weight and zone are in and set those equal to variables. Regex?
delimitedData[0] = [x.lower() for x in delimitedData[0]] #makes the whole header lowercase -- not sure why, list comprehension

zoneWord = []

for i in range(len(delimitedData[0])):
	lowercaseData = delimitedData[0][i].lower()
	if 'zone' in lowercaseData: 
		zoneWord.append(lowercaseData)
if len(zoneWord)==0:
	print('No zone column found.')

#need to find the weight. Multiple ways it might show up (wgt, weight, rated weight, rated wgt)

weightWord = []

for i in range(len(delimitedData[0])):
	lowercaseData = delimitedData[0][i].lower()
	if 'wgt' or 'weight' in lowercaseData: #I tried to define a list here with the two words, and use the list to search in the other list, but it would not work
		if 'rated' in lowercaseData:
			weightWord.append(lowercaseData)
if len(weightWord)==0:
	print('No weight column found.')

#find out the index (columns) that the zone and weight are in 

#delimitedData[0] = [x.lower() for x in delimitedData[0]] #makes the whole header lowercase -- not sure why, list comprehension

weightIndex = delimitedData[0].index(str(weightWord[0]))
zoneIndex = delimitedData[0].index(str(zoneWord[0]))

#find the unique zones in the sheet

uniqueZones = set()

for i in range(1,len(delimitedData)):
	uniqueZones.add(delimitedData[i][zoneIndex])

uniqueZonesList = list(uniqueZones)
#build a function to search for integer zones, sort 1-8 and then append alphabetzones and sort a-z

def createZoneHeader(orderZones): #function to return a header of the unique zones in the data set
	numericZones = []
	numericZonesSorted = []
	alphabeticZones = []
	alphabeticZonesSorted =[]
	for i in range(len(orderZones)):
		try:
			if int(orderZones[i])>0:
				numericZones.append(orderZones[i])
		except ValueError:
			alphabeticZones.append(orderZones[i])
	numericZonesSorted = sorted(numericZones)
	alphabeticZonesSorted = sorted(alphabeticZones)
	combinedHeader = []
	combinedHeader = numericZonesSorted + alphabeticZonesSorted
	return combinedHeader

zoneHeader = createZoneHeader(uniqueZonesList) 

#find the uniques service codes for the data set

def createdServiceGroup(dataSet):
	svcWord=[]
	uniqueServices=set()
	for i in range(len(dataSet[0])):
		uniqueColumnName = dataSet[0][i]
		print(uniqueColumnName)
		if 'svc' or 'service' in str(uniqueColumnName):
			print(uniqueColumnName)
#			svcWord.append(dataSet[0][i])
		else:
			print('There is no service column.')
#	if len(svcWord)==0:
#		print('There is no service column.')
	svcIndex = dataSet[0].index(str(svcWord[0]))
	for i in range(1,len(dataSet)):
		uniqueServices.add(dataSet[i][svcIndex])
	uniqueServicesList = sorted(list(uniqueServices))
	return uniqueServicesList

allServices = createdServiceGroup(delimitedData)




zoneHeader.insert(0,'Weight(lbs)')

#create my excel grid with weight range up to 150. Each weight represents its own list

weightNumbers = list(range(0,151))
#weightNumbers.insert(0,zoneHeader) #need to create a list within the list. Right now it is a string file


numberGrid = [[] for _ in range(len(weightNumbers))] #creating a format of lists of lists must be a more efficient way to do this

for i in range(len(weightNumbers)): #the for loop allows me to drop in my weights in the lists of lists. Essentially my left hand column
	numberGrid[i].append(weightNumbers[i])

numberGrid.insert(0,[])
numberGrid[0].extend(zoneHeader)
