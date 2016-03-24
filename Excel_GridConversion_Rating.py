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

delimitedData[0] = [x.lower() for x in delimitedData[0]] #makes the whole header lowercase -- not sure why, list comprehension

weightIndex = delimitedData[0].index(str(weightWord[0]))
zoneIndex = delimitedData[0].index(str(zoneWord[0]))

#create my excel grid with weight range up to 150. Each weight represents its own list

weightNumbers = list(range(0,151))
weightNumbers.insert(0,'Weight(lbs)') #create a list from weight up to 150lbs

numberGrid = [[] for _ in range(len(weightNumbers))] #creating a format of lists of lists must be a more efficient way to do this

for i in range(len(weightNumbers)): #the for loop allows me to drop in my weights in the lists of lists. Essentially my left hand column
	numberGrid[i].append(weightNumbers[i])

#find the unique zones in the sheet

uniqueZones = set()

for i in range(len(delimitedData)):
	uniqueZones.add(delimitedData[i][19])

uniqueZonesList = list(uniqueZones)
#build a function to search for integer zones, sort 1-8 and then append alphavetzones and sort a-z

def createZoneHeader(orderZones): #all of this is stuck in the local scope now
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
	print(combinedHeader)

zoneHeader = createZoneHeader(uniqueZonesList) #this returns nothing the 2nd time I enter it

