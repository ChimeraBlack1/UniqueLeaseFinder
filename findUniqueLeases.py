import math
import xlrd
import xlwt

loc = ("reportForPerry09062019.xlsm")
wb = xlrd.open_workbook(loc)

#open workbook
sheet = wb.sheet_by_index(0)

#write to workbook
workbook = xlwt.Workbook()
worksheet = workbook.add_sheet('Unique Leases')

LeaseList = []
leaseNumber = 1234
testMin = 0
testMax = len(LeaseList) - 1
target = 67
iteration = 0
NewWorkbookName = "unique leases yo.xls"

excelSize = 10

for x in range(2,excelSize):
  testInput = sheet.cell_value(x,3)
  worksheet.write(x, 0, testInput)

workbook.save(NewWorkbookName)

# if the LeaseList is empty, add the leaseNumber to it
if len(LeaseList) == 0:
  print(str(leaseNumber))
  LeaseList.append(leaseNumber)
elif len(LeaseList) > 1:
  # else if the len(LeaseList) > 1, run the binary search on that list for the leaseNumber
  LNIsFound = BinarySearch(LeaseList, testMin, testMax, leaseNumber)
  print("found: " + str(LNIsFound))
    # if leaseNumber is found, do not write to excel
    # if leaseNumber is NOT found, write to excel
else:
  # else add leaseNumber to the list
  print(str(x))
  LeaseList.append(leaseNumber)


def BinarySearch(theList, xmin, xmax, target):
    """
     This search method is the binary search method. It is a common algorithm for finding a target in a list where the list is already sorted.
    """
    found = False
    while xmin <= xmax and not found:
        arrVal = math.floor((xmin + xmax) / 2)
        if theList[arrVal] == target:
          found = True
        else:
          if theList[arrVal] < target:
            xmin = arrVal + 1
          else:
            xmax = arrVal - 1
    return found
    
