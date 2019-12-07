import csv, json
import openpyxl
from collections import OrderedDict

ch = "yes"
l = []
li = []


with open("xlfile.json","w") as f:

    while ch != "no":
        fname = str(input("First Name: "))
        lname = str(input("Last Name: "))
        roll = str(input("roll: "))
        dob = str(input("dd//mm//yyyy : "))
        gen = str(input("Gender(M/F/O):"))
        phone = str(input("Phone: "))
        mail = str(input("Mail: "))
        
        data = {
            "First_Name" : fname,
            "Last_Name" : lname,
            "Roll_Number" : roll,
            "DOB" : dob,
            "Gender" : gen,
            "Phone" : phone,
            "Mail" : mail
        }
        
        l.append(data)
        ch = str(input("Do you want to add more?\n(yes/no) : "))

    json.dump(l, f, indent=4)


# Opening json in read mode to obtain data from it
with open("xlfile.json","r") as fr:
    data = json.load(fr)

# Opening csv file in write mode to store values in it
f = open("xlfile.csv","w")

fields = ["First_Name","Last_Name","Roll_Number","DOB","Gender","Phone","Mail"]
writ = csv.DictWriter(f, fieldnames=fields)
writ.writeheader()
for d in data:
    writ.writerow(d)

f.close()


# Opening csv in read mode
fc = open("./xlfile.csv","r", newline="")

read = csv.DictReader(fc)
for i in read:
    li.append(i)


# Writing xl file
wb = openpyxl.Workbook()
ws = wb.active
ws.append(fields)

for row in li:
    val = OrderedDict(row)
    values = list(val.values())
    ws.append(values)
    
wb.save("student.xlsx")

print("Your details has been recorded.")