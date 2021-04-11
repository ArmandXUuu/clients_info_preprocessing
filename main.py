import csv
import re

class Client:
    def __init__(self):
        self.name = None
        self.number = 0
        self.communications = []
        self.gender = None
        self.suspecious_communications = []
        self.is_student = "N"
        self.buffer1 = None
        self.buffer2 = None

def process_gender(row):
    MR_pattern = re.compile(r'^.* MR .*$')
    MS_pattern = re.compile(r'^.* MS .*$')
    MRS_pattern = re.compile(r'^.* MRS .*$')
    if MR_pattern.match(row) != None:
        return "MR"
    if MS_pattern.match(row) != None:
        return "MS"
    if MRS_pattern.match(row) != None:
        return "MRS"
    return None

def read_from_csv():
    with open('input.csv','r', encoding="utf-8") as csvfile:
        reader = csv.reader(csvfile)
        strTmp = ""
        rows = [re.sub(' +', ' ', str.lstrip(strTmp.join(row))) for row in reader]
    return rows

def write_to_csv(clients):
    f = open('output.csv', 'w', encoding='utf-8')
    csv_writer = csv.writer(f)
    csv_writer.writerow(["number","gender","name","buffer1",'buffer2',"is_student","contact","suspecious_contact"])
    for client in clients:
        csv_writer.writerow([client.number, client.gender, client.name, client.buffer1, client.buffer2, client.is_student, client.communications, client.suspecious_communications])

def main():
    rows = read_from_csv()

    start_person = re.compile(r'^\d\d\d .*')
    start_person_Car = re.compile(r'^\d\d\d[A-Z] .*')
    start_OSI = re.compile(r'^OSI.*')
    pure_numbers = re.compile(r'^\d.*')
    six_reservation_code = re.compile(r'[A-Za-z0-9]{6}')
    is_student = re.compile(r'^.* SD .*$')
    
    row = None
    clients = []
    client_tmp = Client()

    for row in rows:
        row_tmp = str.split(row, " ")
        if start_person.match(row) != None or start_person_Car.match(row) != None: #Means a new person comes
            clients.append(client_tmp)
            client_tmp = Client()
            client_tmp.number = row_tmp[0]
            client_tmp.name = row_tmp[1]
            client_tmp.buffer1 = row_tmp[2]
            client_tmp.buffer2 = row_tmp[3]
            client_tmp.gender = process_gender(row)
            if is_student.match(row) != None:
                client_tmp.is_student = "Y"
        else :
            if start_OSI.match(row) != None:
                client_tmp.communications.append(row_tmp[2::])
            elif (pure_numbers.match(row) != None):
                client_tmp.suspecious_communications = row
            
    clients.append(client_tmp)
    write_to_csv(clients)
    print("Thank you for using, processing complete.")

if __name__ == '__main__':
    main()