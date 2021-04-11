import csv
import re

class Client:
    def __init__(self):
        self.name = None
        self.number = 0
        self.communications = []
        self.gender = None
        self.suspecious_communications = []

def read_from_csv():
    with open('input.csv','r', encoding="utf-8") as csvfile:
        reader = csv.reader(csvfile)
        strTmp = ""
        rows = [re.sub(' +', ' ', str.lstrip(strTmp.join(row))) for row in reader]
    return rows

def write_to_csv(clients):
    f = open('output.csv', 'w', encoding='utf-8')
    csv_writer = csv.writer(f)
    csv_writer.writerow(["number","name","gender","contact","suspecious_contact"])
    for client in clients:
        csv_writer.writerow([client.number, client.name, client.gender, client.communications, client.suspecious_communications])

def main():
    rows = read_from_csv()

    start_person = re.compile(r'^\d\d\d .*')
    start_person_Car = re.compile(r'^\d\d\d[A-Z] .*')
    start_OSI = re.compile(r'^OSI.*')
    pure_numbers = re.compile(r'^\d.*')
    
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
            client_tmp.gender = row_tmp[2]
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