import csv
import os
import sys
import settings

from trello import TrelloClient

hmh = {
    'leadsource': '5c78123201bc4c7950720e71',
    'client': '5c7812a444111f835dd0ca43',
    'poc': '5c7812b0a95d3333486e74c4',
    'contactnumber': '5c7812bb8ba143056f12ebae',
    'bidvalue': '5c7812e25fe9700eb0f72c03',
    'ptd': '5c781334b9fe9d59892888cf',
    'balance': '5c78133fd2b34b48ed231e7e',

    'SRT': '5c78123e2bc98859a68627f7',
    'JAD': '5c78123201bc4c7950720e72',
    'HMH': '5c7812415f83d8214607414c'
}

class TrelloCSVImporter:

    def __init__(self, board_name, list_name, filename):
        self.board_name = board_name
        self.list_name = list_name
        self.filename = filename
        self.label_mapping = {
                'A': 'green',
                'B': 'yellow'
            }
        self.client = TrelloClient(api_key=os.environ['TRELLO_KEY'], api_secret=os.environ['TRELLO_SECRET'], token=os.environ['OAUTH_TOKEN'])

    def start(self):
        board = self.get_board(self.board_name)
        trello_list = self.get_list(board, self.list_name)
        to_import = self.process_input(self.filename)
        for item in to_import:
            self.create_trello_card(trello_list, item[0], item[1], item[2], item[3], item[4], item[5], item[6], item[7])

    def get_board(self, board_name):
        boards = self.client.list_boards()
        for board in boards:
            if board.name == board_name:
                return board
        raise LookupError("Board Not Found")

    def get_list(self, board, list_name):
        lists = board.all_lists()
        for list in lists:
            if list.name == list_name:
                return list
        raise LookupError("List Not Found")

    def process_input(self, filename):
        l = list()
        csvfile = open(filename, 'r')
        csv_reader = csv.reader(csvfile, delimiter='|')
        for row in csv_reader:
            if len(row) != 8:
                raise StandardError("Improper data format: " + row)
            l.append(row)
        return l

    def create_trello_card(self, trello_list, name, leadsource, client, poc, contactnumber, bid, ptd, balance):
        new_card = trello_list.add_card(name, '') #description missing for now
        # new_card._set_remote_attribute('labels', self.label_mapping[label]) this is cuz they had other args instead of A or B
        new_card._set_remote_attribute('labels', 'green')
        #new_card._set_remote_attribute('customFieldItems', 'test')
        #leadsource = new_card.get_custom_field_by_name('POC')
        new_card._set_remote_attribute_custom_field(hmh['client'],client)
        new_card._set_remote_attribute_custom_field(hmh['poc'],poc)
        new_card._set_remote_attribute_custom_field(hmh['contactnumber'],contactnumber)
        #new_card._set_remote_attribute_custom_field(hmh['bid'],bid) number?
        new_card._set_remote_attribute_custom_field(hmh['ptd'],ptd)
        new_card._set_remote_attribute_custom_field(hmh['balance'],balance)
        if leadsource != '':
            new_card._set_remote_attribute_custom_field_list(hmh['leadsource'],hmh[leadsource])
        print(new_card.fetch_custom_fields_hmh())
        #print(new_card.get_custom_field_by_name('Lead Source'))
        return 0

if __name__ == "__main__":
    if len(sys.argv) != 4:
        raise EnvironmentError("Wrong arguments. Usage: python import-to-trello.py board_name list_name filename")
    importer = TrelloCSVImporter(sys.argv[1],sys.argv[2],sys.argv[3])
    importer.start()
