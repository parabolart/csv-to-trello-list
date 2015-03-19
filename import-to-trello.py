import csv
import os
import sys

from trello import TrelloClient

class TrelloCSVImporter:
    def __init__(self, board_name, list_name, filename):
        self.board_name = board_name
        self.list_name = list_name
        self.filename = filename
        self.label_mapping = {
                'A': 'green',
                'B': 'yellow'
            }
        self.client = TrelloClient(api_key=os.environ['TRELLO_KEY'], api_secret=os.environ['TRELLO_SECRET'], token=os.environ['OAUTH_TOKEN'], token_secret=os.environ['OAUTH_SECRET'])
    
    def start(self):
        board = self.get_board(self.board_name)
        trello_list = self.get_list(board, self.list_name)
        to_import = self.process_input(self.filename)
        for item in to_import:
            self.create_trello_card(trello_list, item[0], item[1], item[2])
    
    def get_board(self, board_name):
        boards = self.client.list_boards()
        for board in boards:
            if board.name == board_name:
                return board
        raise "Board Not Found"
    
    def get_list(self, board, list_name):
        lists = board.all_lists()
        for list in lists:
            if list.name == list_name:
                return list
        raise "List Not Found"
    
    def process_input(self, filename):
        l = list()
        csvfile = open(filename, 'r')
        csv_reader = csv.reader(csvfile, delimiter='|')
        for row in csv_reader:
            if len(row) != 3:
                raise "Improper data format: " + row
            l.append(row)
        return l
    
    def create_trello_card(self, trello_list, name, description, label):
        new_card = trello_list.add_card(name, description)
        new_card._set_remote_attribute('labels', self.label_mapping[label])
        return 0

if __name__ == "__main__":
    if len(sys.argv) != 4:
        raise "Wrong arguments. Usage: python import-to-trello.py board_name list_name filename"
    importer = TrelloCSVImporter(sys.argv[1],sys.argv[2],sys.argv[3])
    importer.start()
    