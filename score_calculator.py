from openpyxl.reader.excel import load_workbook

class Item:
    def __init__(self, name, rank):
        self.name = name
        self.rank = rank

    def __repr__(self):
        return self.name + "':" + str(self.rank)

def calculate():
    girls_workbook = load_workbook(filename="poll/FILLES.xlsx")
    boys_workbook = load_workbook(filename="poll/GARCON.xlsx")

    girls = {}
    boys = {}
    attribution = {}

    for girl in girls_workbook:
        girls[girl.title] = []
        for boy in boys_workbook:
            score = 0
            for i in range(2, 26):
                if boy["B" + str(i)].value == girl["B" + str(i)].value:
                    score += 1
            girls[girl.title].append(Item(boy.title, score))

    for boy in boys_workbook:
        boys[boy.title] = {}
        attribution[boy.title] = None
        for girl in girls_workbook:
            score = 0
            for i in range(2, 26):
                if girl["B" + str(i)].value == boy["B" + str(i)].value:
                    score += 1
            boys[boy.title][girl.title] = score
    return {
        "boys": boys,
        "girls": girls,
        "attribution": attribution,
    }


