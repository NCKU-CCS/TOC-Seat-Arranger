import json
import argparse
from math import ceil
from random import shuffle

import xlwings as xw


START_ROW = 5
AVAILABLE_COLS_IN_EACH_ROOM = {
    '4261': ['B', 'D', 'G', 'J', 'M'],
    '4264': ['B', 'D', 'F', 'H', 'J', 'L', 'N'],
    '4263': ['B', 'D', 'F', 'H', 'J', 'L', 'N']
}


def arrange_seat(input_filename, output_filename):
    with open(input_filename) as input_file:
        students = json.load(input_file)
        shuffle(students)

    student_num = len(students)
    col_num = sum(len(cols) for cols in AVAILABLE_COLS_IN_EACH_ROOM.values())
    col_length = ceil(student_num / col_num)
    available_rows = range(START_ROW, START_ROW + (col_length-1)*2 + 1, 2)

    xb = xw.Book(output_filename)
    counter = 0

    try:
        for room, cols in AVAILABLE_COLS_IN_EACH_ROOM.items():
            xw.Sheet(room).activate()
            print(room)
            for col in cols:
                for row in available_rows:
                    print('{}{}'.format(col, row), students[counter])
                    xw.Range('{}{}'.format(col, row)).value = students[counter]
                    counter += 1
    except IndexError:
        pass
    xb.save()
    xb.close()


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('input_file')
    parser.add_argument('-o', '--output_file', default='seats.xlsx')
    args = parser.parse_args()

    arrange_seat(args.input_file, args.output_file)


if __name__ == "__main__":
    main()
