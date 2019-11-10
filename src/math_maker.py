import docx
import os
import random
import creator

random.seed()

def main():
    name1 = 'multiplication_medium'
    name2 = 'subtraction{}'
    for i in range(1):
        creator.create_simple_sheet(name=name1)
    """
    for i in range(5):
        creator.create_simple_sheet(name=name2.format(i), type='subtraction')
    """

if __name__ == '__main__':
    main()