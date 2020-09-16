import docx
import os
import random
import creator

random.seed()

def main():
    name1 = 'multiplication_by7s2'
    name2 = 'multiplication_by8s3'
    
    creator.create_simple_sheet(name=name1, type='multiplication')

    
if __name__ == '__main__':
    main()