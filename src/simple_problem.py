MAX_A = 99
MIN_A = 0
    
MAX_B = 99
MIN_B = 0

class SimpleProblem:
    def __init__(self, a, b, operation='addition'):
        self.a = a
        self.b = b
        self.operation = operation

    def toString(self):
        return 'a: {} b: {} operation: {}'.format(self.a, self.b, self.operation)