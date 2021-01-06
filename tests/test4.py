class Example:
    def __init__(self, a, b):
        self.a = a
        self.b = b

    def __call__(self, a, b):
        self.a = a
        self.b = b

e = Example(5, 6)
e(3,4)
print(e.a, e.b)