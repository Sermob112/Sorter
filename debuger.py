from sorter import Sorter
from model import File
from collections import defaultdict
import os
# sort = Sorter("НВ600 документация ОССЗ")

# print(sort.count_files(True))
class Person:
    def __init__(self, name):
        self.name = name
    def say_hello(self):
        print("Hello, my name is", self.name)

person1 = Person("Alice")
person1.age = 25
print(person1.__dict__)