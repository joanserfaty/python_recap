# 1 - Array 3D
array = [[[0 for col in range(8)]for col in range(3)]for col in range(5)]
print(array)

# 2 - Formating
subjects=["I", "You"]
verbs=["Play", "Love"]
objects =["Hockey","Football"]
sentence =[]
for sub in subjects:
    for ver in verbs:
        for obj in objects:
            sentence.append('{0} {1} {2}'.format(sub, ver, obj))

print(sentence)

#3 - Compress with Zlib
import zlib
s = 'hello world!hello world!hello world!hello world!'
t = zlib.compress(s)
print t
t = zlib.decompress(t)

#4 - Generate Random password
import string
import random

letters = string.ascii_letters
numbers = string.digits
len_pass = random.randint(5, 10)
rand = [random.choice(letters + numbers) for elem in range(len_pass)]
rand = ''.join(rand)

#5 - Return index number
lis = ['b','o','j','o','r','b']
       #Only First item
ind = lis.index('b')
       # All of them
ind_list = [elem for elem, id in enumerate(lis) if id == 'b']

#6 - Generators
def Generator(n):
    i=0
    while i < n:
        if i % 2 == 0:
            yield i
        i += 1

oven_list=[]
for elem in Generator(9):
    oven_list.append(elem)

#7 - Find All
import re

stri = 'Bonjour je m appel Joan'
r = re.findall('j', stri)

#8 - Fibonacci
def fibo(n):
    if n == 0: return 0
    elif n == 1: return 1
    else: return fibo(n-1)+fibo(n-2)

r = fibo(7)

#9 - Adress mail
import re

mail = 'john.serfaty@google.com'
path = '(\w+).(\w+)@(\w+).com'
r2 = re.match(path, mail)
prenom = r2.group(1)
nom = r2.group(2)
ad = r2.group(3)

#10 - Decorators
def check(func):
    def inside(num1, num2):
        if num2 == 0:
            print('Impossible with 0')
            return
        return func(num1, num2)
    return inside

@check
def div(a, b):
    return a/b

result = div(5, 0)

#11 - Map
my_l = [1,2,3,4,5,6,7,8,9,10]
my_func = (lambda x: x ** 2)
my_new_l = map(my_func, my_l)

#12 - Lambda
my_l = [1,2,3,4,5,6,7,8,9,10]
func = (lambda x: x % 2 == 0)
new_l = filter(func, my_l)

#13 - Tuples
tup = (1,2,3,4,5,6,7,8,9,10)
half = len(tup)//2
first_half = sum(tup[:half])
second_half = sum(tup[half:])

#14 - Dict
dict = {}
dict2 = {}
for elem in range(1,21):
    dict[elem] = elem**2
    dict2.update({elem: elem * elem})

print(dict)
print(dict2)

#15 - Bank Details
bank_details = [('D', 300), ('D', 300), ('W', 200), ('W', 100)]
credit = sum([cash for way, cash in bank_details if way == 'W'])
debit = sum([cash for way, cash in bank_details if way == 'D'])

#16 - List comprehension
my_str = '1,2,3,4,5,6,7,8,9'
my_str = my_str.split(',')
my_str = [elem for elem in my_str if int(elem) % 2 != 0]
my_str = ','.join(my_str)

#17 - Number of letters and digit
my_str = 'hello world! 123'
digit, alpha = 0, 0

for elem in my_str:
    if elem.isdigit():
        digit += 1
    elif elem.isalpha():
        alpha += 1


#18 - Even number
my_numb = [numb for numb in range(1000, 3001)]
final_list = []
for elem in my_numb:
    test_list = []
    for chara in str(elem):
        if int(chara) % 2 == 0:
            test_list.append(True)
        else:
            test_list.append(False)

    if False not in test_list:
        final_list.append(elem)


#19 - list of number divised by 5
def my_list(numb):
    my_list = [int(elem) for elem in numb if int(elem) % 5 == 0]
    return my_list

r = my_list([0100, 0011, 1010, 1001])
print(r)

#20 - string to list and remove duplicates
def convert_str(stri):
    my_list = stri.split(' ')
    my_list = set(my_list)
    return list(my_list)

stri = 'hello world and practice makes perfect and hello world again'
new_stri = convert_str(stri)
print(new_stri)

#21 - Add Elem to Dict
def dict(numb):
    dico = {}
    for elem in range(1, numb+1):
        #d[elem]=elem*elem
        dico.update({elem: elem**2})

    return dico
my_dict = dict(8)
print(my_dict)


#22 - Factorial
def factorial(numb):
    itera = 1
    for elem in range(numb, 0, -1):
        itera *= elem

    return itera

my_fac = factorial(7)
print(my_fac)

#23 - List to string
def number_d():
    my_numb = [str(num) for num in range(2000, 3000) if (num % 7 == 0) and not (num % 5 == 0)]
    my_str = ','.join(my_numb)
    return my_str


my_list = number_d()
print(my_list)


