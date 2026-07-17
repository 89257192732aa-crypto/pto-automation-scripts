
num1 = int(input())   # вводим букву
num2 = input()
num3 = int(input())

if num2 == ('+'):
    print(num1 + num3)
elif num2 == ('-'):
    print(num1 - num3)
elif num2 == ('*'):
    print(num1 * num3)
elif num2 == ('/'):
    print(num1 / num3)
else:
    print('заново')
