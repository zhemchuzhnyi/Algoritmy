# def binary_search(list, item):
#     low = 0
#     high = len(list) - 1
#     while low <= high:
#         mid = (low + high) / 2
#         guess = list[mid]
#         if guess == item:
#             high = mid - 1
#         else:
#             low =mid + 1
#         return None
#
#     my_list = [1, 3, 5, 7, 9]
#
#     print (binary_search(my_list, 9)) # => 1 - нумерация начинается с 0 - вторая ячейка это 1
#     print (binary_search(my_list, -1)) # => None - в пайтоне означает ничего - признак того что элемент не найден


def binary_search(lst, item):
    low = 0
    high = len(lst) - 1

    while low <= high:
        mid = (low + high) // 2  # Целочисленное деление
        guess = lst[mid]

        if guess == item:  # Элемент найден
            return mid
        elif guess < item:  # Ищем в правой части
            low = mid + 1
        else:  # Ищем в левой части
            high = mid - 1

    return None  # Элемент не найден

# Определяем список и вызываем функцию
my_list = [1, 3, 5, 7, 9]

print(binary_search(my_list, 3))  # => 4 (индекс элемента 9)
print(binary_search(my_list, -1))  # => None (элемент не найден)
