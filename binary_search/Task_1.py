def binary_search(lst, item):
    low = 0
    high = len(lst) - 1 # в переменных лоу и хай хранятся границы той части списка в которой выполняется поиск

    while low <= high: # пока эта часть не сократится до 1 элемента
        mid = (low + high) // 2  # Целочисленное деление - проверяем средний элемент
        guess = lst[mid]

        if guess == item:  # Элемент найден
            return mid
        elif guess < item:  # Ищем в правой части - много
            low = mid + 1
        else:  # Ищем в левой части - мало
            high = mid - 1

    return None  # Элемент не найден - значение не существует

# Определяем список и вызываем функцию
my_list = [1, 3, 5, 7, 9]

print(binary_search(my_list, 9))  # => 4 (индекс элемента 9)
print(binary_search(my_list, -1))  # => None (элемент не найден)
