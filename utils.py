def find_max(number):
    try:
        maximum = number[0]
        for i in number:
            if i > maximum:
                maximum = i
        return maximum
    except TypeError:
        return number
