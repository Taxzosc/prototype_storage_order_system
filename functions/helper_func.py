def convert_letter_to_index(column: str) -> int:
    """Returns index of alphabet letter, 0-25, a-z
    
    """
    alphabet = "abcdefghijklmnopqrstuvwxyz"
    return len(alphabet.split(column)[0])

def check_if_already_exists(order_numbers: set, order_number: int|float) -> bool:  # works, but what if order_numbers is big(10k+)
    """Returns True if order_number is in set

    takes a set of order_numbers, converts the set to a list and copies it.
    uses binary search to look for number.
    
    """
    copy_order_numbers = list(order_numbers.copy())
    copy_order_numbers.sort()
    low = 0
    high = len(copy_order_numbers) -1

    while low <= high:
        median = (low+high) // 2
        if copy_order_numbers[median] == order_number:
            return True
        elif copy_order_numbers[median] < order_number:
            low = median +1
        else:
            high = median -1
    return False
    # if order_number in copy_order_numbers:
    #     return True
    # return False
