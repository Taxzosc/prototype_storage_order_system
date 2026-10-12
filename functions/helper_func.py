def convert_letter_to_index(column: str) -> int:
    """Returns index of alphabet letter, 0-25, a-z
    
    """
    alphabet = "abcdefghijklmnopqrstuvwxyz"
    return len(alphabet.split(column)[0])
