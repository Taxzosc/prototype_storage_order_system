def convert_letter_to_index(column: str) -> int:
    alphabet = "abcdefghijklmnopqrstuvwxyz"
    return len(alphabet.split(column)[0])
