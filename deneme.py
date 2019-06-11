import string


dd = ["A"]
alphabet = string.ascii_uppercase
print(len(alphabet))


def convert(text):
    text = list(text)
    newText = []
    for x in text:
        if x in alphabet:
            newText.append(x)
            text.remove(x)
    for


convert("AAA")
