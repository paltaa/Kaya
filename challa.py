def challa(text):
    for i in text:
        if isupper(i):
            return text.split('i',1)[0]

print(challa('tula DE mono'))
