from FileRead import FileRead

if __name__ == '__main__':
    filepath = r'./requirements.txt'
    fr = FileRead(filepath)
    print(fr.readtext())
    print(fr.getinfo())