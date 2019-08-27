import random


IDCharacters = "123456789ABCDEFGHJKLMNOPQRSTUVWXYZ"
def unique_id():
    id = ""
    for x in range(8):
        i = random.randint(0, len(IDCharacters)-1)
        id += IDCharacters[i]
    return id


if __name__ == '__main__':
    print(unique_id())
