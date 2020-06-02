import register


if __name__ == "__main__":
    key = input("Enter key: ")
    reg = register.Register()
    print(reg.encrypted(key))
