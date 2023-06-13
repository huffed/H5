def client_info():
    flag = True
    while flag:
        # client enters mobile number
        mobile_number = input("please enter your mobile number: ")
        try:
            int(mobile_number)
            if mobile_number[0] != "0" or len(mobile_number) != 11:
                print("invalid phone number")
                flag = True
            else:
                print("correct number")
        except ValueError:
            print("sorry that is not a valid mobile number")

        # client enters age
        age = input("please enter your age: ")
        try:
            int(age)
            if age < 0 or age > 100:
                print("that is not a valid age ")
                flag = True
            else:
                print("corect age")
        except ValueError:
            print("that is not a valid age ")


client_info()
